import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface FixtureWindow {
  __cyexcelHeadersReady?: boolean;
  __setHeaderHitMode?(mode: "handled" | "decline" | "throw"): void;
  __headerHitTestCalls?(): number;
  __headerActions?: Array<{
    readonly action: { readonly type: string };
    readonly column: number;
    readonly header: { readonly column: number; readonly title: string };
    readonly structureVersion: number;
  }>;
  __headerErrors?: Array<{ readonly eventName: string; readonly message: string }>;
  __headerDomHasPointerCapture?(): boolean;
  __sheet?: {
    readonly selection: {
      readonly kind: string;
      readonly active: { readonly row: number; readonly column: number };
      readonly range: {
        readonly startRow: number;
        readonly startColumn: number;
        readonly endRow: number;
        readonly endColumn: number;
      };
    };
    readonly scene: {
      getCellRect(row: number, column: number): { readonly width: number };
    };
    getCell(row: number, column: number): unknown;
    dispose(): void;
    insertColumns(start: number, columns: readonly { readonly id: string }[]): void;
    deleteColumns(start: number, count?: number): void;
    moveColumns(start: number, count: number, target: number): void;
  };
}

const GROUP_Y = 14;
const LEAF_Y = 42;

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/headers.html");
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelHeadersReady === true
  );
}

async function nextFrame(page: Page): Promise<void> {
  await page.waitForFunction(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
  );
}

function getCell(page: Page, row: number, column: number): Promise<unknown> {
  return page.evaluate(
    ({ row: r, column: c }) => (window as unknown as FixtureWindow).__sheet?.getCell(r, c),
    { row, column }
  );
}

function getHeaderHitTestCalls(page: Page): Promise<number> {
  return page.evaluate(() =>
    (window as unknown as FixtureWindow).__headerHitTestCalls?.() ?? 0
  );
}

function getSelection(page: Page) {
  return page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.selection);
}

function getColumnWidth(page: Page, column: number): Promise<number | undefined> {
  return page.evaluate(
    (targetColumn) =>
      (window as unknown as FixtureWindow).__sheet?.scene.getCellRect(0, targetColumn).width,
    column
  );
}

async function setHeaderHitMode(
  page: Page,
  mode: "handled" | "decline" | "throw"
): Promise<void> {
  await page.evaluate(
    (nextMode) => (window as unknown as FixtureWindow).__setHeaderHitMode?.(nextMode),
    mode
  );
}

// DomHeaderManager positions the <select>'s wrapper host div, not the
// <select> itself (mirrors DOMCellManager: the factory mounts its own
// content inside the container DomHeaderManager creates and positions).
function statusHeaderHostLeft(page: Page): Promise<string> {
  return page
    .locator("#app select")
    .evaluate((el: HTMLElement) => (el.parentElement as HTMLElement).style.left);
}

test("header canvas paints real, non-transparent pixels", async ({ page }) => {
  await gotoFixture(page);

  const hasPaintedPixel = await page.evaluate(() => {
    const canvas = document.querySelectorAll("#app canvas")[2] as HTMLCanvasElement; // header canvas
    const ctx = canvas.getContext("2d") as CanvasRenderingContext2D;
    const dpr = window.devicePixelRatio || 1;
    const { data } = ctx.getImageData(0, 0, Math.round(400 * dpr), Math.round(28 * dpr));
    for (let i = 3; i < data.length; i += 4) {
      if ((data[i] ?? 0) > 0) {
        return true;
      }
    }
    return false;
  });

  expect(hasPaintedPixel).toBe(true);
});

test("DOM header stays aligned with body content during horizontal scroll (退出条件 1)", async ({
  page
}) => {
  await gotoFixture(page);

  // Column x-order: sku(100, frozen) | name(140) | price(90) | qty(80) | status(100).
  // status's absolute left = 100+140+90+80 = 410; not frozen, so viewport-relative
  // left = 410 - scrollLeft.
  expect(await statusHeaderHostLeft(page)).toBe("410px");

  await page.locator("#app").evaluate((el) => {
    const scroller = el.firstElementChild as HTMLElement;
    scroller.scrollLeft = 50;
    scroller.dispatchEvent(new Event("scroll"));
  });
  await nextFrame(page);

  expect(await statusHeaderHostLeft(page)).toBe("360px");
});

test("a DOM header keeps native pointer ownership and does not change grid selection", async ({ page }) => {
  await gotoFixture(page);
  await page.locator("#app").evaluate((element) => {
    const scroller = element.firstElementChild as HTMLElement;
    scroller.scrollLeft = 100;
    scroller.dispatchEvent(new Event("scroll"));
  });
  await nextFrame(page);

  const select = page.locator("#app select");
  const box = await select.boundingBox();
  if (!box) throw new Error("missing DOM header select box");
  const before = await getSelection(page);

  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down();
  expect(
    await page.evaluate(() =>
      (window as unknown as FixtureWindow).__headerDomHasPointerCapture?.() ?? false
    )
  ).toBe(false);
  expect(await getSelection(page)).toEqual(before);
  await page.mouse.up();

  await select.selectOption({ label: "Archived" });
  await expect(select).toHaveValue("Archived");
  expect(await getSelection(page)).toEqual(before);
  expect(await getHeaderHitTestCalls(page)).toBe(0);
});

test("a native-owned closed-shadow header is terminal at its mounted host", async ({ page }) => {
  await gotoFixture(page);
  await page.locator("#app").evaluate((element) => {
    const scroller = element.firstElementChild as HTMLElement;
    scroller.scrollLeft = 190;
    scroller.dispatchEvent(new Event("scroll"));
  });
  await nextFrame(page);

  const before = await getSelection(page);
  await page.locator("#app .closed-native-header").click();

  expect(await getSelection(page)).toEqual(before);
  expect(await getHeaderHitTestCalls(page)).toBe(0);
});

test("frozen row/column stay pinned; editing a scrolled-into-view body cell still works (先写测试 #7)", async ({
  page
}) => {
  await gotoFixture(page);

  await page.locator("#app").evaluate((el) => {
    const scroller = el.firstElementChild as HTMLElement;
    // 408 = 17 * rowHeight(24): the content canvas's own local y-axis
    // reserves its first 24px (frozenRowCount 1 * rowHeight) for the
    // frozen row band, so scrollable content only starts at local y=24 —
    // row 18 is the first scrollable row whose band (y:[24,48)) begins
    // right there. See doc/architecture.md §6.1.
    scroller.scrollTop = 408;
    scroller.scrollLeft = 0;
    scroller.dispatchEvent(new Event("scroll"));
  });
  await nextFrame(page);

  const contentBox = await page.locator("#app canvas").first().boundingBox();
  if (!contentBox) {
    throw new Error("missing content canvas box");
  }

  // Double-click comfortably inside the scrollable band (past the 24px
  // frozen row band) — must open an editor positioned inside the viewport,
  // not off-screen at its absolute content position.
  await page.mouse.dblclick(contentBox.x + 10, contentBox.y + 40);
  const input = page.locator("#app input");
  await expect(input).toBeVisible();
  const style = await input.evaluate((el: HTMLInputElement) => ({ left: el.style.left, top: el.style.top }));
  expect(style).toEqual({ left: "0px", top: "24px" });
  await input.fill("Edited");
  await page.keyboard.press("Enter");
  await expect(page.locator("#app input")).toHaveCount(0);

  expect(await getCell(page, 18, 0)).toBe("Edited");
});

test("dragging a header column boundary resizes the column and shifts later columns", async ({
  page
}) => {
  await gotoFixture(page);

  const headerBox = await page.locator("#app canvas").nth(2).boundingBox();
  if (!headerBox) {
    throw new Error("missing header canvas box");
  }

  expect(await statusHeaderHostLeft(page)).toBe("410px");

  // sku|name boundary sits at x=100 (sku is frozen, so this position is
  // scroll-independent) — drag it 30px to the right.
  await page.mouse.move(headerBox.x + 100, headerBox.y + LEAF_Y);
  await page.mouse.down();
  await page.mouse.move(headerBox.x + 130, headerBox.y + LEAF_Y);
  await page.mouse.up();
  await page.waitForTimeout(50);

  // The 420px-wide fixture no longer fits the status column after sku grew
  // by 30px — scroll it back into view (DomHeaderManager correctly
  // unmounted it, same virtualization the "scrolling out" unit tests cover)
  // before checking its new position.
  await page.locator("#app").evaluate((el) => {
    const scroller = el.firstElementChild as HTMLElement;
    scroller.scrollLeft = 100;
    scroller.dispatchEvent(new Event("scroll"));
  });
  await nextFrame(page);

  // status absolute left after resize = 130(sku)+140+90+80 = 440; minus scrollLeft 100.
  expect(await statusHeaderHostLeft(page)).toBe("340px");
});

test("hovering a column boundary shows a col-resize cursor; hovering elsewhere doesn't", async ({ page }) => {
  await gotoFixture(page);

  const headerBox = await page.locator("#app canvas").nth(2).boundingBox();
  if (!headerBox) {
    throw new Error("missing header canvas box");
  }

  const cursorAt = (x: number, y: number) =>
    page.evaluate(
      ({ x, y }) => {
        const canvas = document.querySelectorAll("#app canvas")[2] as HTMLCanvasElement;
        const rect = canvas.getBoundingClientRect();
        canvas.dispatchEvent(
          new PointerEvent("pointermove", { clientX: rect.left + x, clientY: rect.top + y, bubbles: true })
        );
        return canvas.style.cursor;
      },
      { x, y }
    );

  // Same sku|name boundary (x=100) the resize-drag test above uses.
  expect(await cursorAt(100, GROUP_Y)).toBe("");
  expect(await cursorAt(100, LEAF_Y)).toBe("col-resize");
  // Well inside a column, away from any boundary.
  expect(await cursorAt(50, LEAF_Y)).toBe("");
});

test("column resize wins over a custom canvas header action", async ({ page }) => {
  await gotoFixture(page);
  const headerBox = await page.locator("#app canvas").nth(2).boundingBox();
  if (!headerBox) throw new Error("missing header canvas box");

  expect(await getColumnWidth(page, 2)).toBe(90);
  await page.mouse.move(headerBox.x + 330, headerBox.y + LEAF_Y);
  await page.mouse.down();
  await page.mouse.move(headerBox.x + 360, headerBox.y + LEAF_Y);
  await page.mouse.up();
  await nextFrame(page);

  expect(await getColumnWidth(page, 2)).toBe(120);
  expect(await getHeaderHitTestCalls(page)).toBe(0);
  expect(
    await page.evaluate(() => (window as unknown as FixtureWindow).__headerActions ?? [])
  ).toEqual([]);
});

test("a handled canvas header action runs hitTest once and preserves selection", async ({ page }) => {
  await gotoFixture(page);
  const headerBox = await page.locator("#app canvas").nth(2).boundingBox();
  if (!headerBox) throw new Error("missing header canvas box");

  await page.mouse.click(headerBox.x + 285, headerBox.y + LEAF_Y);

  expect(await getHeaderHitTestCalls(page)).toBe(1);
  expect(await getSelection(page)).toMatchObject({
    kind: "cell",
    active: { row: 0, column: 0 }
  });
  expect(
    await page.evaluate(() => (window as unknown as FixtureWindow).__headerActions ?? [])
  ).toMatchObject([
    {
      action: { type: "toggle-sort" },
      column: 2,
      header: { column: 2, title: "Price" },
      structureVersion: 0
    }
  ]);
});

test("a declined canvas header action runs hitTest once and falls through to selection", async ({ page }) => {
  await gotoFixture(page);
  await setHeaderHitMode(page, "decline");
  const headerBox = await page.locator("#app canvas").nth(2).boundingBox();
  if (!headerBox) throw new Error("missing header canvas box");

  await page.mouse.click(headerBox.x + 285, headerBox.y + LEAF_Y);

  expect(await getHeaderHitTestCalls(page)).toBe(1);
  expect(await getSelection(page)).toMatchObject({
    kind: "column",
    active: { column: 2 },
    range: { startColumn: 2, endColumn: 2 }
  });
  expect(
    await page.evaluate(() => (window as unknown as FixtureWindow).__headerActions ?? [])
  ).toEqual([]);
});

test("a throwing canvas header action reports once and falls through to selection", async ({ page }) => {
  await gotoFixture(page);
  await setHeaderHitMode(page, "throw");
  const headerBox = await page.locator("#app canvas").nth(2).boundingBox();
  if (!headerBox) throw new Error("missing header canvas box");

  await page.mouse.click(headerBox.x + 285, headerBox.y + LEAF_Y);

  expect(await getHeaderHitTestCalls(page)).toBe(1);
  expect(await getSelection(page)).toMatchObject({
    kind: "column",
    active: { column: 2 },
    range: { startColumn: 2, endColumn: 2 }
  });
  expect(
    await page.evaluate(() => (window as unknown as FixtureWindow).__headerErrors ?? [])
  ).toEqual([
    {
      eventName: "interaction:canvas-header-hit-test:2",
      message: "forced header hit-test failure"
    }
  ]);
});

test("DOM header nodes are cleaned up on dispose()", async ({ page }) => {
  await gotoFixture(page);
  await expect(page.locator("#app select")).toBeVisible();

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.dispose());

  await expect(page.locator("#app select")).toHaveCount(0);
});

// This fixture's `headerGroups` (["basics": columns 0-1, "sales": columns
// 2-3]) reference plain column indices — a structural change would need to
// remap or invalidate them, deferred to Step 19 when the group data
// structure is redesigned anyway (see the Step 17 plan). Until then,
// insert/delete/move must fail fast, before touching the Store, rather
// than leave `#columns`/ColumnManager/LayoutEngine out of sync with it.
test("insertColumns/deleteColumns/moveColumns throw while headerGroups is configured, and leave the sheet untouched", async ({
  page
}) => {
  await gotoFixture(page);

  const results = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sheet;
    const attempts: Array<() => void> = [
      () => sheet?.insertColumns(0, [{ id: "extra" }]),
      () => sheet?.deleteColumns(0),
      () => sheet?.moveColumns(0, 1, 2)
    ];
    return attempts.map((attempt) => {
      try {
        attempt();
        return { threw: false };
      } catch (error) {
        return { threw: true, code: (error as { code?: string }).code, name: (error as Error).name };
      }
    });
  });

  for (const result of results) {
    expect(result).toEqual({
      threw: true,
      code: "header-groups-block-column-structure-change",
      name: "CyExcelError"
    });
  }

  // Untouched: column 0 is still "sku" ("CY-0001" for row 0).
  expect(await getCell(page, 0, 0)).toBe("CY-0001");
});

// Real report: with headerGroups configured, right-clicking a column header
// and clicking a structural action (e.g. "insert column left", contributed
// by createBasicContextMenuPlugin()) used to throw an uncaught exception
// straight out of ContextMenuService's click handler — the menu never
// closed and the error never surfaced anywhere useful, and clicking the
// item silently did nothing from the user's perspective ("插入列删除列没有
// 效果" — a real user report). Fixed at the source instead: basicMenuPlugin
// now checks PluginContext.hasHeaderGroups() and marks these items
// `disabled` up front, so the click can never reach a throwing
// insertColumns()/deleteColumns() call in the first place — DefaultMenuRenderer
// never even attaches a click listener to a disabled item (see its
// #createItemElement()), so a real click on one is a no-op, not just
// visually inert. The underlying catch-and-close-on-throw behavior (for the
// general case of *any* action throwing) is still covered independently by
// ContextMenuService.test.ts's "catches a synchronous throw" unit test.
test("right-clicking a column header with headerGroups configured marks structural actions disabled and inert", async ({
  page
}) => {
  const pageErrors: string[] = [];
  page.on("pageerror", (error) => pageErrors.push(String(error)));

  await gotoFixture(page);
  const box = await page.locator("#app").boundingBox();
  if (!box) throw new Error("missing app box");

  // Column "price" (index 2, part of the "sales" [2,3] group) — x falls
  // within its band regardless of which header sub-row (group/leaf) y lands in.
  await page.mouse.click(box.x + 285, box.y + 10, { button: "right" });
  const insertLeft = page.locator('[data-item-id="cyexcel:insert-column-left"]');
  await expect(insertLeft).toBeVisible();
  expect(await insertLeft.evaluate((el) => (el as HTMLElement).style.pointerEvents)).toBe("none");

  // A disabled item never gets a click listener attached (see
  // DefaultMenuRenderer#createItemElement) — force-clicking it (bypassing
  // the pointer-events:none actionability check a real user's click would
  // already be stopped by) still can't run the blocked action, proving the
  // no-op is enforced at the listener level, not just via CSS.
  await insertLeft.click({ force: true });
  await page.waitForTimeout(50);

  await expect(page.locator('[role="menu"]')).toHaveCount(1); // still open — no action ran to close it
  expect(pageErrors).toEqual([]);
  // Untouched: the blocked action never reached the Store.
  expect(await getCell(page, 0, 0)).toBe("CY-0001");
});
