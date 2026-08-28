import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface FixtureWindow {
  __cyexcelFullLoopReady?: boolean;
  __sheet?: {
    getCell(row: number, column: number): unknown;
    setCell(row: number, column: number, value: unknown): void;
    undo(): void;
    redo(): void;
    registerShortcut(combo: string, handler: () => void | false): () => void;
  };
  __originalData?: Array<Record<string, unknown>>;
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/full-loop.html");
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelFullLoopReady === true
  );
}

async function getContentCanvasBox(page: Page) {
  const box = await page.locator("#app canvas").first().boundingBox();
  if (!box) {
    throw new Error("missing content canvas bounding box");
  }
  return box;
}

function getCell(page: Page, row: number, column: number): Promise<unknown> {
  return page.evaluate(
    ({ row: r, column: c }) => (window as unknown as FixtureWindow).__sheet?.getCell(r, c),
    { row, column }
  );
}

function undo(page: Page): Promise<void> {
  return page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.undo());
}

function redo(page: Page): Promise<void> {
  return page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.redo());
}

test("renders visible cells as real, non-transparent pixels", async ({ page }) => {
  await gotoFixture(page);

  const hasPaintedPixel = await page.evaluate(() => {
    const canvas = document.querySelectorAll("#app canvas")[0] as HTMLCanvasElement;
    const ctx = canvas.getContext("2d") as CanvasRenderingContext2D;
    const dpr = window.devicePixelRatio || 1;
    const { data } = ctx.getImageData(
      Math.round(2 * dpr),
      Math.round(2 * dpr),
      Math.round(100 * dpr),
      Math.round(16 * dpr)
    );
    for (let i = 3; i < data.length; i += 4) {
      if ((data[i] ?? 0) > 0) {
        return true;
      }
    }
    return false;
  });

  expect(hasPaintedPixel).toBe(true);
});

test("double-click opens an editor positioned at the cell rect", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.dblclick(box.x + 10, box.y + 10); // cell (0, 0)

  const input = page.locator("#app input");
  await expect(input).toBeVisible();
  const style = await input.evaluate((el: HTMLInputElement) => ({
    left: el.style.left,
    top: el.style.top,
    width: el.style.width,
    height: el.style.height
  }));
  expect(style).toEqual({ left: "0px", top: "0px", width: "120px", height: "24px" });
});

test("double-click places the cursor at the end of the text instead of selecting it all", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.dblclick(box.x + 10, box.y + 10); // cell (0, 0), value "Item 0"

  const input = page.locator("#app input");
  await expect(input).toBeVisible();
  const selection = await input.evaluate((el: HTMLInputElement) => ({
    start: el.selectionStart,
    end: el.selectionEnd,
    length: el.value.length
  }));
  // Not select-all (which would be {start: 0, end: length}) — the cursor
  // sits at the end, so typing continues/appends instead of replacing
  // everything, matching Excel's double-click/F2 behavior.
  expect(selection).toEqual({ start: selection.length, end: selection.length, length: selection.length });

  // Typing now appends rather than replacing the selected-all text.
  await page.keyboard.type("!");
  await expect(input).toHaveValue("Item 0!");
});

test("typing a new value and pressing Enter commits one undoable/redoable change", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.dblclick(box.x + 10, box.y + 10); // initial value "Item 0"
  // The cursor sits at the end on open (no more select-all — see the
  // dedicated cursor-position test), so replace the whole value explicitly.
  await page.locator("#app input").fill("Changed");
  await page.keyboard.press("Enter");

  await expect(page.locator("#app input")).toHaveCount(0);
  expect(await getCell(page, 0, 0)).toBe("Changed");

  await undo(page);
  expect(await getCell(page, 0, 0)).toBe("Item 0");

  await redo(page);
  expect(await getCell(page, 0, 0)).toBe("Changed");
});

test("setCell() leaves the caller's original row object alone — `data` is an import, not the storage", async ({
  page
}) => {
  await gotoFixture(page);

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.setCell(0, 1, 999));

  // 这条断言的方向在 ADR 0004 §3.2 反了过来：`data` 曾经**就是**表格的存储，
  // `setCell()` 会就地改写调用方传进来的行对象。现在它只被 `rowsToWorkbookSnapshot`
  // 读一遍导进稀疏矩阵，此后两边再无联系。这是意图之内的破坏性变更（4B 设计
  // §3.3），所以这里钉的是「写回真的断了」，而不是删掉一条用例了事。
  const result = await page.evaluate(() => ({
    originalPrice: (window as unknown as FixtureWindow).__originalData?.[0]?.price,
    cell: (window as unknown as FixtureWindow).__sheet?.getCell(0, 1)
  }));

  expect(result.originalPrice).toBe(0);
  expect(result.cell).toBe(999);
});

test("Ctrl+Z/Ctrl+Y undo and redo from the keyboard, not just the programmatic API", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.click(box.x + 10, box.y + 10); // selects and focuses cell (0, 0), "Item 0"
  await page.keyboard.type("Changed");
  await page.keyboard.press("Enter");

  await expect(page.locator("#app input")).toHaveCount(0);
  expect(await getCell(page, 0, 0)).toBe("Changed");

  await page.keyboard.press("Control+z");
  expect(await getCell(page, 0, 0)).toBe("Item 0");

  await page.keyboard.press("Control+y");
  expect(await getCell(page, 0, 0)).toBe("Changed");

  await page.keyboard.press("Control+z");
  expect(await getCell(page, 0, 0)).toBe("Item 0");

  await page.keyboard.press("Control+Shift+z");
  expect(await getCell(page, 0, 0)).toBe("Changed");
});

test("Enter on a selected cell opens the editor when not already editing", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.click(box.x + 10, box.y + 10); // select cell (0, 0), no editor yet
  await expect(page.locator("#app input")).toHaveCount(0);

  await page.keyboard.press("Enter");
  await expect(page.locator("#app input")).toBeVisible();
});

test("typing a printable character on a selected cell starts editing with that character", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.click(box.x + 130, box.y + 10); // select cell (0, 1), "price" column
  await page.keyboard.press("9");

  await expect(page.locator("#app input")).toHaveValue("9");
});

test("Escape cancels editing without changing the value", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.dblclick(box.x + 10, box.y + 10);
  await page.keyboard.type("Discarded");
  await page.keyboard.press("Escape");

  await expect(page.locator("#app input")).toHaveCount(0);
  expect(await getCell(page, 0, 0)).toBe("Item 0");
});

test("Enter during IME composition does not prematurely commit", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.dblclick(box.x + 10, box.y + 10);
  const input = page.locator("#app input");

  await input.evaluate((el) =>
    el.dispatchEvent(new CompositionEvent("compositionstart", { bubbles: true }))
  );
  await page.keyboard.press("Enter");
  await expect(input).toBeVisible(); // still editing — Enter was swallowed by composition

  await input.evaluate((el) =>
    el.dispatchEvent(new CompositionEvent("compositionend", { bubbles: true }))
  );
  await page.keyboard.press("Enter");
  await expect(page.locator("#app input")).toHaveCount(0);
});

test("scrolling commits the active edit per the default scrollCommitPolicy", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.dblclick(box.x + 10, box.y + 10);
  await page.locator("#app input").fill("ScrolledAway");

  await page.locator("#app").evaluate((el) => {
    const scroller = el.firstElementChild as HTMLElement;
    scroller.scrollTop = 100;
    scroller.dispatchEvent(new Event("scroll"));
  });

  await expect(page.locator("#app input")).toHaveCount(0);
  expect(await getCell(page, 0, 0)).toBe("ScrolledAway");
});

test("scrolling repaints the canvas instead of leaving the previous rows on screen", async ({
  page
}) => {
  await gotoFixture(page);

  const sampleTopLeft = () =>
    page.evaluate(() => {
      const canvas = document.querySelectorAll("#app canvas")[0] as HTMLCanvasElement;
      const ctx = canvas.getContext("2d") as CanvasRenderingContext2D;
      const dpr = window.devicePixelRatio || 1;
      const { data } = ctx.getImageData(
        Math.round(2 * dpr),
        Math.round(2 * dpr),
        Math.round(100 * dpr),
        Math.round(16 * dpr)
      );
      return Array.from(data);
    });

  const before = await sampleTopLeft();

  await page.locator("#app").evaluate((el) => {
    const scroller = el.firstElementChild as HTMLElement;
    scroller.scrollTop = 200; // past several rows at rowHeight 24
    scroller.dispatchEvent(new Event("scroll"));
  });
  // The repaint is scheduled via requestAnimationFrame, not synchronous.
  await page.waitForFunction(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
  );

  const after = await sampleTopLeft();

  // Whatever was painted at the top-left corner before the scroll must not
  // still be there after — the previously visible rows scrolled out and
  // different rows scrolled in. Before the ContentLayerRenderer fix, this
  // region was byte-identical because cell rects were never translated by
  // the scroll offset, so the newly "visible" rows were drawn off-canvas
  // and the stale pixels from row 0 were never cleared.
  expect(after).not.toEqual(before);
});

test("the selection box repaints after scrolling instead of staying frozen at its old screen position", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.click(box.x + 10, box.y + 10); // select cell (0, 0)
  await page.waitForFunction(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
  );

  const sampleSelectionCanvas = () =>
    page.evaluate(() => {
      const canvas = document.querySelectorAll("#app canvas")[1] as HTMLCanvasElement;
      const ctx = canvas.getContext("2d") as CanvasRenderingContext2D;
      return Array.from(ctx.getImageData(0, 0, canvas.width, canvas.height).data);
    });

  const before = await sampleSelectionCanvas();

  await page.locator("#app").evaluate((el) => {
    const scroller = el.firstElementChild as HTMLElement;
    scroller.scrollTop = 500;
    scroller.dispatchEvent(new Event("scroll"));
  });
  await page.waitForFunction(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
  );

  const after = await sampleSelectionCanvas();

  // Cell (0, 0) scrolled far out of view, so its selection box should no
  // longer be drawn where it used to be — `invalidateAll()` (called on
  // every scroll) only marked the content layer dirty, never selection, so
  // paintSelection() was never re-invoked and the box just sat at its last
  // painted screen position regardless of scroll.
  expect(after).not.toEqual(before);
});

test("double-clicking a cell after scrolling opens an editor positioned inside the viewport", async ({
  page
}) => {
  await gotoFixture(page);

  await page.locator("#app").evaluate((el) => {
    const scroller = el.firstElementChild as HTMLElement;
    scroller.scrollTop = 600; // deep enough that row 0's absolute position (y=0) is long gone
    scroller.dispatchEvent(new Event("scroll"));
  });
  await page.waitForFunction(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
  );

  const box = await getContentCanvasBox(page);
  // scrollTop 600 at rowHeight 24 puts row 25 right at the top of the viewport.
  await page.mouse.dblclick(box.x + 10, box.y + 10);

  const input = page.locator("#app input");
  await expect(input).toBeVisible();
  const style = await input.evaluate((el: HTMLInputElement) => ({ left: el.style.left, top: el.style.top }));
  // Viewport-relative, not row 25's absolute content position (600px) — and
  // still visible/editable, not silently torn down by a scroll-into-view
  // triggering the scroll-commit path before the assertion runs.
  expect(style).toEqual({ left: "0px", top: "0px" });
  expect(await getCell(page, 25, 0)).toBe("Item 25");
});

test("the selection box hides while editing instead of doubling up with the editor's own border", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  const sampleSelectionCanvas = () =>
    page.evaluate(() => {
      const canvas = document.querySelectorAll("#app canvas")[1] as HTMLCanvasElement;
      const ctx = canvas.getContext("2d") as CanvasRenderingContext2D;
      let nonTransparent = 0;
      const { data } = ctx.getImageData(0, 0, canvas.width, canvas.height);
      for (let i = 3; i < data.length; i += 4) {
        if ((data[i] ?? 0) > 0) {
          nonTransparent += 1;
        }
      }
      return nonTransparent;
    });

  await page.mouse.click(box.x + 10, box.y + 10); // select cell (0, 0)
  await page.waitForFunction(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
  );
  expect(await sampleSelectionCanvas()).toBeGreaterThan(0); // the box is drawn while just selected

  await page.mouse.dblclick(box.x + 10, box.y + 10); // start editing that same cell
  await page.waitForFunction(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
  );
  expect(await sampleSelectionCanvas()).toBe(0); // hidden — the editor's own border is the indicator now

  await page.keyboard.press("Escape"); // cancel back to plain selection
  await page.waitForFunction(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
  );
  expect(await sampleSelectionCanvas()).toBeGreaterThan(0); // reappears once editing ends
});

test("keyboard navigation past the visible viewport scrolls the container to keep the active cell in view", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.click(box.x + 10, box.y + 10); // select cell (0, 0)
  // Container is 300px tall at rowHeight 24 — ~12 rows visible. 20 presses
  // moves the active cell well past that without any scrolling to follow it.
  for (let i = 0; i < 20; i += 1) {
    await page.keyboard.press("ArrowDown");
  }

  const scrollTop = await page
    .locator("#app")
    .evaluate((el) => (el.firstElementChild as HTMLElement).scrollTop);
  expect(scrollTop).toBeGreaterThan(0);

  // Navigating back up past the top of the now-scrolled view should scroll
  // back toward 0, not leave the active cell stranded above the fold.
  for (let i = 0; i < 20; i += 1) {
    await page.keyboard.press("ArrowUp");
  }
  const scrollTopAfterUp = await page
    .locator("#app")
    .evaluate((el) => (el.firstElementChild as HTMLElement).scrollTop);
  expect(scrollTopAfterUp).toBe(0);
});

test("keyboard focus returns to the content canvas after committing an edit, so Enter/arrow keys keep working", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.click(box.x + 10, box.y + 10); // select cell (0, 0)
  await page.keyboard.press("Enter"); // start editing
  await page.keyboard.press("Enter"); // commit unchanged — moveActive(1, 0) makes (1, 0) active
  await expect(page.locator("#app input")).toHaveCount(0);

  const activeIsCanvas = await page.evaluate(() => document.activeElement?.tagName === "CANVAS");
  expect(activeIsCanvas).toBe(true);

  // Enter must still reach the canvas's keydown handler and open an editor
  // for the now-active cell (1, 0) — not silently do nothing.
  await page.keyboard.press("Enter");
  await expect(page.locator("#app input")).toHaveValue("Item 1");
  await page.keyboard.press("Escape");

  const activeIsCanvasAfterCancel = await page.evaluate(
    () => document.activeElement?.tagName === "CANVAS"
  );
  expect(activeIsCanvasAfterCancel).toBe(true);

  // Arrow keys must still reach the canvas too, moving active to (2, 0).
  await page.keyboard.press("ArrowDown");
  await page.keyboard.press("Enter");
  await expect(page.locator("#app input")).toHaveValue("Item 2");
});

test("keyboard focus returns to the content canvas after canceling an edit", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.dblclick(box.x + 10, box.y + 10);
  await page.keyboard.press("Escape");
  await expect(page.locator("#app input")).toHaveCount(0);

  const activeIsCanvas = await page.evaluate(
    () => document.activeElement?.tagName === "CANVAS"
  );
  expect(activeIsCanvas).toBe(true);
});

test("registerShortcut() registers a custom combo that fires on a real keypress", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);
  await page.mouse.click(box.x + 10, box.y + 10);

  await page.evaluate(() => {
    (window as unknown as { __calls: string[] }).__calls = [];
    (window as unknown as FixtureWindow).__sheet?.registerShortcut("Mod+K", () => {
      (window as unknown as { __calls: string[] }).__calls.push("fired");
    });
  });

  await page.keyboard.press("Control+k");

  const calls = await page.evaluate(() => (window as unknown as { __calls: string[] }).__calls);
  expect(calls).toEqual(["fired"]);
});

test("the disposer returned by registerShortcut() unregisters it, without affecting other keys", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);
  await page.mouse.click(box.x + 10, box.y + 10);

  await page.evaluate(() => {
    (window as unknown as { __calls: string[] }).__calls = [];
    const dispose = (window as unknown as FixtureWindow).__sheet?.registerShortcut("Mod+K", () => {
      (window as unknown as { __calls: string[] }).__calls.push("fired");
    });
    (window as unknown as { __dispose: (() => void) | undefined }).__dispose = dispose;
  });
  await page.evaluate(() => (window as unknown as { __dispose?: () => void }).__dispose?.());

  await page.keyboard.press("Control+k");
  expect(await page.evaluate(() => (window as unknown as { __calls: string[] }).__calls)).toEqual([]);

  // Enter (a different, unrelated built-in shortcut) still works after unregistering Mod+K.
  await page.keyboard.press("Enter");
  await expect(page.locator("#app input")).toBeVisible();
});

test("a shortcut registered for an already-built-in combo overrides it by default", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.click(box.x + 10, box.y + 10);
  await page.keyboard.type("Changed");
  await page.keyboard.press("Enter");
  expect(await getCell(page, 0, 0)).toBe("Changed");

  await page.evaluate(() => {
    (window as unknown as { __calls: string[] }).__calls = [];
    (window as unknown as FixtureWindow).__sheet?.registerShortcut("Mod+Z", () => {
      (window as unknown as { __calls: string[] }).__calls.push("intercepted");
    });
  });

  await page.mouse.click(box.x + 10, box.y + 10);
  await page.keyboard.press("Control+z");

  expect(await page.evaluate(() => (window as unknown as { __calls: string[] }).__calls)).toEqual([
    "intercepted"
  ]);
  // Built-in undo() never ran — the value is unchanged.
  expect(await getCell(page, 0, 0)).toBe("Changed");
});

test("repeated edit cycles leave no editor DOM behind", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  for (let i = 0; i < 3; i += 1) {
    await page.mouse.dblclick(box.x + 10, box.y + 10);
    await page.keyboard.press("Escape");
  }
  for (let i = 0; i < 3; i += 1) {
    await page.mouse.dblclick(box.x + 10, box.y + 10);
    await page.keyboard.type("x");
    await page.keyboard.press("Enter");
  }

  const inputCount = await page.evaluate(() => document.querySelectorAll("input").length);
  expect(inputCount).toBe(0);
});
