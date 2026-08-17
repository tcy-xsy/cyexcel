import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface SelectionSnapshot {
  readonly kind: "cell" | "row" | "column" | "all";
  readonly active: { readonly row: number; readonly column: number };
  readonly range: {
    readonly startRow: number;
    readonly startColumn: number;
    readonly endRow: number;
    readonly endColumn: number;
  };
}

interface FixtureWindow {
  __cyexcelRowHeadersReady?: boolean;
  __rowHeaderMenuRuns?: number[];
  __headerCanvasHasCapture?(kind: "column" | "row"): boolean;
  __sheet?: {
    readonly selection: SelectionSnapshot;
    getCell(row: number, column: number): unknown;
    undo(): void;
    startEdit(row: number, column: number): void;
  };
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/row-headers.html");
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelRowHeadersReady === true
  );
}

function canvasBox(page: Page, index: number) {
  return page.locator("#app canvas").nth(index).boundingBox();
}

function readSelection(page: Page): Promise<SelectionSnapshot | undefined> {
  return page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.selection);
}

/** nonTransparentCount for a small logical-pixel box on a given canvas, DPR-aware. */
function sampleAlpha(
  page: Page,
  canvasIndex: number,
  x: number,
  y: number,
  width: number,
  height: number
): Promise<number> {
  return page.evaluate(
    ({ canvasIndex, x, y, width, height }) => {
      const canvas = document.querySelectorAll("#app canvas")[canvasIndex] as HTMLCanvasElement;
      const ctx = canvas.getContext("2d") as CanvasRenderingContext2D;
      const dpr = window.devicePixelRatio || 1;
      const { data } = ctx.getImageData(
        Math.round(x * dpr),
        Math.round(y * dpr),
        Math.max(1, Math.round(width * dpr)),
        Math.max(1, Math.round(height * dpr))
      );
      let nonTransparentCount = 0;
      for (let i = 3; i < data.length; i += 4) {
        if ((data[i] ?? 0) > 0) {
          nonTransparentCount += 1;
        }
      }
      return nonTransparentCount;
    },
    { canvasIndex, x, y, width, height }
  );
}

test("row header canvas paints real, non-transparent pixels", async ({ page }) => {
  await gotoFixture(page);

  const count = await sampleAlpha(page, 3, 0, 0, 40, 250);
  expect(count).toBeGreaterThan(0);
});

test("clicking a row number selects the whole row, not just its first cell", async ({ page }) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  // Row index 5 (not frozen — freeze.rows is 2): absolute y = headerHeight(20) + 5*20 = 120, centered at +10.
  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 130);
  await page.waitForTimeout(50);

  // Sanity: row 5's own top border (y=100 in selection-canvas-local/body coordinates) is visible.
  const topBorder = await sampleAlpha(page, 1, 20, 99, 20, 3);
  expect(topBorder).toBeGreaterThan(0);

  // The column 0/column 1 boundary (x=60) sits *inside* row 5's selected
  // band (y: 100..120) — a single-cell selection (row 5, col 0 only) would
  // draw a right-edge border exactly there; a whole-row selection has no
  // seam at any interior column boundary, only at its own top/bottom and
  // at the grid's outer left/right edges.
  const interiorColumnBoundary = await sampleAlpha(page, 1, 59, 105, 3, 10);
  expect(interiorColumnBoundary).toBe(0);
});

test("dragging across row numbers selects a contiguous band of rows", async ({ page }) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  // Row 3 center: headerHeight(20) + 3*20 + 10 = 90. Row 6 center: 20 + 6*20 + 10 = 150.
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 90);
  await page.mouse.down();
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 150);
  await page.mouse.up();
  await page.waitForTimeout(50);

  // Band top (row 3's own top, content y=60) and bottom (row 6's own
  // bottom, content y=140) are both visible...
  const topBorder = await sampleAlpha(page, 1, 20, 59, 20, 3);
  const bottomBorder = await sampleAlpha(page, 1, 20, 139, 20, 3);
  expect(topBorder).toBeGreaterThan(0);
  expect(bottomBorder).toBeGreaterThan(0);
  // ...but there's no seam at an interior row boundary (row 4/5, content
  // y=100) — one continuous band, not the drag stopping short partway
  // through and leaving a border where it stopped.
  const interiorRowBoundary = await sampleAlpha(page, 1, 20, 99, 20, 3);
  expect(interiorRowBoundary).toBe(0);
});

test("a small jitter after clicking a row number does not extend the selection into the next row", async ({
  page
}) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  // Row 5 center (same position the plain-click test above uses), then a
  // couple of pixels of incidental movement before releasing — well under
  // DRAG_THRESHOLD_PX (4).
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 130);
  await page.mouse.down();
  await page.mouse.move(rowHeaderBox.x + 21, rowHeaderBox.y + 131);
  await page.mouse.up();
  await page.waitForTimeout(50);

  // Row 5's own bottom border (content y=120, row 5/6 boundary) is still
  // there — the selection stayed exactly one row tall, it didn't extend
  // into row 6 from the jitter.
  const row5BottomBorder = await sampleAlpha(page, 1, 20, 119, 20, 3);
  expect(row5BottomBorder).toBeGreaterThan(0);
});

test("a secondary row-header down-move-up does not start drag selection or pointer capture", async ({
  page
}) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 70); // row 2
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 130); // row 5
  await page.mouse.down({ button: "right" });
  expect(
    await page.evaluate(() =>
      (window as unknown as FixtureWindow).__headerCanvasHasCapture?.("row")
    )
  ).toBe(false);
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 190); // row 8
  await page.mouse.up({ button: "right" });
  await page.keyboard.press("Escape");

  expect(await readSelection(page)).toMatchObject({
    kind: "row",
    active: { row: 5, column: 0 },
    range: { startRow: 5, startColumn: 0, endRow: 5, endColumn: 4 }
  });
});

test("row and column headers use their fresh client origin after ancestor reflow", async ({ page }) => {
  await gotoFixture(page);

  await page.locator("#fixture-offset").evaluate((element) => {
    element.style.left = "136px";
    element.style.top = "84px";
  });
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");
  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 130); // row 5
  expect(await readSelection(page)).toMatchObject({
    kind: "row",
    active: { row: 5, column: 0 }
  });

  await page.locator("#fixture-offset").evaluate((element) => {
    element.style.left = "46px";
    element.style.top = "122px";
  });
  const columnHeaderBox = await canvasBox(page, 2);
  if (!columnHeaderBox) throw new Error("missing column header canvas box");
  await page.mouse.click(columnHeaderBox.x + 210, columnHeaderBox.y + 10); // column 3
  expect(await readSelection(page)).toMatchObject({
    kind: "column",
    active: { row: 0, column: 3 }
  });
});

test("shift-clicking a second row number extends the selection to span both rows", async ({ page }) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 70); // row 2
  await page.keyboard.down("Shift");
  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 130); // row 5
  await page.keyboard.up("Shift");
  await page.waitForTimeout(50);

  // Band spans row 2 (content y=40) through row 5 (content y=120).
  const topBorder = await sampleAlpha(page, 1, 20, 39, 20, 3);
  const bottomBorder = await sampleAlpha(page, 1, 20, 119, 20, 3);
  expect(topBorder).toBeGreaterThan(0);
  expect(bottomBorder).toBeGreaterThan(0);
});

test("right-clicking inside a multi-row selection preserves it; right-clicking outside it replaces it", async ({
  page
}) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  // Drag-select rows 3-5 (content y: 60..120).
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 90); // row 3
  await page.mouse.down();
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 130); // row 5
  await page.mouse.up();
  await page.waitForTimeout(50);

  // Right-click row 4, inside the selection — the band must survive
  // unchanged, not collapse to just row 4 (same reasoning as the earlier
  // "right-click inside an existing multi-cell selection preserves it" fix
  // for ordinary cells).
  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 110, { button: "right" }); // row 4
  await expect(page.locator('[role="menu"]')).toBeVisible();
  await page.keyboard.press("Escape");
  await expect(page.locator('[role="menu"]')).toHaveCount(0);
  await page.waitForTimeout(50);

  let topBorder = await sampleAlpha(page, 1, 20, 59, 20, 3);
  const bottomBorder = await sampleAlpha(page, 1, 20, 119, 20, 3);
  expect(topBorder).toBeGreaterThan(0);
  expect(bottomBorder).toBeGreaterThan(0);

  // Right-click row 0, outside the rows 3-5 selection — this one *does*
  // replace it with just row 0, matching a plain click. (Row 0, not the
  // fixture's last visible row: a menu opened near the bottom edge of the
  // viewport can run past it and trigger the browser's native
  // scroll-into-view on the menu's own focus() call — a real, pre-existing
  // gap in ContextMenuService/DefaultMenuRenderer unrelated to this
  // feature, reproducible with an ordinary cell right-click too. Left as a
  // follow-up; picking a row with headroom below it sidesteps it here.)
  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 30, { button: "right" }); // row 0
  await expect(page.locator('[role="menu"]')).toBeVisible();
  await page.keyboard.press("Escape");
  await expect(page.locator('[role="menu"]')).toHaveCount(0);
  await page.waitForTimeout(50);

  // The old band's top border (content y=60) is gone...
  topBorder = await sampleAlpha(page, 1, 20, 59, 20, 3);
  expect(topBorder).toBe(0);
  // ...replaced by row 0's own band (content y: 0..20).
  const row0Top = await sampleAlpha(page, 1, 20, 0, 20, 3);
  const row0Bottom = await sampleAlpha(page, 1, 20, 19, 20, 3);
  expect(row0Top).toBeGreaterThan(0);
  expect(row0Bottom).toBeGreaterThan(0);
});

test("clicking a column header selects the whole column, not just its first cell", async ({ page }) => {
  await gotoFixture(page);
  const headerBox = await canvasBox(page, 2);
  if (!headerBox) throw new Error("missing header canvas box");

  // Column index 3 spans x: [180, 240) — click its header, centered vertically.
  await page.mouse.click(headerBox.x + 210, headerBox.y + 10);
  await page.waitForTimeout(50);

  // Sanity: the selection's own top border (y=0) is visible.
  const topBorder = await sampleAlpha(page, 1, 190, 0, 20, 3);
  expect(topBorder).toBeGreaterThan(0);

  // Row 0/row 1 boundary (y=20) sits inside the selected column's full
  // height — a single-cell selection would have a bottom border exactly
  // there; a whole-column selection has no seam at any interior row
  // boundary.
  const interiorRowBoundary = await sampleAlpha(page, 1, 190, 19, 20, 3);
  expect(interiorRowBoundary).toBe(0);
});

test("dragging across column headers selects a contiguous band of columns", async ({ page }) => {
  await gotoFixture(page);
  const headerBox = await canvasBox(page, 2);
  if (!headerBox) throw new Error("missing header canvas box");

  // Column 1 center (x=90) to column 3 center (x=210) — columns are 60px
  // wide each, c0: [0,60) .. c4: [240,300).
  await page.mouse.move(headerBox.x + 90, headerBox.y + 10);
  await page.mouse.down();
  await page.mouse.move(headerBox.x + 210, headerBox.y + 10);
  await page.mouse.up();
  await page.waitForTimeout(50);

  // Band left (column 1's own left, x=60) and right (column 3's own
  // right, x=240) are both visible...
  const leftBorder = await sampleAlpha(page, 1, 59, 5, 3, 10);
  const rightBorder = await sampleAlpha(page, 1, 239, 5, 3, 10);
  expect(leftBorder).toBeGreaterThan(0);
  expect(rightBorder).toBeGreaterThan(0);
  // ...but no seam at an interior column boundary (column 2/3, x=180).
  const interiorColumnBoundary = await sampleAlpha(page, 1, 179, 5, 3, 10);
  expect(interiorColumnBoundary).toBe(0);
});

test("a small jitter after clicking a column header does not extend the selection into the next column", async ({
  page
}) => {
  await gotoFixture(page);
  const headerBox = await canvasBox(page, 2);
  if (!headerBox) throw new Error("missing header canvas box");

  // Column 3 center (same position the plain-click test above uses), then
  // a couple of pixels of incidental movement before releasing — well
  // under DRAG_THRESHOLD_PX (4).
  await page.mouse.move(headerBox.x + 210, headerBox.y + 10);
  await page.mouse.down();
  await page.mouse.move(headerBox.x + 211, headerBox.y + 11);
  await page.mouse.up();
  await page.waitForTimeout(50);

  // Column 3's own right border (x=240, column 3/4 boundary) is still
  // there — the selection stayed exactly one column wide.
  const column3RightBorder = await sampleAlpha(page, 1, 239, 5, 3, 10);
  expect(column3RightBorder).toBeGreaterThan(0);
});

test("shift-clicking a second column header extends the selection to span both columns", async ({ page }) => {
  await gotoFixture(page);
  const headerBox = await canvasBox(page, 2);
  if (!headerBox) throw new Error("missing header canvas box");

  await page.mouse.click(headerBox.x + 90, headerBox.y + 10); // column 1
  await page.keyboard.down("Shift");
  await page.mouse.click(headerBox.x + 210, headerBox.y + 10); // column 3
  await page.keyboard.up("Shift");
  await page.waitForTimeout(50);

  const leftBorder = await sampleAlpha(page, 1, 59, 5, 3, 10);
  const rightBorder = await sampleAlpha(page, 1, 239, 5, 3, 10);
  expect(leftBorder).toBeGreaterThan(0);
  expect(rightBorder).toBeGreaterThan(0);
});

test("right-clicking inside a multi-column selection preserves it; right-clicking outside it replaces it", async ({
  page
}) => {
  await gotoFixture(page);
  const headerBox = await canvasBox(page, 2);
  if (!headerBox) throw new Error("missing header canvas box");

  // Drag-select columns 1-3 (x: 60..240).
  await page.mouse.move(headerBox.x + 90, headerBox.y + 10); // column 1
  await page.mouse.down();
  await page.mouse.move(headerBox.x + 210, headerBox.y + 10); // column 3
  await page.mouse.up();
  await page.waitForTimeout(50);

  // Right-click column 2, inside the selection — the band must survive
  // unchanged, not collapse to just column 2.
  await page.mouse.click(headerBox.x + 150, headerBox.y + 10, { button: "right" }); // column 2
  await expect(page.locator('[role="menu"]')).toBeVisible();
  await page.keyboard.press("Escape");
  await expect(page.locator('[role="menu"]')).toHaveCount(0);
  await page.waitForTimeout(50);

  let leftBorder = await sampleAlpha(page, 1, 59, 5, 3, 10);
  const rightBorder = await sampleAlpha(page, 1, 239, 5, 3, 10);
  expect(leftBorder).toBeGreaterThan(0);
  expect(rightBorder).toBeGreaterThan(0);

  // Right-click column 4, outside the columns 1-3 selection — this one
  // *does* replace it with just column 4, matching a plain click.
  await page.mouse.click(headerBox.x + 270, headerBox.y + 10, { button: "right" }); // column 4
  await expect(page.locator('[role="menu"]')).toBeVisible();
  await page.keyboard.press("Escape");
  await expect(page.locator('[role="menu"]')).toHaveCount(0);
  await page.waitForTimeout(50);

  // The old band's left border (x=60) is gone...
  leftBorder = await sampleAlpha(page, 1, 59, 5, 3, 10);
  expect(leftBorder).toBe(0);
  // ...replaced by column 4's own band (x: 240..300).
  const column4Left = await sampleAlpha(page, 1, 239, 5, 3, 10);
  expect(column4Left).toBeGreaterThan(0);
});

test("clicking the corner selects everything", async ({ page }) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  // y=10 is within the corner band (headerHeight is 20).
  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 10);

  const interiorColumnBoundary = await sampleAlpha(page, 1, 59, 5, 3, 10);
  const interiorRowBoundary = await sampleAlpha(page, 1, 5, 19, 10, 3);
  expect(interiorColumnBoundary).toBe(0);
  expect(interiorRowBoundary).toBe(0);
});

/** Mean absolute per-byte difference between two ImageData buffers of the same size — text anti-aliasing can jitter a couple of RGB units between identical-looking repaints, so exact byte equality is too strict. */
function meanAbsoluteDiff(a: number[], b: number[]): number {
  let total = 0;
  for (let i = 0; i < a.length; i += 1) {
    total += Math.abs((a[i] ?? 0) - (b[i] ?? 0));
  }
  return total / a.length;
}

function readRowHeaderRegion(page: Page, y: number, height: number): Promise<number[]> {
  return page.evaluate(
    ({ y, height }) => {
      const canvas = document.querySelectorAll("#app canvas")[3] as HTMLCanvasElement;
      const ctx = canvas.getContext("2d") as CanvasRenderingContext2D;
      const dpr = window.devicePixelRatio || 1;
      return Array.from(
        ctx.getImageData(0, Math.round(y * dpr), Math.round(40 * dpr), Math.round(height * dpr)).data
      );
    },
    { y, height }
  );
}

test("the frozen row header band does not move when the grid scrolls, but the scrollable band does", async ({
  page
}) => {
  await gotoFixture(page);

  const beforeFrozen = await readRowHeaderRegion(page, 0, 40);
  const beforeScrollable = await readRowHeaderRegion(page, 150, 20);

  await page.locator("#app").evaluate((el) => {
    const scroller = el.firstElementChild as HTMLElement;
    scroller.scrollTop = 200;
    scroller.dispatchEvent(new Event("scroll"));
  });
  await page.waitForFunction(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
  );

  const afterFrozen = await readRowHeaderRegion(page, 0, 40);
  const afterScrollable = await readRowHeaderRegion(page, 150, 20);

  expect(meanAbsoluteDiff(afterFrozen, beforeFrozen)).toBeLessThan(1);
  expect(meanAbsoluteDiff(afterScrollable, beforeScrollable)).toBeGreaterThan(1);
});

test("right-clicking the row header band opens a menu with a row-header-target action", async ({ page }) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 130, { button: "right" }); // row index 5

  await expect(page.locator('[data-item-id="log-row-header"]')).toBeVisible();
  await page.locator('[data-item-id="log-row-header"]').click();

  const runs = await page.evaluate(() => (window as unknown as FixtureWindow).__rowHeaderMenuRuns ?? []);
  expect(runs).toEqual([5]);
});

function getCell(page: Page, row: number, column: number): Promise<unknown> {
  return page.evaluate(
    ({ row: r, column: c }) => (window as unknown as FixtureWindow).__sheet?.getCell(r, c),
    { row, column }
  );
}

test("createBasicContextMenuPlugin's 'insert row above' inserts a blank row and shifts the rest down", async ({
  page
}) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  expect(await getCell(page, 5, 0)).toBe("5-0");

  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 130, { button: "right" }); // row index 5
  await page.locator('[data-item-id="cyexcel:insert-row-above"]').click();

  expect(await getCell(page, 5, 0)).toBeUndefined(); // the new blank row
  expect(await getCell(page, 6, 0)).toBe("5-0"); // the original row 5, shifted down
});

test("createBasicContextMenuPlugin's 'insert row below' inserts a blank row after the clicked row", async ({
  page
}) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 130, { button: "right" }); // row index 5
  await page.locator('[data-item-id="cyexcel:insert-row-below"]').click();

  expect(await getCell(page, 5, 0)).toBe("5-0"); // row 5 itself untouched
  expect(await getCell(page, 6, 0)).toBeUndefined(); // the new blank row
  expect(await getCell(page, 7, 0)).toBe("6-0"); // the original row 6, shifted down
});

test("createBasicContextMenuPlugin's 'delete row' removes the row and shifts the rest up, and undo restores it", async ({
  page
}) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 130, { button: "right" }); // row index 5
  await page.locator('[data-item-id="cyexcel:delete-row"]').click();

  expect(await getCell(page, 5, 0)).toBe("6-0"); // what was row 6 shifted up to row 5

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.undo());

  expect(await getCell(page, 5, 0)).toBe("5-0"); // fully restored
});

test("deleting the row currently being edited ends the edit instead of leaving a stale editor", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await canvasBox(page, 0); // content canvas
  if (!box) throw new Error("missing content canvas box");
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  // Double-click cell (5, 0) to start editing it — contentCanvas is already
  // positioned below the header band, so its own local y=0 is content y=0;
  // row 5 (rowHeight 20) is centered at local y = 5*20 + 10 = 110.
  await page.mouse.dblclick(box.x + 10, box.y + 5 * 20 + 10);
  await expect(page.locator("#app input")).toBeVisible();

  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 130, { button: "right" }); // row index 5
  await page.locator('[data-item-id="cyexcel:delete-row"]').click();

  await expect(page.locator("#app input")).toHaveCount(0);
});

test("with rowReorder unset (default off), dragging from a selection's edge is still ordinary drag-to-select — no reorder, no indicator", async ({
  page
}) => {
  await gotoFixture(page);
  const rowHeaderBox = await canvasBox(page, 3);
  if (!rowHeaderBox) throw new Error("missing row header canvas box");

  await page.mouse.click(rowHeaderBox.x + 20, rowHeaderBox.y + 130); // selects row 5
  // Row 5's top edge (headerHeight 20 + 5*20 = 120) — where a rowReorder:true
  // fixture would treat this as the reorder handle.
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 120);
  await page.mouse.down();
  await page.mouse.move(rowHeaderBox.x + 20, rowHeaderBox.y + 190); // toward row 8
  await page.mouse.up();
  await page.waitForTimeout(50);

  // Data is untouched — this was a selection drag, not a reorder.
  expect(await getCell(page, 5, 0)).toBe("5-0");
  expect(await getCell(page, 8, 0)).toBe("8-0");
  // The selection extended down to row 8: its own bottom border is visible...
  const bottomBorder = await sampleAlpha(page, 1, 20, 179, 20, 3);
  expect(bottomBorder).toBeGreaterThan(0);
  // ...with no seam at an interior row boundary (row 6/7, content y=140).
  const interiorRowBoundary = await sampleAlpha(page, 1, 20, 139, 20, 3);
  expect(interiorRowBoundary).toBe(0);
});
