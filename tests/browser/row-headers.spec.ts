import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface FixtureWindow {
  __cyexcelRowHeadersReady?: boolean;
  __rowHeaderMenuRuns?: number[];
  __sheet?: {
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

test("clicking a column header selects the whole column, not just its first cell", async ({ page }) => {
  await gotoFixture(page);
  const headerBox = await canvasBox(page, 2);
  if (!headerBox) throw new Error("missing header canvas box");

  // Column index 3 spans x: [180, 240) — click its header, centered vertically.
  await page.mouse.click(headerBox.x + 210, headerBox.y + 10);

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
