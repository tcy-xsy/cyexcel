import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface FixtureWindow {
  __cyexcelFillHandleReady?: boolean;
  __sheet?: {
    getCell(row: number, column: number): unknown;
    undo(): void;
  };
}

const ROW_HEIGHT = 20;
const COLUMN_WIDTH = 60;

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/fill-handle.html");
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelFillHandleReady === true
  );
}

async function getContentCanvasBox(page: Page) {
  const box = await page.locator("#app canvas").first().boundingBox();
  if (!box) {
    throw new Error("missing content canvas bounding box");
  }
  return box;
}

function cellCenter(row: number, column: number): { x: number; y: number } {
  return { x: column * COLUMN_WIDTH + COLUMN_WIDTH / 2, y: row * ROW_HEIGHT + ROW_HEIGHT / 2 };
}

/** The fill handle's on-screen position for a selection ending at (endRow, endColumn) — the bottom-right corner of that cell. */
function handleCorner(endRow: number, endColumn: number): { x: number; y: number } {
  return { x: (endColumn + 1) * COLUMN_WIDTH, y: (endRow + 1) * ROW_HEIGHT };
}

/** Somewhere inside `row`, staying within `column`'s x-range so a straight-down drag doesn't also drift into an adjacent column. */
function pointInRow(row: number, column: number): { x: number; y: number } {
  return { x: column * COLUMN_WIDTH + COLUMN_WIDTH / 2, y: row * ROW_HEIGHT + ROW_HEIGHT / 2 };
}

async function dragSelectRange(
  page: Page,
  box: { x: number; y: number },
  from: { row: number; column: number },
  to: { row: number; column: number }
): Promise<void> {
  const start = cellCenter(from.row, from.column);
  const end = cellCenter(to.row, to.column);
  await page.mouse.move(box.x + start.x, box.y + start.y);
  await page.mouse.down();
  await page.mouse.move(box.x + end.x, box.y + end.y);
  await page.mouse.up();
}

async function dragFillHandle(
  page: Page,
  box: { x: number; y: number },
  from: { x: number; y: number },
  to: { x: number; y: number }
): Promise<void> {
  await page.mouse.move(box.x + from.x, box.y + from.y);
  await page.mouse.down();
  await page.mouse.move(box.x + to.x, box.y + to.y);
  await page.mouse.up();
}

function getCell(page: Page, row: number, column: number): Promise<unknown> {
  return page.evaluate(
    ({ row: r, column: c }) => (window as unknown as FixtureWindow).__sheet?.getCell(r, c),
    { row, column }
  );
}

test("dragging the fill handle down copies a single known value", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await dragSelectRange(page, box, { row: 0, column: 2 }, { row: 0, column: 2 });
  const handle = handleCorner(0, 2);
  const target = pointInRow(3, 2);
  await dragFillHandle(page, box, handle, target);

  expect(await getCell(page, 1, 2)).toBe("X");
  expect(await getCell(page, 2, 2)).toBe("X");
  expect(await getCell(page, 3, 2)).toBe("X");
  expect(await getCell(page, 4, 2)).toBeUndefined();
});

test("dragging the fill handle down extends a numeric sequence linearly", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await dragSelectRange(page, box, { row: 0, column: 3 }, { row: 2, column: 3 });
  const handle = handleCorner(2, 3);
  const target = pointInRow(5, 3);
  await dragFillHandle(page, box, handle, target);

  expect(await getCell(page, 3, 3)).toBe(4);
  expect(await getCell(page, 4, 3)).toBe(5);
  expect(await getCell(page, 5, 3)).toBe(6);
  expect(await getCell(page, 6, 3)).toBeUndefined();
});

test("a fill-handle drag is a single undo step", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await dragSelectRange(page, box, { row: 0, column: 4 }, { row: 0, column: 4 });
  const handle = handleCorner(0, 4);
  const target = pointInRow(2, 4);
  await dragFillHandle(page, box, handle, target);

  expect(await getCell(page, 1, 4)).toBe("Z");
  expect(await getCell(page, 2, 4)).toBe("Z");

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.undo());

  expect(await getCell(page, 0, 4)).toBe("Z"); // the source cell itself is untouched
  expect(await getCell(page, 1, 4)).toBeUndefined();
  expect(await getCell(page, 2, 4)).toBeUndefined();
});

test("double-clicking the fill handle auto-fills down to match the left neighbor column's length", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await dragSelectRange(page, box, { row: 0, column: 1 }, { row: 0, column: 1 });
  const handle = handleCorner(0, 1);
  await page.mouse.dblclick(box.x + handle.x, box.y + handle.y);

  // Column 0 (the immediate left neighbor) has 10 non-empty rows (0-9).
  expect(await getCell(page, 5, 1)).toBe("Y");
  expect(await getCell(page, 9, 1)).toBe("Y");
  expect(await getCell(page, 10, 1)).toBeUndefined();
});

test("a read-only (clipboard.paste: false) column keeps its values unchanged within a multi-column fill", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await dragSelectRange(page, box, { row: 0, column: 5 }, { row: 0, column: 6 });
  const handle = handleCorner(0, 6);
  const target = pointInRow(2, 6);
  await dragFillHandle(page, box, handle, target);

  // Column 5 is read-only — its extended cells stay blank.
  expect(await getCell(page, 1, 5)).toBeUndefined();
  expect(await getCell(page, 2, 5)).toBeUndefined();
  expect(await getCell(page, 0, 5)).toBe("RO"); // the source cell itself is untouched

  // Column 6 is an ordinary column in the same gesture — it fills normally.
  expect(await getCell(page, 1, 6)).toBe("OK");
  expect(await getCell(page, 2, 6)).toBe("OK");
});

test("a fill source spanning a merge reads the merge's anchor value, not the covered cell's blank", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  // Source rows 0-3 span the merge at rows 1-2 (anchor "MERGED"). Known
  // sequence is ["M0", "MERGED", "MERGED", "M3"] (reading the covered row 2
  // through its anchor) — non-numeric, so the extension cyclically repeats
  // it. If row 2 instead read null (the bug), the repeated sequence would
  // put null at row 6 instead of "MERGED".
  await dragSelectRange(page, box, { row: 0, column: 7 }, { row: 3, column: 7 });
  const handle = handleCorner(3, 7);
  const target = pointInRow(7, 7);
  await dragFillHandle(page, box, handle, target);

  expect(await getCell(page, 4, 7)).toBe("M0");
  expect(await getCell(page, 5, 7)).toBe("MERGED");
  expect(await getCell(page, 6, 7)).toBe("MERGED");
  expect(await getCell(page, 7, 7)).toBe("M3");
});
