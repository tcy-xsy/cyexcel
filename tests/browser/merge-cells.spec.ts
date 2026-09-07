import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface FixtureCellRange {
  readonly startRow: number;
  readonly startColumn: number;
  readonly endRow: number;
  readonly endColumn: number;
}

interface FixtureWindow {
  __cyexcelMergeCellsReady?: boolean;
  __sheet?: {
    getCell(row: number, column: number): unknown;
    mergeCells(range: FixtureCellRange): void;
    unmergeCells(range: FixtureCellRange): void;
    undo(): void;
    redo(): void;
  };
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/merge-cells.html");
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelMergeCellsReady === true
  );
}

function getCell(page: Page, row: number, column: number): Promise<unknown> {
  return page.evaluate(
    ({ row: r, column: c }) => (window as unknown as FixtureWindow).__sheet?.getCell(r, c),
    { row, column }
  );
}

function mergeCells(page: Page, range: FixtureCellRange): Promise<void> {
  return page.evaluate(
    (r) => (window as unknown as FixtureWindow).__sheet?.mergeCells(r),
    range
  );
}

function unmergeCells(page: Page, range: FixtureCellRange): Promise<void> {
  return page.evaluate(
    (r) => (window as unknown as FixtureWindow).__sheet?.unmergeCells(r),
    range
  );
}

test("mergeCells clears every non-anchor cell and keeps the anchor; undo restores the cleared values", async ({
  page
}) => {
  await gotoFixture(page);

  expect(await getCell(page, 0, 0)).toBe("0a");
  expect(await getCell(page, 0, 1)).toBe("0b");
  expect(await getCell(page, 1, 0)).toBe("1a");
  expect(await getCell(page, 1, 1)).toBe("1b");

  await mergeCells(page, { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 });

  expect(await getCell(page, 0, 0)).toBe("0a"); // anchor untouched
  expect(await getCell(page, 0, 1)).toBeNull();
  expect(await getCell(page, 1, 0)).toBeNull();
  expect(await getCell(page, 1, 1)).toBeNull();
  // Outside the merged range, nothing changed.
  expect(await getCell(page, 0, 2)).toBe("0c");

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.undo());

  expect(await getCell(page, 0, 1)).toBe("0b");
  expect(await getCell(page, 1, 0)).toBe("1a");
  expect(await getCell(page, 1, 1)).toBe("1b");

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.redo());

  expect(await getCell(page, 0, 1)).toBeNull();
  expect(await getCell(page, 1, 1)).toBeNull();
});

test("mergeCells rejects a range partially overlapping an existing merge; unmergeCells frees it up for a new one", async ({
  page
}) => {
  await gotoFixture(page);

  await mergeCells(page, { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 });

  const overlapMessage = await page.evaluate(() => {
    try {
      (window as unknown as FixtureWindow).__sheet?.mergeCells({
        startRow: 1,
        startColumn: 1,
        endRow: 2,
        endColumn: 2
      });
      return null;
    } catch (error) {
      return error instanceof Error ? error.message : String(error);
    }
  });
  expect(overlapMessage).toBeTruthy();
  // The rejected attempt must not have staged anything.
  expect(await getCell(page, 2, 2)).toBe("2c");

  // unmergeCells targets any merge intersecting the given range — a single
  // cell inside the merge is enough to remove the whole thing (Excel
  // semantics, see UnmergeCellsCommand).
  await unmergeCells(page, { startRow: 1, startColumn: 1, endRow: 1, endColumn: 1 });

  // With the old merge gone, the same range that used to overlap it now
  // merges cleanly.
  await mergeCells(page, { startRow: 1, startColumn: 1, endRow: 2, endColumn: 2 });

  expect(await getCell(page, 1, 2)).toBeNull();
  expect(await getCell(page, 2, 1)).toBeNull();
  expect(await getCell(page, 2, 2)).toBeNull();
});
