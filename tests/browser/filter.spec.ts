import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface FilterState {
  readonly column: number;
}

interface FilterSheet {
  getCell(row: number, column: number): unknown;
  setCell(row: number, column: number, value: unknown): void;
  filterByColumn(column: number, predicate: (value: unknown, row: number) => boolean): void;
  clearFilter(): void;
  sortByColumn(column: number, direction: "asc" | "desc"): boolean;
  deleteVisibleRowsInRange(startRow: number, endRow: number): void;
  undo(): void;
  redo(): void;
  readonly scene: {
    getRowHeight(row: number): number;
  };
  readonly filterState: FilterState | null;
  readonly events: {
    on(name: string, handler: (payload: unknown) => void): void;
    off(name: string, handler: (payload: unknown) => void): void;
  };
}

interface FixtureWindow {
  __filterFixtureReady?: boolean;
  __filter?: {
    sheet: FilterSheet;
    getCommitCount(): number;
    getHeaderActions(): { action: { type: string }; column: number }[];
  };
}

async function gotoFixture(page: Page, query = ""): Promise<void> {
  await page.goto(`/tests/browser/filter.html${query}`);
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__filterFixtureReady === true
  );
}

function headerCanvasBox(page: Page) {
  return page.locator("#app canvas").nth(2).boundingBox();
}

/**
 * Clicks inside the reserved click-to-filter strip — immediately left of
 * the click-to-sort strip, `DEFAULT_HEADER_FILTER_ZONE_WIDTH` (18px) wide,
 * so `columnRightX - 26` always lands inside it (18 for the sort zone plus
 * a few px of margin into the filter zone).
 */
async function clickHeaderFilterZone(page: Page, columnRightX: number): Promise<void> {
  const headerBox = await headerCanvasBox(page);
  if (!headerBox) throw new Error("missing header canvas box");
  await page.mouse.click(headerBox.x + columnRightX - 26, headerBox.y + 10);
}

// Column x-ranges in filter.html: name [0,120), category [120,220), score [220,300).
const CATEGORY_COLUMN_RIGHT_X = 220;

test("filterByColumn hides non-matching rows without touching stored data", async ({ page }) => {
  await gotoFixture(page);

  const before = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__filter!.sheet;
    return [0, 1, 2, 3].map((row) => sheet.scene.getRowHeight(row));
  });
  expect(before).toEqual([20, 20, 20, 20]);

  const result = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__filter!.sheet;
    sheet.filterByColumn(1, (value) => value === "fruit");
    return {
      heights: [0, 1, 2, 3].map((row) => sheet.scene.getRowHeight(row)),
      // Names read straight through: rows 0/1 (fruit) visible, 2/3 (veg) filtered.
      names: [0, 1, 2, 3].map((row) => sheet.getCell(row, 0))
    };
  });

  expect(result.heights).toEqual([20, 20, 0, 0]);
  // Stored data at every index is exactly what it always was — filtering
  // never moves or clears anything (ADR 0003 §3.2).
  expect(result.names).toEqual(["apple", "banana", "carrot", "daikon"]);
});

test("clearFilter restores every row's real height pixel for pixel", async ({ page }) => {
  await gotoFixture(page);

  const heights = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__filter!.sheet;
    sheet.setCell(0, 0, "apple"); // no-op edit, just to exercise the model first
    sheet.filterByColumn(1, (value) => value === "veg");
    sheet.clearFilter();
    return [0, 1, 2, 3].map((row) => sheet.scene.getRowHeight(row));
  });

  expect(heights).toEqual([20, 20, 20, 20]);
});

test("filtering never enters the undo stack and never clears an existing redo entry", async ({ page }) => {
  await gotoFixture(page);

  const outcome = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__filter!.sheet;
    // Sort by score (column 2): scores are [3,1,2,5], so ascending genuinely
    // reorders (unlike name/category, which both happen to already be
    // sorted in the fixture data) — banana(1) sorts to the top.
    sheet.sortByColumn(2, "asc");
    const afterSortTop = sheet.getCell(0, 0);
    sheet.undo();
    const afterUndoTop = sheet.getCell(0, 0);

    // Filtering must not touch that redo entry.
    sheet.filterByColumn(1, (value) => value === "fruit");
    sheet.redo();
    const afterRedoTop = sheet.getCell(0, 0);

    return { afterSortTop, afterUndoTop, afterRedoTop };
  });

  expect(outcome.afterSortTop).toBe("banana"); // lowest score (1) sorts first
  expect(outcome.afterUndoTop).toBe("apple"); // back to original order
  // redo() replayed the sort even though a filter ran in between.
  expect(outcome.afterRedoTop).toBe(outcome.afterSortTop);
});

test("undo skips over a filter entirely, reverting the prior real edit instead", async ({ page }) => {
  await gotoFixture(page);

  const outcome = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__filter!.sheet;
    sheet.setCell(0, 0, "edited");
    sheet.filterByColumn(1, (value) => value === "fruit");
    const beforeUndo = { cell: sheet.getCell(0, 0), row0Height: sheet.scene.getRowHeight(0) };
    sheet.undo();
    const afterUndo = { cell: sheet.getCell(0, 0), row2Height: sheet.scene.getRowHeight(2) };
    return { beforeUndo, afterUndo };
  });

  expect(outcome.beforeUndo).toEqual({ cell: "edited", row0Height: 20 });
  // The setCell was undone...
  expect(outcome.afterUndo.cell).toBe("apple");
  // ...but the filter (never recorded) is still in effect — row 2 (veg) still hidden.
  expect(outcome.afterUndo.row2Height).toBe(0);
});

test("clicking the header's filter zone emits a host-driven notification, applying nothing itself", async ({ page }) => {
  await gotoFixture(page);

  await clickHeaderFilterZone(page, CATEGORY_COLUMN_RIGHT_X);

  const outcome = await page.evaluate(() => {
    const fixture = (window as unknown as FixtureWindow).__filter!;
    return {
      actions: fixture.getHeaderActions(),
      filterState: fixture.sheet.filterState,
      heights: [0, 1, 2, 3].map((row) => fixture.sheet.scene.getRowHeight(row))
    };
  });

  expect(outcome.actions).toEqual([{ action: { type: "toggle-filter" }, column: 1 }]);
  // Core never applies a predicate on its own — no default value-list UI.
  expect(outcome.filterState).toBeNull();
  expect(outcome.heights).toEqual([20, 20, 20, 20]);
});

test("a custom header renderer's toggle-filter is also a pure notification", async ({ page }) => {
  await gotoFixture(page, "?header=custom");

  const headerBox = await headerCanvasBox(page);
  if (!headerBox) throw new Error("missing header canvas box");
  await page.mouse.click(headerBox.x + 130, headerBox.y + 10);

  const actions = await page.evaluate(() =>
    (window as unknown as FixtureWindow).__filter!.getHeaderActions()
  );
  expect(actions).toEqual([{ action: { type: "toggle-filter" }, column: 1 }]);
});

test("a column with filterable: false claims no filter zone", async ({ page }) => {
  await gotoFixture(page, "?unfilterable=1");

  await clickHeaderFilterZone(page, CATEGORY_COLUMN_RIGHT_X);

  const actions = await page.evaluate(() =>
    (window as unknown as FixtureWindow).__filter!.getHeaderActions()
  );
  expect(actions).toEqual([]);
});

test("deleteVisibleRowsInRange only deletes rows visible under the active filter", async ({ page }) => {
  await gotoFixture(page);

  const outcome = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__filter!.sheet;
    sheet.filterByColumn(1, (value) => value === "fruit"); // hides carrot(2)/daikon(3)
    sheet.deleteVisibleRowsInRange(0, 3); // spans the whole table, but only 0/1 are visible
    return [0, 1].map((row) => sheet.getCell(row, 0));
  });

  // apple/banana (visible) deleted; carrot/daikon (filtered-out) survive and
  // shift up to fill the gap.
  expect(outcome).toEqual(["carrot", "daikon"]);
});

test("sorting while filtered leaves filtered-out rows in place and sorts only the visible ones", async ({ page }) => {
  await gotoFixture(page);

  const outcome = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__filter!.sheet;
    sheet.filterByColumn(1, (value) => value === "fruit"); // rows 0,1 visible (apple/banana)
    sheet.sortByColumn(2, "asc"); // score: banana=1, apple=3 -> banana before apple among visible slots
    return [0, 1, 2, 3].map((row) => sheet.getCell(row, 0));
  });

  // Visible rows 0/1 reordered among themselves; filtered rows 2/3 untouched.
  expect(outcome).toEqual(["banana", "apple", "carrot", "daikon"]);
});
