import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface SortSheet {
  getCell(row: number, column: number): unknown;
  setRowHeight(row: number, height: number): void;
  sortByColumn(
    column: number,
    direction: "asc" | "desc",
    options?: { signal?: AbortSignal }
  ): boolean;
  undo(): void;
  readonly scene: {
    getRowHeight(row: number): number;
  };
}

interface FixtureWindow {
  __sortFixtureReady?: boolean;
  __sort?: {
    sheet: SortSheet;
    getCommitCount(): number;
    setOnCompare(hook: (() => void) | null): void;
  };
}

async function gotoFixture(page: Page, query = ""): Promise<void> {
  await page.goto(`/tests/browser/sort.html${query}`);
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__sortFixtureReady === true
  );
}

test("sorting by a column reorders the data and moves row heights with it", async ({ page }) => {
  await gotoFixture(page);

  const before = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sort!.sheet;
    sheet.setRowHeight(0, 60);
    return { top: sheet.getCell(0, 0) as string, height: sheet.scene.getRowHeight(0) };
  });
  expect(before.top).toBe("banana");
  expect(before.height).toBe(60);

  const applied = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sort!.sheet;
    return sheet.sortByColumn(0, "asc");
  });
  expect(applied).toBe(true);

  const after = await page.evaluate((tall: string) => {
    const sheet = (window as unknown as FixtureWindow).__sort!.sheet;
    const column = [0, 1, 2].map((row) => sheet.getCell(row, 0) as string);
    return {
      column,
      // The moved row is located by its content, never by a fixed index:
      // proving the height travelled *with the row* is the whole assertion.
      heightOfMovedRow: sheet.scene.getRowHeight(column.indexOf(tall))
    };
  }, before.top);

  expect(after.column).toEqual(["apple", "banana", "cherry"]);
  // The tall row must still be tall wherever it landed.
  expect(after.heightOfMovedRow).toBe(60);
});

test("sorting descending reverses the ascending order", async ({ page }) => {
  await gotoFixture(page);

  const applied = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sort!.sheet;
    return sheet.sortByColumn(0, "desc");
  });
  expect(applied).toBe(true);

  const column = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sort!.sheet;
    return [0, 1, 2].map((row) => sheet.getCell(row, 0) as string);
  });
  expect(column).toEqual(["cherry", "banana", "apple"]);
});

test("a date column sorts chronologically, not by the raw Date string", async ({ page }) => {
  await gotoFixture(page);

  const applied = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sort!.sheet;
    return sheet.sortByColumn(2, "asc");
  });
  expect(applied).toBe(true);

  const column = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sort!.sheet;
    return [0, 1, 2].map((row) => sheet.getCell(row, 0) as string);
  });

  // Chronological is Jan 1 / Jan 6 / Jan 11 = cherry, apple, banana. Ordering
  // the raw Date values as strings would give their weekday names, Mon < Sat
  // < Wed = apple, banana, cherry — so this also proves the column's cell
  // type `compare` really does reach the sort at runtime.
  expect(column).toEqual(["cherry", "apple", "banana"]);
});

test("one undo restores both the row order and the row heights", async ({ page }) => {
  await gotoFixture(page);

  const before = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sort!.sheet;
    sheet.setRowHeight(0, 60);
    return {
      column: [0, 1, 2].map((row) => sheet.getCell(row, 0) as string),
      heights: [0, 1, 2].map((row) => sheet.scene.getRowHeight(row))
    };
  });
  expect(before.heights).toEqual([60, 20, 20]);

  await page.evaluate(() => {
    (window as unknown as FixtureWindow).__sort!.sheet.sortByColumn(0, "asc");
  });
  await page.evaluate(() => {
    (window as unknown as FixtureWindow).__sort!.sheet.undo();
  });

  const after = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sort!.sheet;
    return {
      column: [0, 1, 2].map((row) => sheet.getCell(row, 0) as string),
      heights: [0, 1, 2].map((row) => sheet.scene.getRowHeight(row))
    };
  });

  expect(after).toEqual(before);
});

test("a column comparator written as a method is called on its column", async ({ page }) => {
  await gotoFixture(page);

  const applied = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sort!.sheet;
    return sheet.sortByColumn(3, "asc");
  });
  expect(applied).toBe(true);

  const column = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sort!.sheet;
    return [0, 1, 2].map((row) => sheet.getCell(row, 0) as string);
  });

  // The "priority" comparator is an object method that reads `this.options`
  // to rank low < medium < high. Detached from its column it throws; ordered
  // alphabetically it would give high < low < medium = banana, cherry, apple.
  expect(column).toEqual(["cherry", "apple", "banana"]);
});

test("an already-aborted signal returns false and writes nothing", async ({ page }) => {
  await gotoFixture(page);

  const result = await page.evaluate(() => {
    const fixture = (window as unknown as FixtureWindow).__sort!;
    const sheet = fixture.sheet;
    const controller = new AbortController();
    controller.abort();
    const columnBefore = [0, 1, 2].map((row) => sheet.getCell(row, 0) as string);
    const commitsBefore = fixture.getCommitCount();

    const applied = sheet.sortByColumn(0, "asc", { signal: controller.signal });

    // Read before any undo(): an implementation that applied the sort *and*
    // returned false is exactly what this test exists to rule out, and an
    // undo() first would hide it.
    return {
      applied,
      columnBefore,
      columnAfter: [0, 1, 2].map((row) => sheet.getCell(row, 0) as string),
      commits: fixture.getCommitCount() - commitsBefore
    };
  });

  expect(result.applied).toBe(false);
  expect(result.columnBefore).toEqual(["banana", "cherry", "apple"]);
  expect(result.columnAfter).toEqual(result.columnBefore);
  // Nothing was committed, so nothing went onto the history stack either.
  expect(result.commits).toBe(0);
});

test("aborting from inside the comparator returns false and writes nothing", async ({ page }) => {
  await gotoFixture(page);

  const result = await page.evaluate(() => {
    const fixture = (window as unknown as FixtureWindow).__sort!;
    const sheet = fixture.sheet;
    const controller = new AbortController();
    const columnBefore = [0, 1, 2].map((row) => sheet.getCell(row, 0) as string);
    const commitsBefore = fixture.getCommitCount();

    // Aborts partway through the sort itself, after the keys have been read
    // — not before sortByColumn() was entered.
    let comparisons = 0;
    fixture.setOnCompare(() => {
      comparisons += 1;
      controller.abort();
    });
    let applied: boolean;
    try {
      applied = sheet.sortByColumn(3, "asc", { signal: controller.signal });
    } finally {
      fixture.setOnCompare(null);
    }

    return {
      applied,
      comparisons,
      columnBefore,
      columnAfter: [0, 1, 2].map((row) => sheet.getCell(row, 0) as string),
      commits: fixture.getCommitCount() - commitsBefore
    };
  });

  // The abort really did happen mid-computation, not before entry.
  expect(result.comparisons).toBeGreaterThan(0);
  expect(result.applied).toBe(false);
  expect(result.columnAfter).toEqual(result.columnBefore);
  expect(result.commits).toBe(0);
});

async function refusedSort(page: Page): Promise<{
  code: string | undefined;
  columnBefore: string[];
  columnAfter: string[];
}> {
  return page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sort!.sheet;
    const columnBefore = [0, 1, 2].map((row) => sheet.getCell(row, 0) as string);
    let code: string | undefined;
    try {
      sheet.sortByColumn(0, "asc");
    } catch (error) {
      code = (error as { code?: string }).code;
    }
    return {
      code,
      columnBefore,
      columnAfter: [0, 1, 2].map((row) => sheet.getCell(row, 0) as string)
    };
  });
}

test("sorting is rejected on a DataSource that cannot reorder", async ({ page }) => {
  await gotoFixture(page, "?dataSource=no-reorder");

  const result = await refusedSort(page);

  expect(result.code).toBe("structural-modification-not-supported");
  expect(result.columnBefore).toEqual(["banana", "cherry", "apple"]);
  expect(result.columnAfter).toEqual(result.columnBefore);
});

test("sorting is rejected on a DataSource that cannot transact at all", async ({ page }) => {
  await gotoFixture(page, "?dataSource=minimal");

  const result = await refusedSort(page);

  // A DataSource with only the four required methods never gets as far as
  // the reorder capability check: it cannot open a transaction at all.
  expect(result.code).toBe("mutation-not-supported");
  expect(result.columnAfter).toEqual(result.columnBefore);
});
