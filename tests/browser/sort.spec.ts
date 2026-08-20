import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface SortState {
  readonly column: number;
  readonly direction: "asc" | "desc";
}

interface SortSheet {
  getCell(row: number, column: number): unknown;
  setRowHeight(row: number, height: number): void;
  sortByColumn(
    column: number,
    direction: "asc" | "desc",
    options?: { signal?: AbortSignal }
  ): boolean;
  undo(): void;
  insertColumns(start: number, columns: readonly Record<string, unknown>[]): void;
  deleteColumns(start: number, count?: number): void;
  readonly scene: {
    getRowHeight(row: number): number;
  };
  readonly sortState: SortState | null;
  readonly events: {
    on(name: string, handler: (payload: unknown) => void): void;
    off(name: string, handler: (payload: unknown) => void): void;
  };
}

interface FixtureWindow {
  __sortFixtureReady?: boolean;
  __sort?: {
    sheet: SortSheet;
    getCommitCount(): number;
    getErrors(): { eventName: string; code?: string; message: string }[];
    getHeaderActions(): { action: { type: string }; column: number }[];
    setOnCompare(hook: (() => void) | null): void;
  };
}

async function gotoFixture(page: Page, query = ""): Promise<void> {
  await page.goto(`/tests/browser/sort.html${query}`);
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__sortFixtureReady === true
  );
}

/** The column-header canvas box: sort.html has no header groups (flat single leaf row), so canvas index 2 is always the column header — confirmed against the fixture's own layer order (0=content,1=selection,2=column-header,3=row-header). */
function headerCanvasBox(page: Page) {
  return page.locator("#app canvas").nth(2).boundingBox();
}

/**
 * Clicks inside the reserved click-to-sort strip at the trailing edge of a
 * *default*-painted header cell (no custom renderer configures its own
 * hitTest) — the strip is `DEFAULT_HEADER_SORT_ZONE_WIDTH` (18px) wide,
 * anchored to the column's own right edge, so `columnRightX - 8` always
 * lands inside it regardless of column width.
 */
async function clickHeaderSortZone(page: Page, columnRightX: number): Promise<void> {
  const headerBox = await headerCanvasBox(page);
  if (!headerBox) throw new Error("missing header canvas box");
  await page.mouse.click(headerBox.x + columnRightX - 8, headerBox.y + 10);
}

// Column x-ranges in sort.html: name [0,120), score [120,200), joined [200,310), priority [310,400).
const NAME_COLUMN_RIGHT_X = 120;

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

test("clicking a default header's sort zone sorts ascending and reports it in sortState", async ({
  page
}) => {
  await gotoFixture(page);

  await clickHeaderSortZone(page, NAME_COLUMN_RIGHT_X);

  const result = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sort!.sheet;
    return {
      column: [0, 1, 2].map((row) => sheet.getCell(row, 0) as string),
      sortState: sheet.sortState
    };
  });

  expect(result.column).toEqual(["apple", "banana", "cherry"]);
  expect(result.sortState).toEqual({ column: 0, direction: "asc" });
});

test("clicking the same header's sort zone twice sorts descending", async ({ page }) => {
  await gotoFixture(page);

  await clickHeaderSortZone(page, NAME_COLUMN_RIGHT_X);
  await clickHeaderSortZone(page, NAME_COLUMN_RIGHT_X);

  const result = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sort!.sheet;
    return {
      column: [0, 1, 2].map((row) => sheet.getCell(row, 0) as string),
      sortState: sheet.sortState
    };
  });

  expect(result.column).toEqual(["cherry", "banana", "apple"]);
  expect(result.sortState).toEqual({ column: 0, direction: "desc" });
});

test("a third click is still a valid sort, never a return to the original row order", async ({
  page
}) => {
  await gotoFixture(page);
  const originalOrder = ["banana", "cherry", "apple"];

  await clickHeaderSortZone(page, NAME_COLUMN_RIGHT_X);
  await clickHeaderSortZone(page, NAME_COLUMN_RIGHT_X);
  await clickHeaderSortZone(page, NAME_COLUMN_RIGHT_X);

  const result = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sort!.sheet;
    return {
      column: [0, 1, 2].map((row) => sheet.getCell(row, 0) as string),
      sortState: sheet.sortState
    };
  });

  // The third click cycles back to ascending, *not* to an "unsorted" state
  // — there is no third indicator value to land on (ADR 0003 §3.4), and the
  // data itself is never the pre-sort insertion order again.
  expect(result.column).toEqual(["apple", "banana", "cherry"]);
  expect(result.column).not.toEqual(originalOrder);
  expect(result.sortState).toEqual({ column: 0, direction: "asc" });
});

test("a header click whose sort is refused leaves the indicator and the data exactly where they were", async ({
  page
}) => {
  // Driven by a real click, never page.evaluate(): a click reaches
  // #routeHeaderCanvasAction from inside an interaction read lease, so the
  // sort runs *deferred*, out of the mutation queue's own drain. That is the
  // only path a real user can take, and it is the path on which a failing
  // sort has no call site left to throw at — the two page.evaluate()-driven
  // refusal tests above run the sort synchronously and so never exercise it.
  await gotoFixture(page, "?dataSource=no-reorder");

  await clickHeaderSortZone(page, NAME_COLUMN_RIGHT_X);

  const result = await page.evaluate(() => {
    const fixture = (window as unknown as FixtureWindow).__sort!;
    return {
      column: [0, 1, 2].map((row) => fixture.sheet.getCell(row, 0) as string),
      sortState: fixture.sheet.sortState,
      errors: fixture.getErrors(),
      headerActions: fixture.getHeaderActions()
    };
  });

  expect(result.column).toEqual(["banana", "cherry", "apple"]);
  // The indicator must still describe the last sort that really happened —
  // here, none at all.
  expect(result.sortState).toBeNull();
  expect(result.errors).toMatchObject([
    { eventName: "interaction:toggle-sort:0", code: "structural-modification-not-supported" }
  ]);
  // ...and no `header-action` either: for a default header the core owns the
  // sort, so announcing `toggle-sort` after refusing to sort would tell every
  // listener a sort just happened. The `error` event is the whole report.
  expect(result.headerActions).toEqual([]);
});

test("a custom header renderer's toggle-sort stays a notification -- the core sorts nothing", async ({
  page
}) => {
  await gotoFixture(page, "?header=custom");

  // The custom renderer's hitTest reports toggle-sort anywhere in the cell,
  // so a click at the column midpoint (not in the default painter's reserved
  // strip, which this column does not have) is enough.
  const headerBox = await headerCanvasBox(page);
  if (!headerBox) throw new Error("missing header canvas box");
  await page.mouse.click(headerBox.x + NAME_COLUMN_RIGHT_X / 2, headerBox.y + 10);

  const result = await page.evaluate(() => {
    const fixture = (window as unknown as FixtureWindow).__sort!;
    return {
      column: [0, 1, 2].map((row) => fixture.sheet.getCell(row, 0) as string),
      sortState: fixture.sheet.sortState,
      commits: fixture.getCommitCount(),
      headerActions: fixture.getHeaderActions()
    };
  });

  // Before this branch `toggle-sort` was purely a notification a host
  // translated into its own command; a custom renderer must keep that
  // contract, or every host following doc/customization-and-hooks.md §7
  // (which tells it to dispatch its own sort from the listener) gets two
  // physical reorders and two history entries per click.
  expect(result.column).toEqual(["banana", "cherry", "apple"]);
  expect(result.sortState).toBeNull();
  expect(result.commits).toBe(0);
  expect(result.headerActions).toEqual([{ action: { type: "toggle-sort" }, column: 0 }]);
});

test("a structural column change clears the sort indicator", async ({ page }) => {
  await gotoFixture(page);

  const afterInsert = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sort!.sheet;
    sheet.sortByColumn(1, "asc");
    const sorted = sheet.sortState;
    sheet.insertColumns(0, [{ id: "extra", field: "extra", width: 50, header: "Extra" }]);
    return { sorted, after: sheet.sortState };
  });

  expect(afterInsert.sorted).toEqual({ column: 1, direction: "asc" });
  // An inserted column shifts every column index right; keeping `column: 1`
  // would draw the indicator on a different column than the one that was
  // sorted. ADR 0003 §3.4's indicator is one-shot ("the last sort applied"),
  // and after the column axis moves that claim is no longer safe to make.
  expect(afterInsert.after).toBeNull();

  const afterDelete = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sort!.sheet;
    sheet.sortByColumn(2, "desc");
    const sorted = sheet.sortState;
    sheet.deleteColumns(2, 1);
    return { sorted, after: sheet.sortState };
  });

  expect(afterDelete.sorted).toEqual({ column: 2, direction: "desc" });
  // The sorted column is gone: a `sortState` naming it would hand callers an
  // index for a column that no longer exists, and would start the *new*
  // column 2 at "desc" on its first click.
  expect(afterDelete.after).toBeNull();
});

test("clicking a header's sort zone emits a toggle-sort header-action event", async ({ page }) => {
  await gotoFixture(page);

  await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sort!.sheet;
    (window as unknown as { __actions: unknown[] }).__actions = [];
    sheet.events.on("header-action", (payload) =>
      (window as unknown as { __actions: unknown[] }).__actions.push(payload)
    );
  });

  await clickHeaderSortZone(page, NAME_COLUMN_RIGHT_X);

  const actions = await page.evaluate(
    () => (window as unknown as { __actions: unknown[] }).__actions
  );

  expect(actions).toMatchObject([
    {
      action: { type: "toggle-sort" },
      column: 0,
      header: { column: 0, title: "Name" }
    }
  ]);
});

test("a direct sortByColumn() call updates sortState the same way a header click does", async ({
  page
}) => {
  await gotoFixture(page);

  const sortState = await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sort!.sheet;
    sheet.sortByColumn(1, "desc");
    return sheet.sortState;
  });

  expect(sortState).toEqual({ column: 1, direction: "desc" });
});

test("a header-action listener reading sortState synchronously inside the callback sees the sort this click just produced, not the previous one", async ({
  page
}) => {
  await gotoFixture(page);

  await page.evaluate(() => {
    const sheet = (window as unknown as FixtureWindow).__sort!.sheet;
    (window as unknown as { __captured: unknown[] }).__captured = [];
    sheet.events.on("header-action", (payload) => {
      const action = payload as { action: { type: string } };
      if (action.action.type !== "toggle-sort") {
        return;
      }
      // Read synchronously, *inside* the listener — this is exactly the
      // pattern that exposed the header-action emit-ordering bug fixed in
      // commit 960c1f0 (#routeHeaderCanvasAction used to emit before
      // applying the sort). Every other test in this file reads sortState
      // only after the click resolves, via a separate page.evaluate() —
      // that reads too late to catch a listener seeing stale state, which
      // is exactly how this bug survived a fully green suite until it was
      // found by hand in the playground.
      //
      // The cell is read alongside it, and that is the half that matters
      // most: when this test was first written the header click applied its
      // sort by calling `sortByColumn` from inside the interaction intent,
      // which only *queued* the reorder behind the drain already running —
      // so `sortState` was current here while the data underneath it still
      // was not. The core now runs the reorder inline in that intent, so
      // both are true together, which is the only version of this guarantee
      // worth having.
      (window as unknown as { __captured: unknown[] }).__captured.push({
        sortState: sheet.sortState,
        topCell: sheet.getCell(0, 0)
      });
    });
  });

  await clickHeaderSortZone(page, NAME_COLUMN_RIGHT_X);

  const captured = await page.evaluate(
    () => (window as unknown as { __captured: unknown[] }).__captured
  );

  // Pre-fix (emit before apply), the listener would run before the sort had
  // written #sortState: sortState would still be null. Pre-*this*-fix,
  // sortState would already read `asc` while `getCell(0, 0)` still returned
  // "banana" — the indicator describing a sort that had not happened.
  expect(captured).toEqual([
    { sortState: { column: 0, direction: "asc" }, topCell: "apple" }
  ]);
});
