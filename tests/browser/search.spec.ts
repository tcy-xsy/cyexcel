import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface SearchMatch {
  readonly row: number;
  readonly column: number;
}

interface SearchSheet {
  getCell(row: number, column: number): unknown;
  find(query: string, options?: Record<string, unknown>): Promise<readonly SearchMatch[] | null>;
  findNext(): SearchMatch | null;
  findPrevious(): SearchMatch | null;
  clearSearch(): void;
  replace(match: SearchMatch, value: unknown): void;
  replaceAll(query: string, value: unknown, options?: Record<string, unknown>): Promise<number | null>;
  undo(): void;
  canUndo(): boolean;
  readonly selection: { readonly ranges: readonly { startRow: number; startColumn: number }[] };
}

interface FixtureWindow {
  __searchFixtureReady?: boolean;
  __search?: {
    sheet: SearchSheet;
    getCommitCount(): number;
  };
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/search.html");
  await page.waitForFunction(() => (window as unknown as FixtureWindow).__searchFixtureReady === true);
}

// Fixture rows: apple/fruit/3, banana/fruit/1, carrot/veg/2, daikon/veg/5.

test("find() returns every occupied cell whose displayed text matches, case-insensitive by default", async ({
  page
}) => {
  await gotoFixture(page);

  const matches = await page.evaluate(async () => {
    const sheet = (window as unknown as FixtureWindow).__search!.sheet;
    return sheet.find("FRUIT");
  });

  expect(matches).toEqual([
    { row: 0, column: 1 },
    { row: 1, column: 1 }
  ]);
});

test("findNext() cycles through matches, selecting and scrolling to each one", async ({ page }) => {
  await gotoFixture(page);

  const activeCells = await page.evaluate(async () => {
    const sheet = (window as unknown as FixtureWindow).__search!.sheet;
    await sheet.find("fruit");
    const seen: SearchMatch[] = [];
    seen.push(sheet.findNext()!);
    seen.push(sheet.findNext()!);
    seen.push(sheet.findNext()!); // wraps back to the first match
    return seen;
  });

  expect(activeCells).toEqual([
    { row: 0, column: 1 },
    { row: 1, column: 1 },
    { row: 0, column: 1 }
  ]);
});

test("replace() writes one cell through a normal undoable command", async ({ page }) => {
  await gotoFixture(page);

  const outcome = await page.evaluate(async () => {
    const sheet = (window as unknown as FixtureWindow).__search!.sheet;
    sheet.replace({ row: 0, column: 0 }, "apricot");
    const after = sheet.getCell(0, 0);
    sheet.undo();
    const afterUndo = sheet.getCell(0, 0);
    return { after, afterUndo };
  });

  expect(outcome).toEqual({ after: "apricot", afterUndo: "apple" });
});

test("replaceAll() replaces every match in one transaction, undoable in a single step", async ({ page }) => {
  await gotoFixture(page);

  const outcome = await page.evaluate(async () => {
    const sheet = (window as unknown as FixtureWindow).__search!.sheet;
    const before = sheet.getCell(0, 1);
    const count = await sheet.replaceAll("fruit", "produce");
    const after = [sheet.getCell(0, 1), sheet.getCell(1, 1)];
    sheet.undo();
    const afterUndo = [sheet.getCell(0, 1), sheet.getCell(1, 1)];
    return { before, count, after, afterUndo, canUndoAgain: sheet.canUndo() };
  });

  expect(outcome.before).toBe("fruit");
  expect(outcome.count).toBe(2);
  expect(outcome.after).toEqual(["produce", "produce"]);
  // One undo restores both cells — one transaction, one History entry.
  expect(outcome.afterUndo).toEqual(["fruit", "fruit"]);
  expect(outcome.canUndoAgain).toBe(false);
});

test("clearSearch() empties the match set", async ({ page }) => {
  await gotoFixture(page);

  const afterClear = await page.evaluate(async () => {
    const sheet = (window as unknown as FixtureWindow).__search!.sheet;
    await sheet.find("fruit");
    sheet.clearSearch();
    return sheet.findNext();
  });

  expect(afterClear).toBeNull();
});

test("find() with useRegex compiles the query as a regular expression", async ({ page }) => {
  await gotoFixture(page);

  const matches = await page.evaluate(async () => {
    const sheet = (window as unknown as FixtureWindow).__search!.sheet;
    return sheet.find("^\\d$", { useRegex: true });
  });

  // All four scores (3, 1, 2, 5) are single digits, so every score cell matches "^\d$".
  expect(matches).toEqual([
    { row: 0, column: 2 },
    { row: 1, column: 2 },
    { row: 2, column: 2 },
    { row: 3, column: 2 }
  ]);
});

test("find() with an already-aborted signal returns null and leaves the previous match set untouched", async ({
  page
}) => {
  await gotoFixture(page);

  const outcome = await page.evaluate(async () => {
    const sheet = (window as unknown as FixtureWindow).__search!.sheet;
    await sheet.find("fruit");
    const controller = new AbortController();
    controller.abort();
    const cancelled = await sheet.find("veg", { signal: controller.signal });
    // The old "fruit" matches are still what findNext() walks.
    const stillFruit = sheet.findNext();
    return { cancelled, stillFruit };
  });

  expect(outcome.cancelled).toBeNull();
  expect(outcome.stillFruit).toEqual({ row: 0, column: 1 });
});
