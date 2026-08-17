import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface FixtureWindow {
  __cyexcelAutoHeightPreflightReady?: boolean;
  __autoHeightPreflight?: {
    readonly sheet: {
      readonly scene: {
        getRowHeight(row: number): number;
        getVisibleRows(): { readonly start: number; readonly count: number };
      };
      getCell(row: number, column: number): unknown;
      deleteColumns(start: number, count?: number): void;
      deleteRows(start: number, count?: number): void;
      setColumnWidth(column: number, width: number): void;
    };
    readonly rowLayoutRows: number[];
    readonly rowLayoutBatches: Array<{ readonly frameId: number; readonly rowCount: number }>;
    readonly errors: Array<{ readonly eventName: string; readonly message: string }>;
    resetBatches(): void;
    scrollBy(delta: number): void;
  };
}

async function gotoScenario(page: Page, scenario: string): Promise<void> {
  await page.goto(`/tests/browser/auto-height-preflight.html?scenario=${scenario}`);
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelAutoHeightPreflightReady === true
  );
  await settlePublications(page);
}

async function settlePublications(page: Page): Promise<void> {
  await page.evaluate(
    () => new Promise<void>((resolve) =>
      requestAnimationFrame(() => requestAnimationFrame(() => requestAnimationFrame(() => resolve())))
    )
  );
}

test("deleting the only wrap column returns auto-owned rows to the default height", async ({ page }) => {
  await gotoScenario(page, "delete-wrap-column");

  const before = await page.evaluate(() => {
    const fixture = (window as unknown as FixtureWindow).__autoHeightPreflight;
    if (!fixture) throw new Error("missing auto-height preflight fixture");
    return fixture.sheet.scene.getRowHeight(0);
  });
  expect(before).toBeGreaterThan(24);

  await page.evaluate(() => {
    const fixture = (window as unknown as FixtureWindow).__autoHeightPreflight;
    if (!fixture) throw new Error("missing auto-height preflight fixture");
    fixture.sheet.deleteColumns(0, 1);
  });
  await settlePublications(page);

  const result = await page.evaluate(() => {
    const fixture = (window as unknown as FixtureWindow).__autoHeightPreflight;
    if (!fixture) throw new Error("missing auto-height preflight fixture");
    return {
      height: fixture.sheet.scene.getRowHeight(0),
      rowLayoutRows: [...fixture.rowLayoutRows],
      errors: [...fixture.errors]
    };
  });
  expect(result).toEqual({ height: 24, rowLayoutRows: [0], errors: [] });
});

test("deleting a top row measures the long row newly entering the visible pane", async ({ page }) => {
  await gotoScenario(page, "delete-top-row");

  const before = await page.evaluate(() => {
    const fixture = (window as unknown as FixtureWindow).__autoHeightPreflight;
    if (!fixture) throw new Error("missing auto-height preflight fixture");
    return {
      visible: fixture.sheet.scene.getVisibleRows(),
      offscreenHeight: fixture.sheet.scene.getRowHeight(4)
    };
  });
  expect(before.visible).toEqual({ start: 0, count: 4 });
  expect(before.offscreenHeight).toBe(24);

  await page.evaluate(() => {
    const fixture = (window as unknown as FixtureWindow).__autoHeightPreflight;
    if (!fixture) throw new Error("missing auto-height preflight fixture");
    fixture.sheet.deleteRows(0, 1);
  });
  await settlePublications(page);

  const result = await page.evaluate(() => {
    const fixture = (window as unknown as FixtureWindow).__autoHeightPreflight;
    if (!fixture) throw new Error("missing auto-height preflight fixture");
    return {
      movedValue: fixture.sheet.getCell(3, 0),
      height: fixture.sheet.scene.getRowHeight(3),
      visible: fixture.sheet.scene.getVisibleRows(),
      rowLayoutRows: [...fixture.rowLayoutRows],
      errors: [...fixture.errors]
    };
  });
  expect(result.movedValue).toBe(
    "newly visible row must be measured before its final scene is published"
  );
  expect(result.height).toBeGreaterThan(24);
  expect(result.visible.start).toBe(0);
  expect(result.rowLayoutRows).toContain(3);
  expect(result.errors).toEqual([]);
});

test("widening a column runs a second measurement pass for a newly revealed row", async ({ page }) => {
  await gotoScenario(page, "widen-column");

  const before = await page.evaluate(() => {
    const fixture = (window as unknown as FixtureWindow).__autoHeightPreflight;
    if (!fixture) throw new Error("missing auto-height preflight fixture");
    return {
      row0: fixture.sheet.scene.getRowHeight(0),
      row1: fixture.sheet.scene.getRowHeight(1),
      visible: fixture.sheet.scene.getVisibleRows()
    };
  });
  expect(before.row0).toBeGreaterThan(95);
  expect(before.row1).toBeGreaterThan(24);
  expect(before.visible).toEqual({ start: 0, count: 1 });

  await page.evaluate(() => {
    const fixture = (window as unknown as FixtureWindow).__autoHeightPreflight;
    if (!fixture) throw new Error("missing auto-height preflight fixture");
    fixture.sheet.setColumnWidth(0, 220);
  });
  await settlePublications(page);

  const result = await page.evaluate(() => {
    const fixture = (window as unknown as FixtureWindow).__autoHeightPreflight;
    if (!fixture) throw new Error("missing auto-height preflight fixture");
    return {
      row0: fixture.sheet.scene.getRowHeight(0),
      row1: fixture.sheet.scene.getRowHeight(1),
      visible: fixture.sheet.scene.getVisibleRows(),
      rowLayoutRows: [...fixture.rowLayoutRows],
      errors: [...fixture.errors]
    };
  });
  expect(result.row0).toBeLessThan(before.row0);
  expect(result.visible.count).toBeGreaterThan(1);
  expect(result.row1).toBeGreaterThan(24);
  expect(result.row1).toBeLessThan(before.row1);
  expect(result.rowLayoutRows).toContain(1);
  expect(result.errors).toEqual([]);
});

/**
 * Migration plan §10's remaining Phase 7 gate: "自动行高一帧最多提交一次 model
 * row-height batch，不逐行触发全套 overlay/render side effects."
 *
 * Batching alone does not satisfy it — N single-row batches spread over N
 * frames and N batches inside one frame produce an identical patch stream, so
 * the fixture stamps every row-layout patch with the animation frame it landed
 * in and the assertions are made per frame.
 */
test("auto-height commits at most one row-height batch per animation frame", async ({ page }) => {
  await gotoScenario(page, "batch-per-frame");
  await page.evaluate(() => {
    (window as unknown as FixtureWindow).__autoHeightPreflight!.resetBatches();
  });

  // Each scroll reveals a screenful of never-measured wrap rows at once.
  for (let step = 0; step < 6; step += 1) {
    await page.evaluate(() => {
      (window as unknown as FixtureWindow).__autoHeightPreflight!.scrollBy(90);
    });
    await settlePublications(page);
  }

  const batches = await page.evaluate(
    () => (window as unknown as FixtureWindow).__autoHeightPreflight!.rowLayoutBatches
  );
  expect(batches.length).toBeGreaterThan(0);

  const perFrame = new Map<number, number>();
  for (const batch of batches) {
    perFrame.set(batch.frameId, (perFrame.get(batch.frameId) ?? 0) + 1);
  }
  expect(Math.max(...perFrame.values())).toBe(1);
});

test("one auto-height pass resizes many rows in a single batch, not one batch per row", async ({
  page
}) => {
  await gotoScenario(page, "batch-per-frame");
  await page.evaluate(() => {
    const fixture = (window as unknown as FixtureWindow).__autoHeightPreflight!;
    fixture.resetBatches();
    fixture.scrollBy(400);
  });
  await settlePublications(page);

  const batches = await page.evaluate(
    () => (window as unknown as FixtureWindow).__autoHeightPreflight!.rowLayoutBatches
  );
  expect(batches.length).toBeGreaterThan(0);
  // Landing a screenful of unmeasured rows must produce a multi-row batch; one
  // row per batch would mean exactly the per-row side effects this gate forbids.
  expect(Math.max(...batches.map((batch) => batch.rowCount))).toBeGreaterThan(1);
});
