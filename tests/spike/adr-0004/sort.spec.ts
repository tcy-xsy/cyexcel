import { test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface SortResult {
  readonly shape: string;
  readonly orderMs: number;
  readonly applyMs: number;
  readonly totalMs: number;
  readonly longTaskCount: number;
  readonly firstValue: unknown;
  readonly lastValue: unknown;
}

interface FixtureWindow {
  __spikeReady?: boolean;
  __spikeSortDone?: boolean;
  __spikeSort?: SortResult;
}

// The sort must be triggered by a real page.click(), never page.evaluate().
// Chromium's Long Tasks API does not attribute script executed via CDP's
// Runtime.evaluate (what page.evaluate() uses under the hood) to a task the
// way it does script run from a genuine DOM event — an identical 800ms+
// busy sort measured longTaskCount 0 on every run under page.evaluate() and
// was only correctly captured once triggered by page.click(). This is
// documented at the top of tests/performance/sort.spec.ts.
async function measure(page: Page, shape: string): Promise<SortResult> {
  await page.goto(`/tests/spike/adr-0004/sort.html?shape=${shape}`);
  await page.waitForFunction(() => (window as unknown as FixtureWindow).__spikeReady === true, undefined, {
    timeout: 180_000
  });
  await page.click("#sort-trigger");
  await page.waitForFunction(() => (window as unknown as FixtureWindow).__spikeSortDone === true, undefined, {
    timeout: 180_000
  });
  const result = await page.evaluate(() => (window as unknown as FixtureWindow).__spikeSort);
  if (!result) {
    throw new Error(`sort fixture returned nothing for shape=${shape}`);
  }
  return result;
}

for (const shape of ["map", "object"] as const) {
  test(`ADR 0004 spike — sort 100k x 50 by row-bucket reorder (${shape})`, async ({ page }) => {
    test.setTimeout(300_000);
    const result = await measure(page, shape);
    console.log(`[spike:sort] ${JSON.stringify(result)}`);
    if (typeof result.firstValue !== "number" || typeof result.lastValue !== "number") {
      throw new Error("sort did not produce numeric endpoints — the fixture sorted the wrong column");
    }
    if (result.firstValue > result.lastValue) {
      throw new Error("sort produced a descending order — the comparator is wrong");
    }
  });
}
