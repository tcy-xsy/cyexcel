import { test } from "@playwright/test";
import type { Page } from "@playwright/test";

import { assertPerfMetric } from "./support/thresholds";

interface SortResult {
  readonly applied: boolean;
  readonly sortMs: number;
  readonly repaintMs: number;
  readonly longTaskCount: number;
}

interface FixtureWindow {
  __cyexcelPerfReady?: boolean;
  __cyexcelSortDone?: boolean;
  __cyexcelSortResult?: SortResult;
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/performance/sort.html");
  await page.waitForFunction(() => (window as unknown as FixtureWindow).__cyexcelPerfReady === true, undefined, {
    timeout: 90_000
  });
}

// Task 7's computeSortOrder deliberately left the sort itself unchunked,
// pending a real measurement here (see its "分片说明" comment in
// .superpowers/sdd/step22-sort-plan/task-7-brief.md): "不要提前实现没有实测
// 支撑的分片". This is that measurement — whether sorting 1,000,000 rows
// produces a long task on the main thread.
//
// The sort MUST be triggered by a genuine page.click(), not page.evaluate():
// Chromium's Long Tasks API (PerformanceObserver({entryTypes:["longtask"]}))
// does not attribute script executed via CDP's Runtime.evaluate (what
// page.evaluate() uses under the hood) to a task the way it does script run
// from a real DOM event. A page.evaluate()-triggered version of this same
// sort measured longTaskCount 0 on every run despite >800ms of synchronous
// blocking; an identical busy-loop triggered by page.click() was correctly
// captured. Only a click makes this measurement trustworthy.
test("sorting 1,000,000 rows by a numeric column completes and reports its main-thread cost", async ({
  page
}) => {
  test.setTimeout(120_000);
  await gotoFixture(page);

  await page.click("#sort-trigger");
  await page.waitForFunction(() => (window as unknown as FixtureWindow).__cyexcelSortDone === true, undefined, {
    timeout: 90_000
  });
  const result = await page.evaluate(() => (window as unknown as FixtureWindow).__cyexcelSortResult);
  if (!result) {
    throw new Error("sort fixture did not return a result");
  }
  if (!result.applied) {
    throw new Error("sortByColumn reported it did not apply");
  }

  assertPerfMetric("sort", "sortMs", result.sortMs, { toleranceFloor: 50 });
  assertPerfMetric("sort", "repaintMs", result.repaintMs, { toleranceFloor: 20 });
  assertPerfMetric("sort", "longTaskCount", result.longTaskCount);
});
