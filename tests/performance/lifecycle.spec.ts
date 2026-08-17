import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

import { assertPerfMetric } from "./support/thresholds";

interface LifecycleResult {
  readonly heapSamples: ReadonlyArray<number | null>;
  readonly domNodeCountAfter: number;
}

interface FixtureWindow {
  __cyexcelPerfReady?: boolean;
  __cyexcelRunLifecycle?: (iterations: number) => Promise<LifecycleResult>;
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/performance/lifecycle.html");
  await page.waitForFunction(() => (window as unknown as FixtureWindow).__cyexcelPerfReady === true, undefined, {
    timeout: 30_000
  });
}

function average(values: readonly number[]): number {
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

test("100x construct+dispose: no heap growth, no lingering Worker, container fully cleaned up", async ({ page }) => {
  test.setTimeout(120_000);
  await gotoFixture(page);

  const result = await page.evaluate(
    (iterations) => (window as unknown as FixtureWindow).__cyexcelRunLifecycle?.(iterations),
    100
  );
  if (!result) {
    throw new Error("lifecycle fixture did not return a result");
  }

  const heapValues = result.heapSamples.filter((value): value is number => value !== null);
  if (heapValues.length >= 4) {
    const quarter = Math.max(1, Math.floor(heapValues.length / 4));
    const heapGrowthBytes = average(heapValues.slice(-quarter)) - average(heapValues.slice(0, quarter));
    // Same GC-timing-noise floor as dom-renderer-scroll.spec.ts — see that
    // file's comment.
    assertPerfMetric("lifecycle", "heapGrowthBytes", Math.max(0, heapGrowthBytes), { toleranceFloor: 5_000_000 });
  }

  expect(result.domNodeCountAfter, "container must be empty after the last dispose()").toBe(0);
  expect(page.workers().length, "no lingering dedicated Worker after dispose()").toBe(0);
});
