import { test } from "@playwright/test";
import type { Page } from "@playwright/test";

import { assertPerfMetric, SCROLL_DURATION_MS } from "./support/thresholds";

interface Sample {
  readonly tMs: number;
  readonly domNodeCount: number;
  readonly heapBytes: number | null;
}

interface FixtureWindow {
  __cyexcelPerfReady?: boolean;
  __cyexcelRunScrollWithDomTrend?: (durationMs: number, sampleIntervalMs: number) => Promise<Sample[]>;
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/performance/dom-renderer-scroll.html");
  await page.waitForFunction(() => (window as unknown as FixtureWindow).__cyexcelPerfReady === true, undefined, {
    timeout: 90_000
  });
}

function average(values: readonly number[]): number {
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

test("DOM-renderer columns stay pooled over a long continuous scroll — no DOM/heap growth", async ({ page }) => {
  test.setTimeout(SCROLL_DURATION_MS + 60_000);
  await gotoFixture(page);

  const samples = await page.evaluate(
    ({ durationMs, sampleIntervalMs }) =>
      (window as unknown as FixtureWindow).__cyexcelRunScrollWithDomTrend?.(durationMs, sampleIntervalMs),
    { durationMs: SCROLL_DURATION_MS, sampleIntervalMs: 1000 }
  );
  if (!samples || samples.length < 4) {
    throw new Error("dom-renderer-scroll fixture did not return enough samples");
  }

  // Compare the average of the first vs. last quarter of samples rather
  // than single endpoints — GC timing is noisy sample-to-sample, but a real
  // leak shows up as a trend across the whole run.
  const quarterSize = Math.max(1, Math.floor(samples.length / 4));
  const firstQuarter = samples.slice(0, quarterSize);
  const lastQuarter = samples.slice(-quarterSize);

  const domNodeGrowth = average(lastQuarter.map((s) => s.domNodeCount)) - average(firstQuarter.map((s) => s.domNodeCount));
  assertPerfMetric("domRendererScroll", "domNodeGrowth", Math.max(0, domNodeGrowth));

  const heapSamples = samples.filter((s) => s.heapBytes !== null).map((s) => s.heapBytes as number);
  if (heapSamples.length >= 4) {
    const heapQuarter = Math.max(1, Math.floor(heapSamples.length / 4));
    const heapGrowthBytes = average(heapSamples.slice(-heapQuarter)) - average(heapSamples.slice(0, heapQuarter));
    // performance.memory is a coarse, GC-timing-dependent sample — a real
    // 5-minute run reads a few hundred KB of apparent growth purely from
    // sampling noise even though domNodeGrowth (the structural signal that
    // actually indicates a pool leak) stays flat at 0. 5MB is generously
    // above that noise floor while still catching a real unbounded leak.
    assertPerfMetric("domRendererScroll", "heapGrowthBytes", Math.max(0, heapGrowthBytes), {
      toleranceFloor: 5_000_000
    });
  }
});
