import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

import { assertPerfMetric } from "./support/thresholds";

interface PerfMetrics {
  readonly constructMs: number;
  readonly readyMs: number;
  readonly heapBytes: number | null;
  readonly domNodeCount: number;
}

interface FixtureWindow {
  __cyexcelPerfReady?: boolean;
  __cyexcelPerfMetrics?: PerfMetrics;
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/performance/first-paint.html");
  await page.waitForFunction(() => (window as unknown as FixtureWindow).__cyexcelPerfReady === true, undefined, {
    // Generating 1,000,000 rows x 50 columns and constructing the sheet
    // legitimately takes longer than Playwright's default 30s poll window
    // in a sandboxed/headless environment.
    timeout: 90_000
  });
}

test("1,000,000 rows x 50 columns: construct + first paint stays fast, DOM stays viewport-bound", async ({
  page
}) => {
  test.setTimeout(120_000);
  await gotoFixture(page);

  const metrics = await page.evaluate(
    () => (window as unknown as FixtureWindow).__cyexcelPerfMetrics as PerfMetrics
  );

  // Small floors, same reasoning as the other fast-metric specs: after the
  // buildDataSource() fix (no more O(rows * columns) eager row clone — see
  // src/core/CyExcelCore.ts's `data` option JSDoc and doc/progress.md's
  // Step 15 entry), these numbers are small enough (single/double-digit
  // ms) that ordinary rAF/scheduler jitter alone can cross a pure 35%
  // relative window.
  assertPerfMetric("firstPaint", "constructMs", metrics.constructMs, { toleranceFloor: 15 });
  assertPerfMetric("firstPaint", "readyMs", metrics.readyMs, { toleranceFloor: 30 });
  if (metrics.heapBytes !== null) {
    assertPerfMetric("firstPaint", "heapBytes", metrics.heapBytes);
  }
  assertPerfMetric("firstPaint", "domNodeCount", metrics.domNodeCount);

  // DOM node count must stay small/viewport-bound regardless of the
  // 1,000,000-row logical dataset (PRD §3.2 "DOM 单元格数量: 与可视区域相关，
  // 不随总行数增长") — a generous fixed ceiling catches an accidental
  // full-table DOM render without needing a second fixture at a smaller
  // row count just to diff against.
  expect(metrics.domNodeCount, "DOM node count must not scale with row count").toBeLessThan(500);
});
