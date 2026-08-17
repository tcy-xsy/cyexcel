import { test } from "@playwright/test";
import type { Page } from "@playwright/test";

import { assertPerfMetric, SCROLL_DURATION_MS } from "./support/thresholds";

interface ScrollResult {
  readonly fps: number;
  readonly p95FrameMs: number;
  readonly longTaskCount: number;
}

interface FixtureWindow {
  __cyexcelPerfReady?: boolean;
  __cyexcelRunScroll?: (durationMs: number) => Promise<ScrollResult>;
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/performance/scroll.html");
  await page.waitForFunction(() => (window as unknown as FixtureWindow).__cyexcelPerfReady === true, undefined, {
    timeout: 90_000
  });
}

test("1,000,000-row continuous scroll holds FPS and produces no long tasks", async ({ page }) => {
  test.setTimeout(SCROLL_DURATION_MS + 60_000);
  await gotoFixture(page);

  const result = await page.evaluate(
    (durationMs) => (window as unknown as FixtureWindow).__cyexcelRunScroll?.(durationMs),
    SCROLL_DURATION_MS
  );
  if (!result) {
    throw new Error("scroll fixture did not return a result");
  }

  assertPerfMetric("scroll", "fps", result.fps);
  assertPerfMetric("scroll", "p95FrameMs", result.p95FrameMs, { toleranceFloor: 10 });
  assertPerfMetric("scroll", "longTaskCount", result.longTaskCount);
});
