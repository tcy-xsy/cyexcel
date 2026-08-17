import { test } from "@playwright/test";
import type { Page } from "@playwright/test";

import { assertPerfMetric } from "./support/thresholds";

interface BatchUpdateResult {
  readonly transactionMs: number;
  readonly repaintMs: number;
}

interface FixtureWindow {
  __cyexcelPerfReady?: boolean;
  __cyexcelRunBatchUpdate?: () => Promise<BatchUpdateResult>;
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/performance/batch-update.html");
  await page.waitForFunction(() => (window as unknown as FixtureWindow).__cyexcelPerfReady === true, undefined, {
    timeout: 90_000
  });
}

test("pasting 100,000 cells commits as one transaction and repaints quickly", async ({ page }) => {
  test.setTimeout(120_000);
  await gotoFixture(page);

  const result = await page.evaluate(() => (window as unknown as FixtureWindow).__cyexcelRunBatchUpdate?.());
  if (!result) {
    throw new Error("batch-update fixture did not return a result");
  }

  assertPerfMetric("batchUpdate", "transactionMs", result.transactionMs, { toleranceFloor: 40 });
  assertPerfMetric("batchUpdate", "repaintMs", result.repaintMs, { toleranceFloor: 20 });
});
