import { test } from "@playwright/test";
import type { Page } from "@playwright/test";

import { assertPerfMetric } from "./support/thresholds";

interface UpdateResult {
  readonly commandMs: number;
  readonly paintMs: number;
}

interface FixtureWindow {
  __cyexcelPerfReady?: boolean;
  __cyexcelUpdateOne?: () => Promise<UpdateResult>;
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/performance/single-cell-update.html");
  await page.waitForFunction(() => (window as unknown as FixtureWindow).__cyexcelPerfReady === true, undefined, {
    timeout: 90_000
  });
}

test("a single cell write on a 1,000,000-row sheet stays cheap regardless of row count", async ({ page }) => {
  test.setTimeout(120_000);
  await gotoFixture(page);

  const result = await page.evaluate(() => (window as unknown as FixtureWindow).__cyexcelUpdateOne?.());
  if (!result) {
    throw new Error("single-cell-update fixture did not return a result");
  }

  assertPerfMetric("singleCellUpdate", "commandMs", result.commandMs, { toleranceFloor: 5 });
  assertPerfMetric("singleCellUpdate", "paintMs", result.paintMs, { toleranceFloor: 20 });
});
