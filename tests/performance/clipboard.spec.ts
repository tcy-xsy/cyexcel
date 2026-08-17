import { test } from "@playwright/test";
import type { Page } from "@playwright/test";

import { assertPerfMetric } from "./support/thresholds";

interface PlainCopyResult {
  readonly planMs: number;
  readonly copyMs: number;
}

interface HugeArrayCopyResult {
  readonly copyMs: number;
  readonly longTaskCount: number;
}

interface FixtureWindow {
  __cyexcelPerfReady?: boolean;
  __cyexcelRunPlainCopy?: () => Promise<PlainCopyResult>;
  __cyexcelRunHugeArrayCopy?: () => Promise<HugeArrayCopyResult>;
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/performance/clipboard.html");
  await page.waitForFunction(() => (window as unknown as FixtureWindow).__cyexcelPerfReady === true, undefined, {
    timeout: 90_000
  });
}

test("copying 10,000 plain cells plans and serializes quickly", async ({ page, context }) => {
  test.setTimeout(60_000);
  await context.grantPermissions(["clipboard-read", "clipboard-write"]);
  await gotoFixture(page);

  const result = await page.evaluate(() => (window as unknown as FixtureWindow).__cyexcelRunPlainCopy?.());
  if (!result) {
    throw new Error("plain-copy fixture did not return a result");
  }

  assertPerfMetric("clipboardPlain", "planMs", result.planMs, { toleranceFloor: 10 });
  // copyMs's floor is wider than the other timing floors in this file — it
  // crosses into the real OS/browser clipboard IPC path (navigator.clipboard
  // via context.grantPermissions()), which has its own scheduling variance
  // on top of ordinary JS wall-clock noise.
  assertPerfMetric("clipboardPlain", "copyMs", result.copyMs, { toleranceFloor: 80 });
});

test("copying a single cell with a 100k-item array produces no long task", async ({ page, context }) => {
  test.setTimeout(60_000);
  await context.grantPermissions(["clipboard-read", "clipboard-write"]);
  await gotoFixture(page);

  const result = await page.evaluate(() => (window as unknown as FixtureWindow).__cyexcelRunHugeArrayCopy?.());
  if (!result) {
    throw new Error("huge-array-copy fixture did not return a result");
  }

  assertPerfMetric("clipboardHugeArray", "copyMs", result.copyMs, { toleranceFloor: 80 });
  assertPerfMetric("clipboardHugeArray", "longTaskCount", result.longTaskCount);
});
