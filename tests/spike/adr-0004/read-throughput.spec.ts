import { test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface ReadResult {
  readonly shape: string;
  readonly cellsPerFrame: number;
  readonly medianMs: number;
  readonly p95Ms: number;
  readonly maxMs: number;
}

interface FixtureWindow {
  __spikeReady?: boolean;
  __spikeRead?: ReadResult;
}

async function measure(page: Page, shape: string): Promise<ReadResult> {
  await page.goto(`/tests/spike/adr-0004/read-throughput.html?shape=${shape}`);
  await page.waitForFunction(() => (window as unknown as FixtureWindow).__spikeReady === true, undefined, {
    timeout: 180_000
  });
  const result = await page.evaluate(() => (window as unknown as FixtureWindow).__spikeRead);
  if (!result) {
    throw new Error(`read fixture returned nothing for shape=${shape}`);
  }
  return result;
}

for (const shape of ["map", "object"] as const) {
  test(`ADR 0004 spike — viewport read cost (${shape})`, async ({ page }) => {
    test.setTimeout(240_000);
    const result = await measure(page, shape);
    console.log(`[spike:read] ${JSON.stringify(result)}`);
  });
}
