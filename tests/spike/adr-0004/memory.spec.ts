import { test } from "@playwright/test";
import type { Page } from "@playwright/test";

test.use({
  launchOptions: {
    args: ["--enable-precise-memory-info", "--js-flags=--expose-gc"]
  }
});

interface MemoryResult {
  readonly shape: string;
  readonly occupied: number;
  readonly beforeBytes: number | null;
  readonly afterBytes: number | null;
  readonly deltaBytes: number | null;
  readonly bytesPerCell: number | null;
  readonly gcAvailable: boolean;
}

interface FixtureWindow {
  __spikeReady?: boolean;
  __spikeMemory?: MemoryResult;
  gc?: () => void;
}

async function measure(page: Page, shape: string): Promise<MemoryResult> {
  await page.goto(`/tests/spike/adr-0004/memory.html?shape=${shape}`);
  await page.waitForFunction(() => (window as unknown as FixtureWindow).__spikeReady === true, undefined, {
    timeout: 180_000
  });

  const result = await page.evaluate(() => (window as unknown as FixtureWindow).__spikeMemory);
  if (!result) {
    throw new Error(`memory fixture returned nothing for shape=${shape}`);
  }
  return result;
}

// Spike：不断言阈值，只把数字打出来供 Task 6 的报告引用。
// 断言会把一个"我们还不知道答案"的问题伪装成"我们知道答案且它没退化"。
for (const shape of ["map", "object", "rows"] as const) {
  test(`ADR 0004 spike — heap for 100k x 50 (${shape}) — run 1`, async ({ page }) => {
    test.setTimeout(240_000);
    const result = await measure(page, shape);
    console.log(`[spike:memory] ${JSON.stringify(result)}`);
    if (result.deltaBytes === null) {
      throw new Error("performance.memory unavailable — run this spec under --project=chrome");
    }
    if (!result.gcAvailable) {
      throw new Error("gc() unavailable — run with --js-flags=--expose-gc");
    }
  });

  test(`ADR 0004 spike — heap for 100k x 50 (${shape}) — run 2`, async ({ page }) => {
    test.setTimeout(240_000);
    const result = await measure(page, shape);
    console.log(`[spike:memory] ${JSON.stringify(result)}`);
    if (result.deltaBytes === null) {
      throw new Error("performance.memory unavailable — run this spec under --project=chrome");
    }
    if (!result.gcAvailable) {
      throw new Error("gc() unavailable — run with --js-flags=--expose-gc");
    }
  });

  test(`ADR 0004 spike — heap for 100k x 50 (${shape}) — run 3`, async ({ page }) => {
    test.setTimeout(240_000);
    const result = await measure(page, shape);
    console.log(`[spike:memory] ${JSON.stringify(result)}`);
    if (result.deltaBytes === null) {
      throw new Error("performance.memory unavailable — run this spec under --project=chrome");
    }
    if (!result.gcAvailable) {
      throw new Error("gc() unavailable — run with --js-flags=--expose-gc");
    }
  });
}
