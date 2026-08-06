import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface FixtureWindow {
  __cyexcelDistEsmReady?: boolean;
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/build/dist-esm.html");
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelDistEsmReady === true
  );
}

test("importing CyExcelCore from the built dist/core.js ESM artifact creates a real, painted table", async ({
  page
}) => {
  await gotoFixture(page);

  const hasPaintedPixel = await page.evaluate(() => {
    const canvas = document.querySelectorAll("#app canvas")[0] as HTMLCanvasElement;
    const ctx = canvas.getContext("2d") as CanvasRenderingContext2D;
    const dpr = window.devicePixelRatio || 1;
    const { data } = ctx.getImageData(0, 0, Math.round(60 * dpr), Math.round(20 * dpr));
    for (let i = 3; i < data.length; i += 4) {
      if ((data[i] ?? 0) > 0) {
        return true;
      }
    }
    return false;
  });

  expect(hasPaintedPixel).toBe(true);
});
