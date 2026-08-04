import { expect, test } from "@playwright/test";

test("configureCanvasForDpr sizes the backing store to devicePixelRatio", async ({ page }) => {
  await page.addInitScript(() => {
    Object.defineProperty(window, "devicePixelRatio", { value: 2, configurable: true });
  });

  await page.goto("/tests/browser/canvas-render.html");
  await page.waitForFunction(
    () => (window as unknown as { __cyexcelRenderReady?: boolean }).__cyexcelRenderReady === true
  );

  const canvasSize = await page.evaluate(() => {
    const canvas = document.getElementById("canvas") as HTMLCanvasElement;
    return { width: canvas.width, height: canvas.height };
  });

  expect(canvasSize).toEqual({ width: 600, height: 200 });
});

test("paints visible cells with real, non-transparent pixels", async ({ page }) => {
  await page.addInitScript(() => {
    Object.defineProperty(window, "devicePixelRatio", { value: 2, configurable: true });
  });

  await page.goto("/tests/browser/canvas-render.html");
  await page.waitForFunction(
    () => (window as unknown as { __cyexcelRenderReady?: boolean }).__cyexcelRenderReady === true
  );

  const hasPaintedPixel = await page.evaluate(() => {
    const canvas = document.getElementById("canvas") as HTMLCanvasElement;
    const ctx = canvas.getContext("2d") as CanvasRenderingContext2D;
    const dpr = window.devicePixelRatio || 1;

    // Sample most of cell (0, 0)'s pixel rect (x:0-60, y:0-20 logical),
    // shrunk slightly to avoid edge artifacts, converted to physical pixels.
    const x = Math.round(2 * dpr);
    const y = Math.round(2 * dpr);
    const width = Math.round(50 * dpr);
    const height = Math.round(16 * dpr);

    const { data } = ctx.getImageData(x, y, width, height);
    for (let i = 3; i < data.length; i += 4) {
      if ((data[i] ?? 0) > 0) {
        return true;
      }
    }
    return false;
  });

  expect(hasPaintedPixel).toBe(true);
});
