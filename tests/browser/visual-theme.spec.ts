import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface FixtureWindow {
  __cyexcelVisualThemeReady?: boolean;
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/visual-theme.html");
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelVisualThemeReady === true
  );
}

async function getContentCanvasBox(page: Page) {
  const box = await page.locator("#app canvas").first().boundingBox();
  if (!box) {
    throw new Error("missing content canvas bounding box");
  }
  return box;
}

/** Average RGBA of a small logical-pixel box, scaled for devicePixelRatio. */
function sampleBox(
  page: Page,
  canvasIndex: number,
  x: number,
  y: number,
  width: number,
  height: number
): Promise<{ r: number; g: number; b: number; a: number; nonTransparentCount: number }> {
  return page.evaluate(
    ({ canvasIndex, x, y, width, height }) => {
      const canvas = document.querySelectorAll("#app canvas")[canvasIndex] as HTMLCanvasElement;
      const ctx = canvas.getContext("2d") as CanvasRenderingContext2D;
      const dpr = window.devicePixelRatio || 1;
      const { data } = ctx.getImageData(
        Math.round(x * dpr),
        Math.round(y * dpr),
        Math.max(1, Math.round(width * dpr)),
        Math.max(1, Math.round(height * dpr))
      );
      let r = 0;
      let g = 0;
      let b = 0;
      let a = 0;
      let nonTransparentCount = 0;
      const pixelCount = data.length / 4;
      for (let i = 0; i < data.length; i += 4) {
        r += data[i] ?? 0;
        g += data[i + 1] ?? 0;
        b += data[i + 2] ?? 0;
        a += data[i + 3] ?? 0;
        if ((data[i + 3] ?? 0) > 0) {
          nonTransparentCount += 1;
        }
      }
      return { r: r / pixelCount, g: g / pixelCount, b: b / pixelCount, a: a / pixelCount, nonTransparentCount };
    },
    { canvasIndex, x, y, width, height }
  );
}

test("a light gridline is visible at an ordinary cell boundary", async ({ page }) => {
  await gotoFixture(page);

  // Boundary between column 2 and column 3 (x=300), row 2 (y=60..84, avoids
  // the frozen-row separator at y=24) — an ordinary, non-frozen boundary.
  const sample = await sampleBox(page, 0, 299, 65, 2, 6);

  expect(sample.nonTransparentCount).toBeGreaterThan(0);
  // GRID_LINE_COLOR is a light gray (#e2e5ea) — clearly lighter than black text.
  expect(sample.r).toBeGreaterThan(150);
});

test("the header band has a distinct background color, not transparent", async ({ page }) => {
  await gotoFixture(page);

  // Interior of column 1's header cell (x=150..190), away from the text and
  // any border pixels.
  const sample = await sampleBox(page, 2, 160, 8, 20, 12);

  expect(sample.nonTransparentCount).toBeGreaterThan(0);
  // HEADER_BACKGROUND_COLOR is #f3f4f6 — a very light, near-white gray.
  expect(sample.r).toBeGreaterThan(220);
  expect(sample.g).toBeGreaterThan(220);
  expect(sample.b).toBeGreaterThan(220);
});

test("the frozen column boundary is drawn darker than an ordinary gridline", async ({ page }) => {
  await gotoFixture(page);

  // x=100 is both the frozen/scrollable column boundary and an ordinary
  // cell edge; y=65 avoids the frozen row separator at y=24.
  const separator = await sampleBox(page, 0, 99, 65, 2, 6);
  // x=300 is an ordinary (non-frozen) column boundary at the same row band.
  const ordinaryGridline = await sampleBox(page, 0, 299, 65, 2, 6);

  expect(separator.nonTransparentCount).toBeGreaterThan(0);
  expect(ordinaryGridline.nonTransparentCount).toBeGreaterThan(0);
  // FROZEN_SEPARATOR_COLOR (#9ca3af) is noticeably darker than GRID_LINE_COLOR (#e2e5ea).
  expect(separator.r).toBeLessThan(ordinaryGridline.r - 20);
});

test("the frozen row boundary is drawn darker than an ordinary gridline", async ({ page }) => {
  await gotoFixture(page);

  // y=24 is the frozen/scrollable row boundary; x=250 avoids any column edge.
  const separator = await sampleBox(page, 0, 245, 23, 10, 2);
  // y=48 is an ordinary (non-frozen) row boundary at the same column.
  const ordinaryGridline = await sampleBox(page, 0, 245, 47, 10, 2);

  expect(separator.nonTransparentCount).toBeGreaterThan(0);
  expect(ordinaryGridline.nonTransparentCount).toBeGreaterThan(0);
  expect(separator.r).toBeLessThan(ordinaryGridline.r - 20);
});

test("selecting a single cell draws a solid fill-handle square at its bottom-right corner", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  // Row 2, column 2 (a scrollable, non-frozen cell): rect is {x:200, y:48, width:100, height:24}.
  await page.mouse.click(box.x + 250, box.y + 60);
  await page.waitForFunction(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
  );

  // A small box centered on the corner (300, 72) should be mostly filled —
  // more than a thin border line passing through the same-sized box could
  // produce (the corner has both border edges crossing plus the handle fill).
  const cornerBox = await sampleBox(page, 1, 296, 68, 8, 8);
  // A same-sized box centered on the top edge's midpoint, far from any
  // corner, only ever crosses the thin selection border line.
  const edgeBox = await sampleBox(page, 1, 246, 44, 8, 8);

  expect(cornerBox.nonTransparentCount).toBeGreaterThan(edgeBox.nonTransparentCount * 2);
});
