import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface FixtureWindow {
  __cyexcelVisualReady?: boolean;
  __sheet?: {
    startEdit(row: number, column: number): void;
  };
}

/**
 * Must run before page.goto() — devicePixelRatio is read at construction/
 * paint time, not observed live. Same technique as
 * tests/browser/canvas-render.spec.ts's DPR override.
 */
async function gotoFixtureAtDpr(page: Page, dpr: number): Promise<void> {
  await page.addInitScript((value) => {
    Object.defineProperty(window, "devicePixelRatio", { value, configurable: true });
  }, dpr);
  await page.goto("/tests/visual/dpr-consistency.html");
  await page.waitForFunction(() => (window as unknown as FixtureWindow).__cyexcelVisualReady === true);
}

async function getContentCanvasBox(page: Page) {
  const box = await page.locator("#app canvas").first().boundingBox();
  if (!box) {
    throw new Error("missing content canvas bounding box");
  }
  return box;
}

/**
 * Alpha-weighted average RGB of a small logical-pixel box, scaled for
 * devicePixelRatio. Weighted by alpha (not a plain per-pixel average like
 * tests/browser/visual-theme.spec.ts's DPR=1-only `sampleBox`) — a thin
 * (1-2 logical px) anti-aliased gridline stroke only covers part of a
 * fixed-size sample box, and *which* part shifts with `Math.round(x *
 * dpr)`'s sub-pixel rounding at different DPR. An unweighted average mixes
 * the line's real color with fully-transparent (r=0, a=0) background
 * pixels in a DPR-dependent proportion — e.g. this box's "r" read 225.5 at
 * DPR=1 but 113 at DPR=2 for the exact same rendered gridline (verified by
 * hand while building this test: the frozen separator, a solid `fillRect`
 * block, sampled a rock-stable 156 at every DPR the whole time — this was
 * a sampling artifact specific to thin strokes, not a real DPR rendering
 * bug). Weighting by alpha reports the color of whatever fraction of the
 * box is actually painted, independent of how much transparent margin the
 * box happens to include.
 */
function sampleBox(
  page: Page,
  canvasIndex: number,
  x: number,
  y: number,
  width: number,
  height: number
): Promise<{ r: number; nonTransparentCount: number }> {
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
      let rWeighted = 0;
      let alphaSum = 0;
      let nonTransparentCount = 0;
      for (let i = 0; i < data.length; i += 4) {
        const alpha = data[i + 3] ?? 0;
        rWeighted += (data[i] ?? 0) * alpha;
        alphaSum += alpha;
        if (alpha > 0) {
          nonTransparentCount += 1;
        }
      }
      return { r: alphaSum > 0 ? rWeighted / alphaSum : 0, nonTransparentCount };
    },
    { canvasIndex, x, y, width, height }
  );
}

for (const dpr of [1, 2, 3]) {
  test.describe(`devicePixelRatio = ${dpr}`, () => {
    test("content canvas backing store scales to the container size x DPR", async ({ page }) => {
      await gotoFixtureAtDpr(page, dpr);

      const canvasSize = await page.evaluate(() => {
        const canvas = document.querySelectorAll("#app canvas")[0] as HTMLCanvasElement;
        return { width: canvas.width, height: canvas.height };
      });

      // #app is 400x300 logical px (dpr-consistency.html), no header/row-header band.
      expect(canvasSize).toEqual({ width: 400 * dpr, height: 300 * dpr });
    });

    test("the frozen column boundary stays darker than an ordinary gridline", async ({ page }) => {
      await gotoFixtureAtDpr(page, dpr);

      // x=100 is the frozen/scrollable column boundary (freeze: {columns: 1},
      // 100px-wide columns); y=65 sits inside row 2 (rows are 24px tall,
      // frozen row separator is at y=24), avoiding it.
      const separator = await sampleBox(page, 0, 99, 65, 2, 6);
      // x=300 is an ordinary column boundary (col 2/3) at the same row band.
      const ordinaryGridline = await sampleBox(page, 0, 299, 65, 2, 6);

      expect(separator.nonTransparentCount).toBeGreaterThan(0);
      expect(ordinaryGridline.nonTransparentCount).toBeGreaterThan(0);
      expect(separator.r).toBeLessThan(ordinaryGridline.r - 20);
    });

    test("the frozen row boundary stays darker than an ordinary gridline", async ({ page }) => {
      await gotoFixtureAtDpr(page, dpr);

      // y=24 is the frozen/scrollable row boundary; x=250 sits inside
      // column 2 (100px-wide columns), avoiding any column edge.
      const separator = await sampleBox(page, 0, 245, 23, 10, 2);
      // y=72 is an ordinary row boundary (row 2/3) at the same column.
      const ordinaryGridline = await sampleBox(page, 0, 245, 71, 10, 2);

      expect(separator.nonTransparentCount).toBeGreaterThan(0);
      expect(ordinaryGridline.nonTransparentCount).toBeGreaterThan(0);
      expect(separator.r).toBeLessThan(ordinaryGridline.r - 20);
    });

    test("selecting a scrollable cell still draws a solid fill-handle square at its corner", async ({ page }) => {
      await gotoFixtureAtDpr(page, dpr);
      const box = await getContentCanvasBox(page);

      // Row 2, column 2 (a scrollable, non-frozen cell): rect is {x:200, y:48, width:100, height:24}.
      await page.mouse.click(box.x + 250, box.y + 60);
      await page.waitForFunction(
        () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
      );

      // A small box centered on the corner (300, 72) — both border edges
      // plus the fill-handle square cross it.
      const cornerBox = await sampleBox(page, 1, 296, 68, 8, 8);
      // A same-sized box centered on the top edge's midpoint, far from any
      // corner — only ever crosses the thin selection border line.
      const edgeBox = await sampleBox(page, 1, 246, 44, 8, 8);

      expect(cornerBox.nonTransparentCount).toBeGreaterThan(edgeBox.nonTransparentCount * 2);
    });

    test("the DOM editor overlay stays pixel-aligned with the cell rect", async ({ page }) => {
      await gotoFixtureAtDpr(page, dpr);

      // The editor's CSS position comes straight from LayoutEngine's cell
      // rect in logical/CSS pixels — devicePixelRatio only scales the
      // Canvas backing store, never DOM layout — so this should hold
      // identically at every DPR. Regression guard against a future change
      // accidentally threading DPR into this coordinate path (this
      // codebase has twice shipped real scroll-offset coordinate bugs in
      // exactly this editor-positioning area; see doc/progress.md).
      await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.startEdit(2, 2));

      const input = page.locator("#app input");
      await expect(input).toBeVisible();
      const style = await input.evaluate((el: HTMLInputElement) => ({
        left: el.style.left,
        top: el.style.top,
        width: el.style.width,
        height: el.style.height
      }));
      expect(style).toEqual({ left: "200px", top: "48px", width: "100px", height: "24px" });
    });
  });
}
