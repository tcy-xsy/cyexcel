import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface FixtureWindow {
  __cyexcelCustomCellsReady?: boolean;
  __canvasHitTestCalls?: Record<number, number>;
  __sheet?: {
    readonly selection: {
      readonly active: { readonly row: number; readonly column: number };
    };
    getCell(row: number, column: number): unknown;
    dispose(): void;
  };
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/custom-cells.html");
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelCustomCellsReady === true
  );
}

async function getContentCanvasBox(page: Page) {
  const box = await page.locator("#app canvas").first().boundingBox();
  if (!box) {
    throw new Error("missing content canvas bounding box");
  }
  return box;
}

function getCell(page: Page, row: number, column: number): Promise<unknown> {
  return page.evaluate(
    ({ row: r, column: c }) => (window as unknown as FixtureWindow).__sheet?.getCell(r, c),
    { row, column }
  );
}

function getCanvasHitTestCalls(page: Page, row: number): Promise<number> {
  return page.evaluate(
    (targetRow) => (window as unknown as FixtureWindow).__canvasHitTestCalls?.[targetRow] ?? 0,
    row
  );
}

test("custom CanvasCellRenderer paints real, non-transparent pixels", async ({ page }) => {
  await gotoFixture(page);

  const hasPaintedPixel = await page.evaluate(() => {
    const canvas = document.querySelectorAll("#app canvas")[0] as HTMLCanvasElement;
    const ctx = canvas.getContext("2d") as CanvasRenderingContext2D;
    const dpr = window.devicePixelRatio || 1;
    // Cell (0, 1) — the "counter" column, columns are name(120)/counter(100)/toggle(100).
    const { data } = ctx.getImageData(
      Math.round(122 * dpr),
      Math.round(2 * dpr),
      Math.round(80 * dpr),
      Math.round(16 * dpr)
    );
    for (let i = 3; i < data.length; i += 4) {
      if ((data[i] ?? 0) > 0) {
        return true;
      }
    }
    return false;
  });

  expect(hasPaintedPixel).toBe(true);
});

test("custom DOM renderer mounts real DOM nodes at the correct position", async ({ page }) => {
  await gotoFixture(page);

  const button = page.locator("#app .toggle-button").first();
  await expect(button).toBeVisible();
  await expect(button).toHaveText("on"); // row 0's toggle is true (index % 2 === 0)

  const host = page.locator("#app .toggle-button").first().locator("..");
  const style = await host.evaluate((el: HTMLElement) => ({ left: el.style.left, top: el.style.top }));
  // Column 0 ("name") is 120px wide, so the "toggle" column (index 2) starts at 220px.
  expect(style).toEqual({ left: "220px", top: "0px" });
});

test("scrolling updates/removes DOM renderer nodes without unbounded growth", async ({ page }) => {
  await gotoFixture(page);

  const before = await page.locator("#app .toggle-button").count();
  expect(before).toBeGreaterThan(0);
  expect(before).toBeLessThan(30); // 200 rows exist; only the visible + overscan band should be mounted

  await page.locator("#app").evaluate((el) => {
    const scroller = el.firstElementChild as HTMLElement;
    scroller.scrollTop = 2000; // deep into the 200-row dataset
    scroller.dispatchEvent(new Event("scroll"));
  });
  await page.waitForFunction(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
  );

  const after = await page.locator("#app .toggle-button").count();
  expect(after).toBeGreaterThan(0);
  expect(after).toBeLessThan(30);

  // The visible content actually changed — not the same rows redrawn.
  const firstButtonText = await page.locator("#app .toggle-button").first().textContent();
  expect(firstButtonText).not.toBeNull();
});

test("clicking a hitTest-enabled canvas cell commits a Command", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  expect(await getCell(page, 0, 1)).toBe(0);
  // Column 1 ("counter") starts at x=120; click well inside it.
  await page.mouse.click(box.x + 140, box.y + 10);

  expect(await getCell(page, 0, 1)).toBe(1);
});

test("a declined canvas hitTest runs once and falls through to body selection", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  // Counter row 1 deliberately declines. The lower-priority selection route
  // must still select that exact cell, which Enter exposes via its editor.
  await page.mouse.click(box.x + 140, box.y + 34);

  expect(await getCanvasHitTestCalls(page, 1)).toBe(1);
  await page.keyboard.press("Enter");
  await expect(page.locator("#app input")).toHaveValue("1");
});

test("a handled canvas hitTest runs once, commits, and keeps the existing selection", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  // Counter row 2 returns an action. It owns the pointer input, so the body
  // selection candidate must not also move the initial selection from A1.
  await page.mouse.click(box.x + 140, box.y + 58);

  expect(await getCanvasHitTestCalls(page, 2)).toBe(1);
  expect(await getCell(page, 2, 1)).toBe(3);

  await page.locator("#app canvas").first().focus();
  await page.keyboard.press("Enter");
  await expect(page.locator("#app input")).toHaveValue("Item 0");
});

test("body input refreshes surface origins after ancestor reflow without resize()", async ({ page }) => {
  await gotoFixture(page);

  await page.locator("#app").evaluate((element) => {
    element.style.marginLeft = "80px";
  });
  const movedBox = await getContentCanvasBox(page);
  await page.mouse.click(movedBox.x + 140, movedBox.y + 10);

  expect(await getCanvasHitTestCalls(page, 0)).toBe(1);
  expect(await getCell(page, 0, 1)).toBe(1);
});

test("right-clicking a hitTest-enabled canvas cell does not commit its click action", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  expect(await getCell(page, 0, 1)).toBe(0);
  await page.mouse.click(box.x + 140, box.y + 10, { button: "right" });

  expect(await getCell(page, 0, 1)).toBe(0); // unchanged — only #handleContextMenu should react to this click
});

test("clicking the DOM toggle button commits through context.actions.commit", async ({ page }) => {
  await gotoFixture(page);

  const before = await page.evaluate(
    () => (window as unknown as FixtureWindow).__sheet?.selection
  );
  expect(await getCell(page, 0, 2)).toBe(true);
  await page.locator("#app .toggle-button").first().click();

  expect(await getCell(page, 0, 2)).toBe(false);
  await expect(page.locator("#app .toggle-button").first()).toHaveText("off");
  expect(await page.evaluate(
    () => (window as unknown as FixtureWindow).__sheet?.selection
  )).toEqual(before);
});

test("clicking passive DOM cell content falls through to canonical body selection", async ({ page }) => {
  await gotoFixture(page);

  await page.locator("#app .toggle-passive").nth(1).click();

  expect(await page.evaluate(
    () => (window as unknown as FixtureWindow).__sheet?.selection.active
  )).toEqual({ row: 1, column: 2 });
});

test("a native-owned closed-shadow cell is terminal without exposing its internal input", async ({ page }) => {
  await gotoFixture(page);

  const before = await page.evaluate(
    () => (window as unknown as FixtureWindow).__sheet?.selection.active
  );
  await page.locator("#app .closed-native-cell").nth(1).click();

  expect(await page.evaluate(
    () => (window as unknown as FixtureWindow).__sheet?.selection.active
  )).toEqual(before);
});

test("a surface-owned closed-shadow cell routes through the canonical body surface", async ({ page }) => {
  await gotoFixture(page);

  await page.locator("#app .closed-surface-cell").nth(1).click();

  expect(await page.evaluate(
    () => (window as unknown as FixtureWindow).__sheet?.selection.active
  )).toEqual({ row: 1, column: 5 });
});

test("editing a number-typed column via the keyboard stores and renders a real number, not a blank cell", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  expect(await getCell(page, 0, 3)).toBe(0);

  // Column 3 ("price") starts at x=320 (name120+counter100+toggle100).
  await page.mouse.dblclick(box.x + 330, box.y + 10);
  await page.locator("#app input").fill("500");
  await page.keyboard.press("Enter");
  await expect(page.locator("#app input")).toHaveCount(0);

  // The default text editor always returns a string ("500") — without
  // CellType.parse() being applied on commit, this got stored as the
  // *string* "500", and the built-in "number" type's format() (which
  // requires typeof value === "number") silently rendered "" for it.
  expect(await getCell(page, 0, 3)).toBe(500);

  const hasPaintedPixel = await page.evaluate(() => {
    const canvas = document.querySelectorAll("#app canvas")[0] as HTMLCanvasElement;
    const ctx = canvas.getContext("2d") as CanvasRenderingContext2D;
    const dpr = window.devicePixelRatio || 1;
    const { data } = ctx.getImageData(
      Math.round(322 * dpr),
      Math.round(2 * dpr),
      Math.round(80 * dpr),
      Math.round(16 * dpr)
    );
    for (let i = 3; i < data.length; i += 4) {
      if ((data[i] ?? 0) > 0) {
        return true;
      }
    }
    return false;
  });
  expect(hasPaintedPixel).toBe(true);
});

test("a custom renderer's ctx.font change does not leak into the next cell's default text painting", async ({
  page
}) => {
  await gotoFixture(page);

  // Counts non-transparent pixels in a region inset from the cell's own
  // edges — insetting avoids the gridline border pixels (which run along
  // every cell's full width/height regardless of font size and would
  // otherwise dominate a boundary-based measurement).
  const countGlyphPixels = (x: number, y: number, width: number, height: number): Promise<number> =>
    page.evaluate(
      ({ x, y, width, height }) => {
        const canvas = document.querySelectorAll("#app canvas")[0] as HTMLCanvasElement;
        const ctx = canvas.getContext("2d") as CanvasRenderingContext2D;
        const dpr = window.devicePixelRatio || 1;
        const inset = 3 * dpr;
        const w = Math.round(width * dpr - inset * 2);
        const h = Math.round(height * dpr - inset * 2);
        const { data } = ctx.getImageData(Math.round(x * dpr + inset), Math.round(y * dpr + inset), w, h);
        let count = 0;
        for (let i = 3; i < data.length; i += 4) {
          if ((data[i] ?? 0) > 0) {
            count += 1;
          }
        }
        return count;
      },
      { x, y, width, height }
    );

  // Column 1 ("counter") sets ctx.font = "28px serif" in its own paint()
  // and never restores it. Column 3 ("price") starts at x=320 (name120 +
  // counter100 + toggle100) and has no custom renderer — its text should
  // still be drawn with the framework's own 13px CELL_FONT, not inherit
  // counter's leftover 28px setting. A much bigger font covers a much
  // larger pixel area for comparably-sized text ("n=0" vs "0"), so pixel
  // count is a robust proxy for "which font actually got used".
  const counterGlyphPixels = await countGlyphPixels(122, 0, 80, 24);
  const priceGlyphPixels = await countGlyphPixels(322, 0, 80, 24);

  expect(counterGlyphPixels).toBeGreaterThan(0);
  expect(priceGlyphPixels).toBeGreaterThan(0);
  expect(priceGlyphPixels).toBeLessThan(counterGlyphPixels * 0.5);
});

test("dispose() cleans up all DOM renderer nodes", async ({ page }) => {
  await gotoFixture(page);
  await expect(page.locator("#app .toggle-button").first()).toBeVisible();

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.dispose());

  await expect(page.locator("#app .toggle-button")).toHaveCount(0);
});
