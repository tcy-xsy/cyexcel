import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface FixtureWindow {
  __cyexcelHeadersReady?: boolean;
  __sheet?: {
    getCell(row: number, column: number): unknown;
    dispose(): void;
  };
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/headers.html");
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelHeadersReady === true
  );
}

async function nextFrame(page: Page): Promise<void> {
  await page.waitForFunction(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
  );
}

function getCell(page: Page, row: number, column: number): Promise<unknown> {
  return page.evaluate(
    ({ row: r, column: c }) => (window as unknown as FixtureWindow).__sheet?.getCell(r, c),
    { row, column }
  );
}

// DomHeaderManager positions the <select>'s wrapper host div, not the
// <select> itself (mirrors DOMCellManager: the factory mounts its own
// content inside the container DomHeaderManager creates and positions).
function statusHeaderHostLeft(page: Page): Promise<string> {
  return page
    .locator("#app select")
    .evaluate((el: HTMLElement) => (el.parentElement as HTMLElement).style.left);
}

test("header canvas paints real, non-transparent pixels", async ({ page }) => {
  await gotoFixture(page);

  const hasPaintedPixel = await page.evaluate(() => {
    const canvas = document.querySelectorAll("#app canvas")[2] as HTMLCanvasElement; // header canvas
    const ctx = canvas.getContext("2d") as CanvasRenderingContext2D;
    const dpr = window.devicePixelRatio || 1;
    const { data } = ctx.getImageData(0, 0, Math.round(400 * dpr), Math.round(28 * dpr));
    for (let i = 3; i < data.length; i += 4) {
      if ((data[i] ?? 0) > 0) {
        return true;
      }
    }
    return false;
  });

  expect(hasPaintedPixel).toBe(true);
});

test("DOM header stays aligned with body content during horizontal scroll (退出条件 1)", async ({
  page
}) => {
  await gotoFixture(page);

  // Column x-order: sku(100, frozen) | name(140) | price(90) | qty(80) | status(100).
  // status's absolute left = 100+140+90+80 = 410; not frozen, so viewport-relative
  // left = 410 - scrollLeft.
  expect(await statusHeaderHostLeft(page)).toBe("410px");

  await page.locator("#app").evaluate((el) => {
    const scroller = el.firstElementChild as HTMLElement;
    scroller.scrollLeft = 50;
    scroller.dispatchEvent(new Event("scroll"));
  });
  await nextFrame(page);

  expect(await statusHeaderHostLeft(page)).toBe("360px");
});

test("frozen row/column stay pinned; editing a scrolled-into-view body cell still works (先写测试 #7)", async ({
  page
}) => {
  await gotoFixture(page);

  await page.locator("#app").evaluate((el) => {
    const scroller = el.firstElementChild as HTMLElement;
    // 408 = 17 * rowHeight(24): the content canvas's own local y-axis
    // reserves its first 24px (frozenRowCount 1 * rowHeight) for the
    // frozen row band, so scrollable content only starts at local y=24 —
    // row 18 is the first scrollable row whose band (y:[24,48)) begins
    // right there. See doc/architecture.md §6.1.
    scroller.scrollTop = 408;
    scroller.scrollLeft = 0;
    scroller.dispatchEvent(new Event("scroll"));
  });
  await nextFrame(page);

  const contentBox = await page.locator("#app canvas").first().boundingBox();
  if (!contentBox) {
    throw new Error("missing content canvas box");
  }

  // Double-click comfortably inside the scrollable band (past the 24px
  // frozen row band) — must open an editor positioned inside the viewport,
  // not off-screen at its absolute content position.
  await page.mouse.dblclick(contentBox.x + 10, contentBox.y + 40);
  const input = page.locator("#app input");
  await expect(input).toBeVisible();
  const style = await input.evaluate((el: HTMLInputElement) => ({ left: el.style.left, top: el.style.top }));
  expect(style).toEqual({ left: "0px", top: "24px" });
  await input.fill("Edited");
  await page.keyboard.press("Enter");
  await expect(page.locator("#app input")).toHaveCount(0);

  expect(await getCell(page, 18, 0)).toBe("Edited");
});

test("dragging a header column boundary resizes the column and shifts later columns", async ({
  page
}) => {
  await gotoFixture(page);

  const headerBox = await page.locator("#app canvas").nth(2).boundingBox();
  if (!headerBox) {
    throw new Error("missing header canvas box");
  }

  expect(await statusHeaderHostLeft(page)).toBe("410px");

  // sku|name boundary sits at x=100 (sku is frozen, so this position is
  // scroll-independent) — drag it 30px to the right.
  await page.mouse.move(headerBox.x + 100, headerBox.y + 14);
  await page.mouse.down();
  await page.mouse.move(headerBox.x + 130, headerBox.y + 14);
  await page.mouse.up();
  await page.waitForTimeout(50);

  // The 420px-wide fixture no longer fits the status column after sku grew
  // by 30px — scroll it back into view (DomHeaderManager correctly
  // unmounted it, same virtualization the "scrolling out" unit tests cover)
  // before checking its new position.
  await page.locator("#app").evaluate((el) => {
    const scroller = el.firstElementChild as HTMLElement;
    scroller.scrollLeft = 100;
    scroller.dispatchEvent(new Event("scroll"));
  });
  await nextFrame(page);

  // status absolute left after resize = 130(sku)+140+90+80 = 440; minus scrollLeft 100.
  expect(await statusHeaderHostLeft(page)).toBe("340px");
});

test("DOM header nodes are cleaned up on dispose()", async ({ page }) => {
  await gotoFixture(page);
  await expect(page.locator("#app select")).toBeVisible();

  await page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.dispose());

  await expect(page.locator("#app select")).toHaveCount(0);
});
