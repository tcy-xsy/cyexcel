import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface FixtureWindow {
  __cyexcelFullLoopReady?: boolean;
  __sheet?: {
    getCell(row: number, column: number): unknown;
    undo(): void;
    redo(): void;
  };
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/full-loop.html");
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelFullLoopReady === true
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

function undo(page: Page): Promise<void> {
  return page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.undo());
}

function redo(page: Page): Promise<void> {
  return page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.redo());
}

test("renders visible cells as real, non-transparent pixels", async ({ page }) => {
  await gotoFixture(page);

  const hasPaintedPixel = await page.evaluate(() => {
    const canvas = document.querySelectorAll("#app canvas")[0] as HTMLCanvasElement;
    const ctx = canvas.getContext("2d") as CanvasRenderingContext2D;
    const dpr = window.devicePixelRatio || 1;
    const { data } = ctx.getImageData(
      Math.round(2 * dpr),
      Math.round(2 * dpr),
      Math.round(100 * dpr),
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

test("double-click opens an editor positioned at the cell rect", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.dblclick(box.x + 10, box.y + 10); // cell (0, 0)

  const input = page.locator("#app input");
  await expect(input).toBeVisible();
  const style = await input.evaluate((el: HTMLInputElement) => ({
    left: el.style.left,
    top: el.style.top,
    width: el.style.width,
    height: el.style.height
  }));
  expect(style).toEqual({ left: "0px", top: "0px", width: "120px", height: "24px" });
});

test("typing a new value and pressing Enter commits one undoable/redoable change", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.dblclick(box.x + 10, box.y + 10); // initial value "Item 0", selected on focus
  await page.keyboard.type("Changed");
  await page.keyboard.press("Enter");

  await expect(page.locator("#app input")).toHaveCount(0);
  expect(await getCell(page, 0, 0)).toBe("Changed");

  await undo(page);
  expect(await getCell(page, 0, 0)).toBe("Item 0");

  await redo(page);
  expect(await getCell(page, 0, 0)).toBe("Changed");
});

test("Enter on a selected cell opens the editor when not already editing", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.click(box.x + 10, box.y + 10); // select cell (0, 0), no editor yet
  await expect(page.locator("#app input")).toHaveCount(0);

  await page.keyboard.press("Enter");
  await expect(page.locator("#app input")).toBeVisible();
});

test("typing a printable character on a selected cell starts editing with that character", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.click(box.x + 130, box.y + 10); // select cell (0, 1), "price" column
  await page.keyboard.press("9");

  await expect(page.locator("#app input")).toHaveValue("9");
});

test("Escape cancels editing without changing the value", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.dblclick(box.x + 10, box.y + 10);
  await page.keyboard.type("Discarded");
  await page.keyboard.press("Escape");

  await expect(page.locator("#app input")).toHaveCount(0);
  expect(await getCell(page, 0, 0)).toBe("Item 0");
});

test("Enter during IME composition does not prematurely commit", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.dblclick(box.x + 10, box.y + 10);
  const input = page.locator("#app input");

  await input.evaluate((el) =>
    el.dispatchEvent(new CompositionEvent("compositionstart", { bubbles: true }))
  );
  await page.keyboard.press("Enter");
  await expect(input).toBeVisible(); // still editing — Enter was swallowed by composition

  await input.evaluate((el) =>
    el.dispatchEvent(new CompositionEvent("compositionend", { bubbles: true }))
  );
  await page.keyboard.press("Enter");
  await expect(page.locator("#app input")).toHaveCount(0);
});

test("scrolling commits the active edit per the default scrollCommitPolicy", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  await page.mouse.dblclick(box.x + 10, box.y + 10);
  await page.keyboard.type("ScrolledAway");

  await page.locator("#app").evaluate((el) => {
    const scroller = el.firstElementChild as HTMLElement;
    scroller.scrollTop = 100;
    scroller.dispatchEvent(new Event("scroll"));
  });

  await expect(page.locator("#app input")).toHaveCount(0);
  expect(await getCell(page, 0, 0)).toBe("ScrolledAway");
});

test("repeated edit cycles leave no editor DOM behind", async ({ page }) => {
  await gotoFixture(page);
  const box = await getContentCanvasBox(page);

  for (let i = 0; i < 3; i += 1) {
    await page.mouse.dblclick(box.x + 10, box.y + 10);
    await page.keyboard.press("Escape");
  }
  for (let i = 0; i < 3; i += 1) {
    await page.mouse.dblclick(box.x + 10, box.y + 10);
    await page.keyboard.type("x");
    await page.keyboard.press("Enter");
  }

  const inputCount = await page.evaluate(() => document.querySelectorAll("input").length);
  expect(inputCount).toBe(0);
});
