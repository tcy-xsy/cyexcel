import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface FixtureWindow {
  __cyexcelContextMenuBoundsReady?: boolean;
  __sheet?: {
    getCell(row: number, column: number): unknown;
  };
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/context-menu-bounds.html");
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelContextMenuBoundsReady === true
  );
}

function getCell(page: Page, row: number, column: number): Promise<unknown> {
  return page.evaluate(
    ({ row: r, column: c }) => (window as unknown as FixtureWindow).__sheet?.getCell(r, c),
    { row, column }
  );
}

// Real user report: right-clicking a column header near the right edge of a
// narrow container opened a menu that visibly ran off the edge ("在右侧插入"
// showed as truncated text), and clicking the truncated item silently did
// nothing ("插入列没有作用") — the overflowing portion was clipped by the
// scroller's own `overflow: auto`, which also makes it unclickable, not just
// invisible. Root cause: DefaultMenuRenderer positioned purely from
// anchorPoint with no awareness of the available space. Fixed via
// ContextMenuService's `getBounds` -> DefaultMenuRenderer's #clampToBounds().
test("a context menu opened near the container's right edge stays fully within it, instead of being clipped", async ({
  page
}) => {
  await gotoFixture(page);
  const appBox = await page.locator("#app").boundingBox();
  if (!appBox) throw new Error("missing app box");

  // Column 2 (x:[160,240)) header, near its right edge — the menu's
  // minWidth (160px) can't fit between x=230 and the container's right
  // edge at x=250.
  await page.mouse.click(appBox.x + 230, appBox.y + 10, { button: "right" });

  const menu = page.locator('[role="menu"]');
  await expect(menu).toBeVisible();
  const menuBox = await menu.boundingBox();
  if (!menuBox) throw new Error("missing menu box");

  expect(menuBox.x + menuBox.width).toBeLessThanOrEqual(appBox.x + appBox.width + 1); // +1: sub-pixel rounding
  expect(menuBox.x).toBeGreaterThanOrEqual(appBox.x - 1);
});

test("a menu item that would have been clipped is still fully clickable after the clamp", async ({ page }) => {
  await gotoFixture(page);
  const appBox = await page.locator("#app").boundingBox();
  if (!appBox) throw new Error("missing app box");

  expect(await getCell(page, 0, 2)).toBe("0-2");

  await page.mouse.click(appBox.x + 230, appBox.y + 10, { button: "right" }); // column 2
  await page.locator('[data-item-id="cyexcel:insert-column-left"]').click();

  // The insert genuinely ran (not silently swallowed by a clipped/
  // unclickable menu item) — column 2's old data shifted to column 3.
  expect(await getCell(page, 0, 2)).toBeUndefined();
  expect(await getCell(page, 0, 3)).toBe("0-2");
});
