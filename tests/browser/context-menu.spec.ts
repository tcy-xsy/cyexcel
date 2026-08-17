import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

interface CellActionRun {
  readonly target: string;
  readonly row?: number;
  readonly column?: number;
}

interface FixtureWindow {
  __cyexcelContextMenuReady?: boolean;
  __cellActionRuns?: CellActionRun[];
  __sheet?: {
    getCell(row: number, column: number): unknown;
    setCell(row: number, column: number, value: unknown): void;
  };
}

async function gotoFixture(page: Page): Promise<void> {
  await page.goto("/tests/browser/context-menu.html");
  await page.waitForFunction(
    () => (window as unknown as FixtureWindow).__cyexcelContextMenuReady === true
  );
}

async function getAppBox(page: Page) {
  const box = await page.locator("#app").boundingBox();
  if (!box) {
    throw new Error("missing #app bounding box");
  }
  return box;
}

function getRuns(page: Page): Promise<CellActionRun[]> {
  return page.evaluate(() => (window as unknown as FixtureWindow).__cellActionRuns ?? []);
}

test("right-clicking a cell opens a menu with a cell-target action", async ({ page }) => {
  await gotoFixture(page);
  const box = await getAppBox(page);

  await page.mouse.click(box.x + 50, box.y + 36, { button: "right" });

  await expect(page.locator('[role="menu"]')).toBeVisible();
  await expect(page.locator('[data-item-id="log-cell"]')).toBeVisible();
});

test("keyboard navigation runs the highlighted action and returns focus to the content canvas", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getAppBox(page);

  await page.mouse.click(box.x + 50, box.y + 36, { button: "right" }); // cell (0, 0)
  // The cell target's merged menu is [cyexcel:copy, cyexcel:paste, log-cell]
  // (basicMenuPlugin's provider registers before the fixture's own
  // "cell-actions" provider) — three ArrowDown presses reaches "log-cell",
  // which also exercises cycling through multiple items.
  await page.keyboard.press("ArrowDown");
  await page.keyboard.press("ArrowDown");
  await page.keyboard.press("ArrowDown");
  await page.keyboard.press("Enter");

  await expect(page.locator('[role="menu"]')).toHaveCount(0);
  const runs = await getRuns(page);
  expect(runs).toEqual([{ target: "cell", row: 0, column: 0 }]);

  const isCanvasFocused = await page.evaluate(() => {
    const canvas = document.querySelectorAll("#app canvas")[0];
    return document.activeElement === canvas;
  });
  expect(isCanvasFocused).toBe(true);
});

test("Escape closes the menu without running an action, and returns focus to the content canvas", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getAppBox(page);

  await page.mouse.click(box.x + 50, box.y + 36, { button: "right" });
  await page.keyboard.press("ArrowDown");
  await page.keyboard.press("Escape");

  await expect(page.locator('[role="menu"]')).toHaveCount(0);
  expect(await getRuns(page)).toEqual([]);

  const isCanvasFocused = await page.evaluate(() => {
    const canvas = document.querySelectorAll("#app canvas")[0];
    return document.activeElement === canvas;
  });
  expect(isCanvasFocused).toBe(true);
});

test("right-clicking the column header band opens a menu with a column-header-target action", async ({
  page
}) => {
  await gotoFixture(page);
  const box = await getAppBox(page);

  // y < headerHeight (24px) — the header band, not the body.
  await page.mouse.click(box.x + 150, box.y + 10, { button: "right" }); // "age" column header

  // basicMenuPlugin now also contributes items for "column-header" (Step
  // 17), so this fixture's own item is no longer first — click it directly
  // by id instead of counting ArrowDown presses through the merged menu.
  await expect(page.locator('[data-item-id="log-column-header"]')).toBeVisible();
  await page.locator('[data-item-id="log-column-header"]').click();

  expect(await getRuns(page)).toEqual([{ target: "column-header", column: 1 }]);
});

test("right-clicking past the data extent opens a menu with a blank-target action", async ({ page }) => {
  await gotoFixture(page);
  const box = await getAppBox(page);

  // Well below the 2 data rows (headerHeight 24 + 2 * rowHeight 24 = 72px of content).
  await page.mouse.click(box.x + 50, box.y + 250, { button: "right" });

  await expect(page.locator('[data-item-id="log-blank"]')).toBeVisible();
  await page.keyboard.press("ArrowDown");
  await page.keyboard.press("Enter");

  expect(await getRuns(page)).toEqual([{ target: "blank" }]);
});

test("right-clicking a cell inside an existing multi-cell selection preserves it instead of collapsing to one cell", async ({
  page,
  context
}) => {
  await context.grantPermissions(["clipboard-read", "clipboard-write"]);
  await gotoFixture(page);
  const box = await getAppBox(page);

  // Drag-select cell (0,0) "Alice" through cell (1,0) "Bob".
  await page.mouse.move(box.x + 50, box.y + 36);
  await page.mouse.down();
  await page.mouse.move(box.x + 50, box.y + 36 + 24);
  await page.mouse.up();

  // Right-click one of the two already-selected cells.
  await page.mouse.click(box.x + 50, box.y + 36, { button: "right" });
  await page.locator('[data-item-id="cyexcel:copy"]').click();
  await expect(page.locator('[role="menu"]')).toHaveCount(0);

  await expect
    .poll(async () => {
      const text = await page.evaluate(() => navigator.clipboard.readText());
      return text.split("\n").map((line) => line.trim());
    })
    .toEqual(["Alice", "Bob"]); // both rows — the selection wasn't collapsed to just "Alice"
});

test("clicking outside an open context menu closes it without running any action", async ({ page }) => {
  await gotoFixture(page);
  const box = await getAppBox(page);

  await page.mouse.click(box.x + 50, box.y + 36, { button: "right" });
  await expect(page.locator('[role="menu"]')).toBeVisible();

  // A plain left click elsewhere on the grid, well away from the menu.
  await page.mouse.click(box.x + 50, box.y + 250);

  await expect(page.locator('[role="menu"]')).toHaveCount(0);
  expect(await getRuns(page)).toEqual([]);
});

test("createBasicContextMenuPlugin's Copy/Paste actions round-trip through the real system clipboard", async ({
  page,
  context
}) => {
  await context.grantPermissions(["clipboard-read", "clipboard-write"]);
  await gotoFixture(page);
  const box = await getAppBox(page);

  await page.mouse.click(box.x + 50, box.y + 36, { button: "right" }); // cell (0, 0), value "Alice"
  await page.locator('[data-item-id="cyexcel:copy"]').click();
  await expect(page.locator('[role="menu"]')).toHaveCount(0);

  const copiedText = await page.evaluate(() => navigator.clipboard.readText());
  expect(copiedText).toBe("Alice");

  await page.evaluate(async () => {
    await navigator.clipboard.writeText("Zed");
  });
  await page.mouse.click(box.x + 50, box.y + 36 + 24, { button: "right" }); // cell (1, 0), value "Bob"
  await page.locator('[data-item-id="cyexcel:paste"]').click();
  await expect(page.locator('[role="menu"]')).toHaveCount(0);

  await expect
    .poll(() => page.evaluate(() => (window as unknown as FixtureWindow).__sheet?.getCell(1, 0)))
    .toBe("Zed");
});
