import { expect, test } from "@playwright/test";
import type { WorkbookHandle } from "../../src/createWorkbook.js";

declare global {
  interface Window {
    localeFixture: { workbook: WorkbookHandle };
  }
}

test.beforeEach(async ({ page }) => {
  await page.setViewportSize({ width: 2800, height: 1000 });
  await page.goto("/tests/browser/locale.html");
  await expect(page.locator("#english .cyx-toolbar")).toBeVisible();
});

test("English, Chinese and custom locales coexist, including workbook controls", async ({ page }) => {
  await expect(page.locator("#english").getByRole("button", { name: "Undo", exact: true })).toBeVisible();
  await expect(page.locator("#chinese").getByRole("button", { name: "撤销", exact: true })).toBeVisible();
  await expect(page.locator("#custom").getByRole("button", { name: "Annuler", exact: true })).toBeVisible();
  await expect(page.locator("#custom").getByRole("button", { name: "Redo", exact: true })).toBeVisible();
  await expect(page.locator("#english").getByLabel("Cell address", { exact: true })).toBeVisible();
  await page.locator("#english").getByRole("button", { name: "Add sheet", exact: true }).click();
  await expect(page.locator("#english")).toContainText("Sheet2");
  await expect(page.locator("#english").getByRole("button", { name: "Delete sheet Sheet2", exact: true })).toBeVisible();
});

test("workbook snapshots custom dictionaries for added, lazy and replacement sheets", async ({ page }) => {
  const labels = await page.evaluate(async () => {
    const modulePath = "/src/createWorkbook.ts";
    const { createWorkbook } = await import(modulePath);
    const host = document.createElement("div");
    host.style.cssText = "width:1200px;height:320px";
    document.body.append(host);
    const messages = { Undo: "First undo" };
    const lazyMessages = { Undo: "Lazy undo" };
    const workbook = createWorkbook(host, {
      locale: { name: "custom", messages },
      sheets: [{}, { locale: { name: "lazy", messages: lazyMessages } }]
    });
    const result = [workbook.getActiveSheet().t("Undo")];
    messages.Undo = "Changed undo";
    lazyMessages.Undo = "Changed lazy undo";
    workbook.addSheet();
    result.push(workbook.getActiveSheet().t("Undo"));
    workbook.setActiveSheet("sheet-002");
    result.push(workbook.getActiveSheet().t("Undo"));
    const replacementMessages = { Undo: "Replacement undo" };
    const ids = workbook.replaceSheets([{}, { locale: { name: "replacement", messages: replacementMessages } }]);
    replacementMessages.Undo = "Changed replacement undo";
    workbook.setActiveSheet(ids[1]);
    result.push(workbook.getActiveSheet().t("Undo"));
    workbook.dispose();
    host.remove();
    return result;
  });
  expect(labels).toEqual(["First undo", "First undo", "Lazy undo", "Replacement undo"]);
});

test("context menu and search dialog use English", async ({ page }) => {
  await page.locator("#english canvas").first().click({ button: "right", position: { x: 30, y: 14 } });
  await expect(page.getByRole("menu", { name: "Spreadsheet actions", exact: true })).toBeVisible();
  await expect(page.getByRole("menuitem", { name: /Copy/ })).toBeVisible();
  await page.keyboard.press("Escape");
  await page.locator("#english").getByRole("button", { name: "Find and replace", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "Find and replace", exact: true });
  await expect(dialog.getByRole("button", { name: "Find all", exact: true })).toBeVisible();
  await dialog.getByRole("button", { name: "Find all", exact: true }).click();
  await expect(dialog).toContainText("Enter a search term first.");
  await dialog.getByRole("button", { name: "Close", exact: true }).first().click();
  await expect(dialog).toHaveCount(0);
});

test("data validation options and form errors use English", async ({ page }) => {
  await page.locator("#english").getByRole("button", { name: "Data validation", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "Data validation", exact: true });
  await dialog.getByRole("combobox", { name: "Allow", exact: true }).selectOption("whole");
  await dialog.getByRole("combobox", { name: "Data", exact: true }).selectOption("equal");
  await dialog.getByRole("button", { name: "OK", exact: true }).click();
  await expect(dialog).toContainText("Enter Value.");
  await expect(dialog.getByRole("combobox", { name: "Error style", exact: true })).toContainText("Stop (reject input)");
});

test("validation rejection dialogs keep English labels and the original cell value", async ({ page }) => {
  await page.evaluate(() => window.localeFixture.workbook.getActiveSheet().setDataValidation(
    { startRow: 0, endRow: 0, startColumn: 1, endColumn: 1 },
    { criteria: { type: "whole", operator: "between", value1: 1, value2: 10 } }
  ));
  await page.locator("#english canvas").first().dblclick({ position: { x: 150, y: 14 } });
  await page.locator('#english input[style*="position: absolute"]').fill("100");
  await page.keyboard.press("Enter");
  const dialog = page.getByRole("alertdialog");
  await expect(dialog).toContainText("Whole number must be between 1 and 10");
  await expect(dialog).toContainText("Your changes were not saved.");
  await page.setViewportSize({ width: 320, height: 480 });
  await expect.poll(async () => {
    const box = await dialog.boundingBox();
    return box!.x >= 0 && box!.x + box!.width <= 320;
  }).toBe(true);
  await dialog.getByRole("button", { name: "OK", exact: true }).click();
  expect(await page.evaluate(() => window.localeFixture.workbook.getActiveSheet().getCell(0, 1))).toBe(5);
});

test("color panels and toolbar overflow use English on narrow screens", async ({ page }) => {
  await page.locator("#english").getByRole("button", { name: "Text color (more)", exact: true }).click();
  await expect(page.getByRole("button", { name: "Red #dc2626", exact: true })).toBeVisible();
  await expect(page.getByLabel("Open color picker", { exact: true })).toBeVisible();
  await page.keyboard.press("Escape");
  await page.setViewportSize({ width: 420, height: 900 });
  await page.locator("#english").evaluate(element => { element.style.width = "380px"; });
  await page.locator("#english").getByRole("button", { name: "More", exact: true }).click();
  await expect(page.locator(".cyx-toolbar-panel").getByRole("button", { name: "Find and replace", exact: true })).toBeVisible();
  await page.locator(".cyx-toolbar-panel").getByRole("button", { name: "Find and replace", exact: true }).click();
  await expect(page.getByRole("dialog", { name: "Find and replace", exact: true })).toBeVisible();
});

test("English forms keep aligned labels and fit small viewports after resizing", async ({ page }) => {
  await page.locator("#english").getByRole("button", { name: "Data validation", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "Data validation", exact: true });
  await dialog.getByRole("combobox", { name: "Allow", exact: true }).selectOption("list");
  await expect(dialog).toContainText("Comma-separated, with optional colors:");
  const allow = await dialog.getByRole("combobox", { name: "Allow", exact: true }).boundingBox();
  const source = await dialog.getByRole("textbox", { name: "Value", exact: true }).boundingBox();
  expect(source!.x).toBeCloseTo(allow!.x, 0);
  expect(source!.width).toBeCloseTo(allow!.width, 0);

  for (const width of [375, 320]) {
    await page.setViewportSize({ width, height: 480 });
    await expect.poll(async () => {
      const box = await dialog.boundingBox();
      return box!.x >= 0 && box!.x + box!.width <= width && box!.y >= 0 && box!.y + box!.height <= 480;
    }).toBe(true);
    const overflowing = await dialog.evaluate(root => [...root.querySelectorAll("input, select, button, label")]
      .filter(element => element.getClientRects().length > 0)
      .some(element => {
        const box = element.getBoundingClientRect();
        const bounds = root.getBoundingClientRect();
        return box.left < bounds.left || box.right > bounds.right;
      }));
    expect(overflowing).toBe(false);
    await dialog.getByRole("button", { name: "OK", exact: true }).scrollIntoViewIfNeeded();
    await expect(dialog.getByRole("button", { name: "OK", exact: true })).toBeInViewport();
  }
});

test("English search actions remain reachable on a short mobile screen", async ({ page }) => {
  await page.locator("#english").getByRole("button", { name: "Find and replace", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "Find and replace", exact: true });
  await page.setViewportSize({ width: 320, height: 320 });
  await expect.poll(async () => {
    const box = await dialog.boundingBox();
    return box!.x >= 0 && box!.x + box!.width <= 320 && box!.y >= 0 && box!.y + box!.height <= 320;
  }).toBe(true);
  await dialog.getByRole("button", { name: "Replace all", exact: true }).click();
  await expect(dialog).toContainText("Enter a search term first.");
  const body = dialog.locator(".cyx-toolbar-dialog-body");
  expect(await body.evaluate(element => element.scrollWidth <= element.clientWidth)).toBe(true);
});

test("search dialog sizes to its content and keeps actions in two aligned rows", async ({ page }) => {
  await page.addStyleTag({ content: 'body { font-family: "Helvetica Neue", sans-serif; }' });
  await page.locator("#english").getByRole("button", { name: "Find and replace", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "Find and replace", exact: true });
  const buttons = dialog.locator(".cyx-toolbar-dialog-buttons button");
  const initial = (await dialog.boundingBox())!.width;
  await buttons.first().evaluate(element => { element.textContent = "Find all matching spreadsheet cells"; });
  await expect.poll(async () => (await dialog.boundingBox())!.width).toBeGreaterThan(initial);
  for (const width of [1440, 375, 320]) {
    await page.setViewportSize({ width, height: 800 });
    await expect.poll(async () => {
      const bounds = await dialog.boundingBox();
      return bounds!.x >= 0 && bounds!.x + bounds!.width <= width;
    }).toBe(true);
    const boxes = await buttons.evaluateAll(elements => elements.map(element => {
      const { x, y, width, height } = element.getBoundingClientRect();
      return { x, y, width, height, fits: element.scrollWidth <= element.clientWidth };
    }));
    for (let column = 0; column < 3; column++) {
      expect(boxes[column]!.x).toBeCloseTo(boxes[column + 3]!.x, 0);
      expect(boxes[column]!.y).toBe(boxes[0]!.y);
      expect(boxes[column + 3]!.y).toBe(boxes[3]!.y);
    }
    expect(boxes.every(box => box.fits)).toBe(true);
    expect(boxes[3]!.y).toBeGreaterThan(boxes[0]!.y);
  }
});

test("dialog can be enlarged from its resize corner and still fits after viewport changes", async ({ page }) => {
  await page.locator("#english").getByRole("button", { name: "Find and replace", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "Find and replace", exact: true });
  const initial = (await dialog.boundingBox())!;
  expect(initial.width).toBeGreaterThanOrEqual(480);
  await page.mouse.move(initial.x + initial.width - 3, initial.y + initial.height - 3);
  await page.mouse.down();
  await page.mouse.move(initial.x + initial.width + 180, initial.y + initial.height + 100, { steps: 10 });
  await page.mouse.up();
  const enlarged = (await dialog.boundingBox())!;
  expect(enlarged.width).toBeGreaterThan(initial.width + 150);
  expect(enlarged.height).toBeGreaterThan(initial.height + 70);
  await expect(dialog.getByRole("textbox", { name: "Find what", exact: true })).toHaveJSProperty("value", "");
  await page.setViewportSize({ width: 320, height: 400 });
  await expect.poll(async () => {
    const bounds = (await dialog.boundingBox())!;
    return bounds.x >= 0 && bounds.y >= 0 && bounds.x + bounds.width <= 320 && bounds.y + bounds.height <= 400;
  }).toBe(true);
  await dialog.getByRole("button", { name: "Replace all", exact: true }).click();
  await expect(dialog).toContainText("Enter a search term first.");
});

test("long English dropdown items stay inside the sheet and remain selectable", async ({ page }) => {
  const errors: string[] = [];
  page.on("pageerror", error => errors.push(error.message));
  await page.setViewportSize({ width: 320, height: 800 });
  await page.locator("#english").evaluate(element => { element.style.width = "296px"; });
  const items = ["Awaiting manager approval before shipment", "InternationalDeliveryTrackingReferenceWithoutSpaces", "Completed"] as const;
  await page.evaluate(items => window.localeFixture.workbook.getActiveSheet().setDataValidation(
    { startRow: 0, endRow: 0, startColumn: 1, endColumn: 1 },
    { criteria: { type: "list", items } }
  ), items);
  const canvas = page.locator("#english canvas").first();
  await canvas.click({ position: { x: 230, y: 14 } });
  const list = page.getByRole("listbox");
  const sheetBounds = (await canvas.boundingBox())!;
  const listBounds = (await list.boundingBox())!;
  expect(listBounds.x).toBeGreaterThanOrEqual(sheetBounds.x);
  expect(listBounds.x + listBounds.width).toBeLessThanOrEqual(sheetBounds.x + sheetBounds.width);
  expect(listBounds.y + listBounds.height).toBeLessThanOrEqual(sheetBounds.y + sheetBounds.height);
  expect(await list.evaluate(element => element.scrollWidth <= element.clientWidth)).toBe(true);
  await list.getByRole("option", { name: items[2], exact: true }).click();
  expect(await page.evaluate(() => window.localeFixture.workbook.getActiveSheet().getCell(0, 1))).toBe(items[2]);
  expect(errors).toEqual([]);
});

test("overflow select labels have their own line and long tooltips fit the viewport", async ({ page }) => {
  await page.setViewportSize({ width: 320, height: 800 });
  await page.locator("#english").evaluate(element => { element.style.width = "296px"; });
  await page.locator("#english").getByRole("button", { name: "More", exact: true }).click();
  const panel = page.locator(".cyx-toolbar-panel");
  const row = panel.locator(".cyx-toolbar-panel-row").filter({ has: page.getByRole("combobox", { name: "Number format", exact: true }) });
  const label = await row.evaluate(element => {
    const range = document.createRange();
    range.selectNodeContents(element.firstChild!);
    return { lines: range.getClientRects().length, bottom: range.getBoundingClientRect().bottom };
  });
  expect(label.lines).toBe(1);
  expect((await row.getByRole("combobox").boundingBox())!.y).toBeGreaterThanOrEqual(label.bottom);
  expect(await panel.evaluate(element => element.scrollWidth <= element.clientWidth)).toBe(true);
  await page.keyboard.press("Escape");
  // Use the public schema API to exercise an integration's longer English tooltip.
  await page.evaluate(async () => {
    const modulePath = "/src/createSheet.ts";
    const { createSheet } = await import(modulePath);
    const host = document.createElement("div");
    host.style.cssText = "position:fixed;bottom:0;left:0;width:280px;height:100px";
    document.body.append(host);
    createSheet(host, {
      locale: "en-US",
      toolbar: { items: [{ id: "long-help", group: "history", order: 0, type: "button", label: "Help", tooltip: "Formatting and spreadsheet actions for the currently selected cells, including number formats, borders, and data validation rules." }] }
    });
  });
  await page.getByRole("button", { name: "Help", exact: true }).hover();
  const tooltip = page.locator(".cyx-toolbar-tooltip");
  await expect(tooltip).toBeVisible();
  const bounds = (await tooltip.boundingBox())!;
  expect(bounds.x).toBeGreaterThanOrEqual(0);
  expect(bounds.x + bounds.width).toBeLessThanOrEqual(320);
  expect(bounds.y).toBeGreaterThanOrEqual(0);
  expect(bounds.y + bounds.height).toBeLessThanOrEqual(800);
});
