import { expect, test, type Page } from "@playwright/test";
import type { CyExcelCore } from "../../src/core";

declare global {
  interface Window {
    validationSheet: CyExcelCore;
  }
}

const editor = (page: Page) => page.locator('#app input[style*="position: absolute"]');

async function clickCell(page: Page, column: number, row = 0, edit = false): Promise<void> {
  const box = await page.locator("#app canvas").first().boundingBox();
  if (!box) {
    throw new Error("missing body canvas");
  }
  const x = box.x + [75, 195, 300, 410][column]!;
  const y = box.y + row * 30 + 15;
  if (edit) {
    await page.mouse.dblclick(x, y);
  } else {
    await page.mouse.click(x, y);
  }
}

test.beforeEach(async ({ page }) => {
  await page.goto("/tests/browser/data-validation.html");
  await page.waitForFunction(() => Boolean(window.validationSheet));
});

test("acknowledging a rejected edit restores the value and lets other cells be selected", async ({ page }) => {
  await clickCell(page, 1, 0, true);
  await editor(page).fill("500");
  await page.keyboard.press("Enter");
  await expect(page.getByRole("alertdialog")).toHaveCount(1);
  await page.getByRole("button", { name: "确定", exact: true }).click();
  await expect(editor(page)).toHaveCount(0);
  await clickCell(page, 0);
  await expect(page.getByRole("alertdialog")).toHaveCount(0);
  expect(await page.evaluate(() => window.validationSheet.getCell(0, 1))).toBe(8);
  expect(await page.evaluate(() => window.validationSheet.selection.active)).toEqual({ row: 0, column: 0 });
  await clickCell(page, 1, 0, true);
  await editor(page).fill("42");
  await page.keyboard.press("Enter");
  await expect(editor(page)).toHaveCount(0);
  expect(await page.evaluate(() => window.validationSheet.getCell(0, 1))).toBe(42);
});

test("click-away rejection and Escape abandon the draft without reopening", async ({ page }) => {
  await clickCell(page, 1, 0, true);
  await editor(page).fill("abc");
  await clickCell(page, 0);
  await expect(page.getByRole("alertdialog")).toHaveCount(1);
  await page.keyboard.press("Escape");
  await expect(editor(page)).toHaveCount(0);
  await clickCell(page, 0);
  await expect(page.getByRole("alertdialog")).toHaveCount(0);
  expect(await page.evaluate(() => window.validationSheet.getCell(0, 1))).toBe(8);
});

test("an existing invalid value can be left after acknowledging, then corrected", async ({ page }) => {
  await clickCell(page, 1, 1, true);
  await page.keyboard.press("Enter");
  await expect(page.getByRole("alertdialog")).toHaveCount(1);
  await page.getByRole("button", { name: "确定", exact: true }).click();
  await clickCell(page, 0);
  await expect(page.getByRole("alertdialog")).toHaveCount(0);
  expect(await page.evaluate(() => window.validationSheet.getCell(1, 1))).toBe(250);
  await clickCell(page, 1, 1, true);
  await editor(page).fill("12");
  await page.keyboard.press("Enter");
  await expect(editor(page)).toHaveCount(0);
  expect(await page.evaluate(() => window.validationSheet.validateCell(1, 1).valid)).toBe(true);
});

test("valid numbers round-trip and warning values commit without a dialog", async ({ page }) => {
  await clickCell(page, 1, 0, true);
  await page.keyboard.press("Enter");
  await expect(editor(page)).toHaveCount(0);
  await expect(page.getByRole("alertdialog")).toHaveCount(0);
  expect(await page.evaluate(() => window.validationSheet.getCell(0, 1))).toBe(8);
  await clickCell(page, 3, 0, true);
  await editor(page).fill("a longer note");
  await page.keyboard.press("Enter");
  await expect(editor(page)).toHaveCount(0);
  await expect(page.getByRole("alertdialog")).toHaveCount(0);
  expect(await page.evaluate(() => window.validationSheet.getCell(0, 3))).toBe("a longer note");
  expect(await page.evaluate(() => window.validationSheet.validateCell(0, 3).valid)).toBe(false);
});

test("hover hints distinguish severity, fit inside the sheet and do not block editing", async ({ page }) => {
  await page.evaluate(() => {
    window.validationSheet.setCell(0, 3, "long warning note");
    window.validationSheet.setCell(1, 3, "long information note");
  });
  const box = await page.locator("#app canvas").first().boundingBox();
  if (!box) {
    throw new Error("missing body canvas");
  }
  const hint = page.getByRole("tooltip");
  for (const entry of [
    { x: 195, y: 45, severity: "stop", title: "输入无效" },
    { x: 410, y: 15, severity: "warning", title: "注意" },
    { x: 410, y: 45, severity: "information", title: "提示" }
  ]) {
    await page.mouse.move(box.x + entry.x, box.y + entry.y);
    await expect(hint).toHaveAttribute("data-severity", entry.severity);
    await expect(hint).toContainText(entry.title);
    await expect(hint).toHaveCSS("pointer-events", "none");
    const hintBox = await hint.boundingBox();
    expect(hintBox).not.toBeNull();
    expect(hintBox!.width).toBe(240);
    expect(hintBox!.height).toBeLessThan(140);
    expect(hintBox!.x).toBeGreaterThanOrEqual(box.x + Math.min(entry.x - 45, box.width - 244));
    expect(hintBox!.x + hintBox!.width).toBeLessThanOrEqual(box.x + box.width);
  }
  await expect(hint).not.toContainText("错误");
  await clickCell(page, 3, 1, true);
  await expect(editor(page)).toHaveCount(1);
  await expect(hint).toHaveCount(0);
  await page.keyboard.press("Escape");
  await page.mouse.move(1, 1);
  await expect(hint).toHaveCount(0);
});
