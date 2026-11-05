import { expect, test } from "@playwright/test";
import type { WorkbookHandle } from "../../src/createWorkbook";

declare global {
  interface Window {
    workbookForTest: WorkbookHandle;
  }
}

test.beforeEach(async ({ page }) => {
  await page.goto("/tests/browser/workbook-delete.html");
  await expect(page.getByRole("tab")).toHaveCount(3);
});

test("deletes an inactive unmounted sheet and guards the final sheet", async ({ page }) => {
  const result = await page.evaluate(() => {
    const workbook = window.workbookForTest;
    const initial = workbook.getActiveSheet();
    initial.setCell(0, 0, "preserved");
    workbook.deleteSheet("sheet-002");
    const sameActive = initial === workbook.getActiveSheet();
    workbook.deleteSheet("sheet-003");
    const errors: string[] = [];
    for (const id of ["sheet-001", "missing", "sheet-002"]) {
      try {
        workbook.deleteSheet(id);
      } catch (error) {
        errors.push((error as Error).message);
      }
    }
    return { sameActive, value: initial.getCell(0, 0), errors };
  });
  expect(result).toEqual({
    sameActive: true,
    value: "preserved",
    errors: ["cannot delete the last sheet", 'unknown sheet "missing"', 'unknown sheet "sheet-002"']
  });
  await expect(page.getByRole("tab")).toHaveCount(1);
  await expect(page.getByRole("button", { name: "删除工作表 First", exact: true })).toBeDisabled();
  await page.getByRole("button", { name: "添加工作表", exact: true }).click();
  await expect(page.locator(".cyx-workbook-close-sheet:disabled")).toHaveCount(0);
});

test("deletes active sheets, selects neighbors and releases mounted hosts", async ({ page }) => {
  await page.getByRole("tab", { name: "Second", exact: true }).click();
  await page.evaluate(() => window.workbookForTest.deleteSheet("sheet-002"));
  await expect(page.getByRole("tab", { name: "Third", exact: true })).toHaveAttribute("aria-selected", "true");
  await page.evaluate(() => window.workbookForTest.deleteSheet("sheet-003"));
  await expect(page.getByRole("tab", { name: "First", exact: true })).toHaveAttribute("aria-selected", "true");
  await expect(page.locator(".cyx-workbook-sheet")).toHaveCount(1);
  await page.evaluate(() => window.workbookForTest.addSheet({ name: "Fourth" }));
  await expect(page.getByRole("tab", { name: "Fourth", exact: true })).toHaveAttribute("aria-selected", "true");
});

test("close icons support cancellation and confirmation", async ({ page }) => {
  const close = page.getByRole("button", { name: "删除工作表 First", exact: true });
  page.once("dialog", dialog => dialog.dismiss());
  await close.click();
  await expect(page.getByRole("tab")).toHaveCount(3);
  page.once("dialog", dialog => dialog.accept());
  await close.click();
  await expect(page.getByRole("tab")).toHaveCount(2);
  await expect(page.getByRole("tab", { name: "Second", exact: true })).toHaveAttribute("aria-selected", "true");
});

test("closing an inactive sheet keeps the current tab and supports keyboard activation", async ({ page }) => {
  const close = page.getByRole("button", { name: "删除工作表 Third", exact: true });
  await close.focus();
  page.once("dialog", dialog => dialog.accept());
  await page.keyboard.press("Enter");
  await expect(page.getByRole("tab")).toHaveCount(2);
  await expect(page.getByRole("tab", { name: "First", exact: true })).toHaveAttribute("aria-selected", "true");
  await expect(page.getByRole("tab", { name: "First", exact: true })).toBeFocused();
  await page.getByRole("tab", { name: "First", exact: true }).click({ button: "right" });
  await expect(page.locator(".cyx-workbook-tab-menu")).toHaveCount(0);
  await page.evaluate(() => window.workbookForTest.dispose());
  await expect(page.locator(".cyx-workbook")).toHaveCount(0);
});

test("renames inline with Enter, blur, Escape and blank-name protection", async ({ page }) => {
  const nameInput = page.getByRole("textbox", { name: "工作表名称", exact: true });
  await page.getByRole("tab", { name: "First", exact: true }).dblclick();
  await expect(nameInput).toHaveValue("First");
  await nameInput.fill("  Renamed  ");
  await nameInput.press("Enter");
  const renamed = page.getByRole("tab", { name: "Renamed", exact: true });
  await expect(renamed).toBeFocused();
  await expect(renamed).toHaveAttribute("aria-selected", "true");
  await renamed.press("F2");
  await nameInput.fill("Discarded");
  await nameInput.press("Escape");
  await expect(renamed).toBeVisible();
  await renamed.dblclick();
  await nameInput.fill("   ");
  await nameInput.press("Enter");
  await expect(renamed).toBeVisible();
  await renamed.dblclick();
  await nameInput.fill("Saved on blur");
  await page.getByRole("tab", { name: "Second", exact: true }).click();
  await expect(page.getByRole("tab", { name: "Saved on blur", exact: true })).toBeVisible();
  await expect(page.getByRole("tab", { name: "Second", exact: true })).toHaveAttribute("aria-selected", "true");
  page.once("dialog", async dialog => {
    expect(dialog.message()).toContain("Saved on blur");
    await dialog.dismiss();
  });
  await page.getByRole("button", { name: "删除工作表 Saved on blur", exact: true }).click();
  await expect(page.getByRole("tab")).toHaveCount(3);
});

test("renames unmounted sheets through the API and preserves the name when switching", async ({ page }) => {
  const errors = await page.evaluate(() => {
    const workbook = window.workbookForTest;
    workbook.renameSheet("sheet-003", "  API name  ");
    const errors: string[] = [];
    for (const [id, name] of [["missing", "New"], ["sheet-003", "   "]]) {
      try { workbook.renameSheet(id!, name!); }
      catch (error) { errors.push((error as Error).message); }
    }
    return errors;
  });
  expect(errors).toEqual(['unknown sheet "missing"', "sheet name cannot be blank"]);
  await page.getByRole("tab", { name: "API name", exact: true }).click();
  await expect(page.getByRole("tab", { name: "API name", exact: true })).toHaveAttribute("aria-selected", "true");
});
