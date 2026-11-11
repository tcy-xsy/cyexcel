import { expect, test } from "@playwright/test";
import type { WorkbookHandle } from "../../src/createWorkbook";

declare global {
  interface Window {
    workbookForExcelImportTest: WorkbookHandle;
    createExcelImportFixture(): Promise<number[]>;
  }
}

test.beforeEach(async ({ page }) => {
  await page.goto("/tests/browser/excel-import.html");
  await page.waitForFunction(() => Boolean(window.workbookForExcelImportTest));
});

test("loads the parser only on import, replaces tabs and preserves formulas, styles and merges", async ({ page }) => {
  const before = await page.evaluate(() => performance.getEntriesByType("resource").map(entry => entry.name));
  expect(before.some(name => /exceljs/i.test(name))).toBe(false);
  const bytes = await page.evaluate(() => window.createExcelImportFixture());
  await page.getByRole("button", { name: "导入 Excel", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "导入 Excel", exact: true });
  await expect(dialog.getByRole("button", { name: "导入并替换" })).toBeDisabled();
  await dialog.getByLabel("Excel 文件").setInputFiles({ name: "workbook.xlsx", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", buffer: Buffer.from(bytes) });
  await expect(dialog.getByRole("button", { name: "导入并替换" })).toBeDisabled();
  await dialog.getByLabel("替换当前工作簿", { exact: true }).check();
  await dialog.getByRole("button", { name: "导入并替换" }).click();
  await expect(dialog).toHaveCount(0);
  await expect(page.getByRole("tab")).toHaveText(["Imported sales", "Imported stock"]);
  const result = await page.evaluate(() => {
    const core = window.workbookForExcelImportTest.getActiveSheet();
    const snapshot = core.toSnapshot();
    const sheet = snapshot.sheets[snapshot.sheetOrder[0]!]!;
    return { text: core.getCell(0, 0), formula: sheet.cellData[1]![1]!.f, value: core.getCell(1, 1), style: core.getCellStyle(1, 1), merges: sheet.mergeData };
  });
  expect(result).toMatchObject({ text: "00123", formula: "=B1*2", value: 8, style: { bold: true, color: "#ff8000", numFmt: "0.00" }, merges: [{ startRow: 2, startColumn: 0, endRow: 2, endColumn: 1 }] });
  await expect(page.locator(".cyx-workbook-sheet")).toHaveCount(1);
  await page.getByRole("tab", { name: "Imported stock", exact: true }).click();
  expect(await page.evaluate(() => window.workbookForExcelImportTest.getActiveSheet().getCell(0, 0))).toBe("stock");
});

test("a corrupt file reports an inline error and retains the current workbook", async ({ page }) => {
  const errors: string[] = [];
  page.on("pageerror", error => errors.push(error.message));
  await page.getByRole("button", { name: "导入 Excel", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "导入 Excel", exact: true });
  await dialog.getByLabel("Excel 文件").setInputFiles({ name: "broken.xlsx", mimeType: "application/octet-stream", buffer: Buffer.from("broken") });
  await dialog.getByLabel("替换当前工作簿", { exact: true }).check();
  await dialog.getByRole("button", { name: "导入并替换" }).click();
  await expect(dialog.getByRole("status")).toContainText("导入失败");
  await expect(page.getByRole("tab")).toHaveText(["Original"]);
  expect(await page.evaluate(() => window.workbookForExcelImportTest.getActiveSheet().getCell(0, 0))).toBe("original");
  expect(errors).toEqual([]);
  await dialog.getByRole("button", { name: "关闭", exact: true }).click();
  await expect(dialog).toHaveCount(0);
});

test("replacement validates all snapshots before changing existing tabs", async ({ page }) => {
  const error = await page.evaluate(() => {
    const workbook = window.workbookForExcelImportTest;
    const snapshot = workbook.getActiveSheet().toSnapshot();
    try {
      workbook.replaceSheets([{ name: "Valid", snapshot }, { name: "Invalid", snapshot: { ...snapshot, version: -1 } }]);
    } catch (error) {
      return String(error);
    }
    return null;
  });
  expect(error).toContain("Unsupported snapshot version");
  await expect(page.getByRole("tab")).toHaveText(["Original"]);
  await expect(page.locator(".cyx-workbook-sheet")).toHaveCount(1);
});
