import { expect, test } from "@playwright/test";
import { readFile } from "node:fs/promises";
import type { WorkbookHandle } from "../../src/createWorkbook";
import { readXlsxEntries } from "../support/xlsx";

declare global {
  interface Window {
    exportWorkbookForTest: WorkbookHandle;
  }
}

test.beforeEach(async ({ page }) => {
  await page.goto("/tests/browser/excel-export.html");
  await page.waitForFunction(() => Boolean(window.exportWorkbookForTest));
});

test("toolbar exports renamed sheets, unopened data, styles, merges and validation rules", async ({ page }, testInfo) => {
  const pending = page.waitForEvent("download");
  await page.getByRole("button", { name: "导出 Excel", exact: true }).click();
  const download = await pending;
  expect(download.suggestedFilename()).toMatch(/\.xlsx$/);
  const path = testInfo.outputPath("workbook.xlsx");
  await download.saveAs(path);
  const files = readXlsxEntries(await readFile(path));
  expect(files.get("xl/workbook.xml")).toContain('name="销售汇总"');
  expect(files.get("xl/workbook.xml")).toContain('name="未打开的表"');
  expect(files.get("xl/workbook.xml")).not.toContain("Deleted");
  expect(files.get("xl/workbook.xml")).toContain('state="hidden"');
  const first = files.get("xl/worksheets/sheet1.xml")!;
  expect(first).toContain('<mergeCell ref="A2:B2"/>');
  expect(first).toContain('sqref="B3:B4"');
  expect(first).toContain('errorStyle="warning"');
  expect(first).toMatch(/<c r="B4" s="\d+"><v>3<\/v>/);
  expect(files.get("xl/styles.xml")).toContain("<b/>");
  expect(files.get("xl/styles.xml")).toContain('rgb="FFFF8000"');
  const second = files.get("xl/worksheets/sheet2.xml")!;
  expect(second).toContain("second-last");
  expect(second).toContain('sqref="B2:B3"');
  expect(second).toContain('type="decimal" operator="greaterThan"');
  expect(files.get("xl/worksheets/sheet3.xml")).toContain("a,b");
  await expect(page.locator(".cyx-workbook-sheet")).toHaveCount(1);
  await expect(page.getByRole("tab", { name: "销售汇总", exact: true })).toHaveAttribute("aria-selected", "true");
});

test("Blob API exports the same workbook without custom headers and without mounting inactive sheets", async ({ page }) => {
  const bytes = await page.evaluate(async () => {
    const blob = await window.exportWorkbookForTest.exportExcel({ download: false, includeHeaders: false });
    return Array.from(new Uint8Array(await (blob as Blob).arrayBuffer()));
  });
  const files = readXlsxEntries(new Uint8Array(bytes));
  const first = files.get("xl/worksheets/sheet1.xml")!;
  expect(first).toContain('<mergeCell ref="A1:B1"/>');
  expect(first).toContain('sqref="B2:B3"');
  expect(first).toMatch(/<c r="B3" s="\d+"><v>3<\/v>/);
  await expect(page.locator(".cyx-workbook-sheet")).toHaveCount(1);
  await page.getByRole("tab", { name: "未打开的表", exact: true }).click();
  expect(await page.evaluate(() => window.exportWorkbookForTest.getActiveSheet().getCell(1, 0))).toBe("second-last");
});
