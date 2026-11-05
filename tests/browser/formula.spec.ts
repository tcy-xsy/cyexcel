import { expect, test } from "@playwright/test";
import type { WorkbookHandle } from "../../src/createWorkbook";
import { readXlsxEntries } from "../support/xlsx";

declare global { interface Window { formulaWorkbook: WorkbookHandle } }
test.beforeEach(async ({ page }) => {
  await page.goto('/tests/browser/formula.html');
  await page.waitForFunction(() => Boolean(window.formulaWorkbook));
});

test('formula bar edits, transitive recalculation, undo and redo retain formula source and styles', async ({ page }) => {
  await page.evaluate(() => {
    const sheet = window.formulaWorkbook.getActiveSheet();
    sheet.setCellStyle(0, 1, { bold: true });
    sheet.selectCell(0, 1);
  });
  const bar = page.getByRole('textbox', { name: '单元格内容', exact: true });
  await expect(bar).toHaveValue('=SUM(A1:A2)');
  await bar.fill('=A1*4');
  await bar.press('Enter');
  expect(await page.evaluate(() => {
    const sheet = window.formulaWorkbook.getActiveSheet();
    const values = [sheet.getCell(0, 1), sheet.getCell(0, 2)];
    sheet.setCell(0, 0, 7); values.push(sheet.getCell(0, 2));
    sheet.undo(); values.push(sheet.getCell(0, 2));
    sheet.redo(); values.push(sheet.getCell(0, 2));
    return { values, formula: sheet.getCellFormula(0, 1), styles: sheet.toSnapshot().sheets[sheet.toSnapshot().sheetOrder[0]!]!.cellData[0]?.[1]?.s };
  })).toEqual({ values: [8, 16, 56, 16, 56], formula: '=A1*4', styles: expect.any(String) });
});

test('numeric cell editor accepts a formula and validation checks its calculated value', async ({ page }) => {
  const canvas = await page.locator('#app canvas').first().boundingBox();
  if (!canvas) throw new Error('Missing canvas');
  await page.mouse.dblclick(canvas.x + 75, canvas.y + 15);
  const editor = page.locator('#app input[style*="position: absolute"]');
  await expect(editor).toHaveCount(1);
  await editor.fill('=2+5');
  await editor.press('Enter');
  await expect(page.getByRole('alertdialog')).toHaveCount(0);
  expect(await page.evaluate(() => {
    const sheet = window.formulaWorkbook.getActiveSheet();
    return [sheet.getCell(0, 0), sheet.getCellFormula(0, 0), sheet.getCell(0, 2)];
  })).toEqual([7, '=2+5', 20]);
  await page.mouse.dblclick(canvas.x + 75, canvas.y + 15);
  await expect(editor).toHaveValue('=2+5');
  await editor.fill('=200');
  await editor.press('Enter');
  await expect(page.getByRole('alertdialog')).toHaveCount(1);
});

test('exports fresh formula caches in active and unopened sheets without mounting them', async ({ page }) => {
  const bytes = await page.evaluate(async () => {
    const blob = await window.formulaWorkbook.exportExcel({ download: false });
    return [...new Uint8Array(await (blob as Blob).arrayBuffer())];
  });
  const entries = readXlsxEntries(new Uint8Array(bytes));
  expect(entries.get('xl/worksheets/sheet1.xml')).toContain('<f>SUM(A2:A3)</f><v>5</v>');
  expect(entries.get('xl/worksheets/sheet2.xml')).toContain('<f>A2*3</f><v>21</v>');
  await expect(page.locator('.cyx-workbook-sheet')).toHaveCount(1);
});

test('pastes formula text as one undoable calculation transaction', async ({ page, context }) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  const result = await page.evaluate(async () => {
    const sheet = window.formulaWorkbook.getActiveSheet();
    await navigator.clipboard.writeText('=A1+1\t=B2*2');
    await sheet.paste({ range: { startRow: 1, endRow: 1, startColumn: 1, endColumn: 2 } });
    const values = [sheet.getCell(1, 1), sheet.getCell(1, 2)];
    const formulas = [sheet.getCellFormula(1, 1), sheet.getCellFormula(1, 2)];
    sheet.undo();
    return { values, formulas, restored: [sheet.getCell(1, 1), sheet.getCell(1, 2)] };
  });
  expect(result).toEqual({ values: [3, 6], formulas: ['=A1+1', '=B2*2'], restored: [0, 0] });
});
