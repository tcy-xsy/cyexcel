import ExcelJS from "exceljs";
import { describe, expect, it, vi } from "vitest";
import type { WorkbookHandle } from "cyexcel";
import { WorkbookModel } from "cyexcel/core";
import { createFormulaEngine } from "cyexcel/formula";
import { createExcelImportPlugin, readExcelWorkbook } from "../src/index";
import { createExcelImportSample } from "../examples/sample";

async function fixture(): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Sales");
  sheet.getCell("A1").value = "00123";
  sheet.getCell("B1").value = false;
  sheet.getCell("C1").value = 4;
  sheet.getCell("C2").value = { formula: "C1*2", result: 8 };
  sheet.fillFormula("C3:C4", "C1*2", [8, 16]);
  sheet.getCell("C5").value = { formula: "1/0", result: { error: "#DIV/0!" } };
  sheet.getCell("D2").value = new Date("2026-01-15T00:00:00Z");
  sheet.getCell("D2").numFmt = "yyyy-mm-dd";
  sheet.getCell("E2").value = { richText: [{ text: "hello " }, { text: "world", font: { bold: true } }] };
  sheet.getCell("F2").value = { text: "Link", hyperlink: "https://example.com" };
  sheet.getCell("A1").font = { bold: true, size: 15, color: { argb: "FFFF8000" } };
  sheet.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
  sheet.getCell("A1").alignment = { horizontal: "center", wrapText: true };
  sheet.getCell("A1").border = { bottom: { style: "thin", color: { argb: "FF112233" } } };
  sheet.getCell("J12").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF336699" } };
  sheet.getRow(3).height = 30;
  sheet.getRow(8).height = 24;
  sheet.getRow(8).hidden = true;
  sheet.getColumn(2).width = 20;
  sheet.getColumn(7).hidden = true;
  sheet.mergeCells("A6:B6");
  sheet.getCell("A6").value = "merged";
  sheet.views = [{ state: "frozen", xSplit: 1, ySplit: 1 }];
  sheet.getCell("B3").dataValidation = { type: "list", formulae: ['"A,B"'] };
  workbook.addWorksheet("Hidden", { state: "hidden" }).getCell("A1").value = "retained";
  return new Uint8Array(await workbook.xlsx.writeBuffer());
}

describe("optional Excel import plugin", () => {
  it("imports the documentation sample with all six feature worksheets", async () => {
    const result = await readExcelWorkbook(await createExcelImportSample());
    expect(result.sheets.map(sheet => sheet.name)).toEqual(["Values", "Number formats", "Formulas", "Styles", "Layout", "Hidden sheet"]);
    const snapshots = result.sheets.map(sheet => sheet.snapshot!);
    const [values, formats, formulas, styles, layout] = snapshots.map(snapshot => snapshot.sheets[snapshot.sheetOrder[0]!]!);
    expect(values!.cellData[2]![1]!.v).toBe("00123");
    expect(values!.cellData[7]![1]!.v).toBe(false);
    expect(values!.cellData[9]![1]!.v).toBe(46332);
    expect(snapshots[0]!.styles[values!.cellData[10]![1]!.s!]).toMatchObject({ background: "#ffe4a3" });
    expect(snapshots[1]!.styles[formats!.cellData[5]![1]!.s!]).toMatchObject({ numFmt: "0.0%" });
    expect(formulas!.cellData[1]![2]).toMatchObject({ f: "=SUM(E2:E4)", v: 60 });
    expect(formulas!.cellData[15]![2]).toMatchObject({ f: "=1/0", fe: "#DIV/0!" });
    expect(formulas!.cellData[16]![2]).toMatchObject({ f: "=E2*2", v: 20 });
    expect(formulas!.cellData[17]![2]).toMatchObject({ f: "=E3*2", v: 40 });
    const formulaSheet = WorkbookModel.fromSnapshot(snapshots[2]!).sheet();
    formulaSheet.enableFormulas(createFormulaEngine());
    for (let row = 1; row <= 17; row++) {
      expect(formulaSheet.getCell(row, 2)).toEqual(formulaSheet.getCell(row, 1));
      expect(formulaSheet.getRecord(row, 2)?.f).toMatch(/^=/);
      expect(formulaSheet.getRecord(row, 1)?.f).toBeUndefined();
    }
    formulaSheet.setCell(1, 4, 100);
    expect(formulaSheet.getCell(1, 2)).toBe(150);
    expect(formulaSheet.getCell(1, 1)).toBe(60);
    expect(snapshots[3]!.styles[styles!.cellData[5]![1]!.s!]).toMatchObject({ fontFamily: "Georgia", fontSize: 24 });
    expect(layout!.freeze).toEqual({ rows: 1, columns: 1 });
    expect(layout!.mergeData).toEqual([{ startRow: 1, startColumn: 1, endRow: 2, endColumn: 2 }]);
    expect(layout!.rowData[4]!.hd).toBe(true);
    expect(layout!.columnData[3]!.hd).toBe(true);
    expect(layout!.cellData[19]![7]!.v).toBe("Retained beyond empty rows and columns");
    for (const snapshot of snapshots) expect(() => WorkbookModel.fromSnapshot(snapshot)).not.toThrow();
  });

  it("reads compressed XLSX, preserves types, formulas, styles, sparse blanks and tab order", async () => {
    const onWarning = vi.fn();
    const result = await readExcelWorkbook(await fixture(), { onWarning });
    expect(result.sheets.map(sheet => sheet.name)).toEqual(["Sales", "Hidden"]);
    const snapshot = result.sheets[0]!.snapshot!;
    const sheet = snapshot.sheets[snapshot.sheetOrder[0]!]!;
    expect(sheet.cellData[0]![0]!.v).toBe("00123");
    expect(sheet.cellData[0]![1]!.v).toBe(false);
    expect(sheet.cellData[1]![2]).toMatchObject({ v: 8, f: "=C1*2" });
    expect(sheet.cellData[3]![2]).toMatchObject({ v: 16, f: "=C2*2" });
    expect(sheet.cellData[4]![2]).toMatchObject({ v: "#DIV/0!", fe: "#DIV/0!" });
    expect(sheet.cellData[1]![3]!.v).toBe(46037);
    expect(sheet.cellData[1]![4]!.v).toBe("hello world");
    expect(sheet.cellData[1]![5]!.v).toBe("Link");
    expect(snapshot.styles[sheet.cellData[0]![0]!.s!]).toMatchObject({ bold: true, fontSize: 20, color: "#ff8000", background: "#ffff00", align: "center", border: { bottom: { color: "#112233" } } });
    expect(snapshot.styles[sheet.cellData[11]![9]!.s!]).toMatchObject({ background: "#336699" });
    expect(sheet.rowData[2]).toEqual({ h: 40, hp: "manual" });
    expect(sheet.rowData[7]).toEqual({ h: 32, hp: "manual", hd: true });
    expect(sheet.columnData[1]!.w).toBe(145);
    expect(sheet.columnData[6]!.hd).toBe(true);
    expect(sheet.mergeData).toEqual([{ startRow: 5, startColumn: 0, endRow: 5, endColumn: 1 }]);
    expect(sheet.cellData[5]?.[1]?.v).toBeUndefined();
    expect(sheet.freeze).toEqual({ rows: 1, columns: 1 });
    expect(result.sheets[0]!.columns![0]!.wrap).toBe(true);
    expect(result.warnings).toContain("Data validation rules are not imported.");
    expect(onWarning).toHaveBeenCalled();
    for (const imported of result.sheets) expect(() => WorkbookModel.fromSnapshot(imported.snapshot!)).not.toThrow();
  });

  it("normalizes 1904 date workbooks to the model's 1900 date system", async () => {
    const workbook = new ExcelJS.Workbook();
    workbook.properties.date1904 = true;
    const cell = workbook.addWorksheet("Dates").getCell("A1");
    cell.value = new Date("2026-01-15T00:00:00Z");
    cell.numFmt = "yyyy-mm-dd";
    const result = await readExcelWorkbook(new Uint8Array(await workbook.xlsx.writeBuffer()));
    const snapshot = result.sheets[0]!.snapshot!;
    expect(snapshot.sheets[snapshot.sheetOrder[0]!]!.cellData[0]![0]!.v).toBe(46037);
  });

  it("rejects invalid, oversized, over-limit, empty and cancelled files without applying data", async () => {
    const plugin = createExcelImportPlugin();
    const replaceSheets = vi.fn();
    plugin.attach({ replaceSheets } as unknown as WorkbookHandle);
    await expect(plugin.importExcel(new Uint8Array([1, 2, 3]))).rejects.toThrow(".xlsx");
    await expect(plugin.importExcel(new Uint8Array([80, 75, 1, 2]))).rejects.toThrow("Unable to read");
    const bytes = await fixture();
    await expect(plugin.importExcel(bytes, { maxFileSize: 1 })).rejects.toThrow("byte limit");
    await expect(plugin.importExcel(bytes, { maxCells: 1 })).rejects.toThrow("cell import limit");
    const empty = new ExcelJS.Workbook();
    await expect(plugin.importExcel(new Uint8Array(await empty.xlsx.writeBuffer()))).rejects.toThrow("no worksheets");
    const controller = new AbortController();
    controller.abort();
    await expect(plugin.importExcel(bytes, { signal: controller.signal })).rejects.toThrow();
    const later = new AbortController();
    const pending = plugin.importExcel(bytes, { signal: later.signal });
    later.abort();
    await expect(pending).rejects.toThrow();
    expect(replaceSheets).not.toHaveBeenCalled();
  });

  it("applies a complete result once and prevents overlapping imports and disposal races", async () => {
    const bytes = await fixture();
    const plugin = createExcelImportPlugin();
    const replaceSheets = vi.fn();
    plugin.attach({ replaceSheets } as unknown as WorkbookHandle);
    const pending = plugin.importExcel(bytes);
    await expect(plugin.importExcel(bytes)).rejects.toThrow("already in progress");
    const result = await pending;
    expect(replaceSheets).toHaveBeenCalledExactlyOnceWith(result.sheets);
    const cancelled = plugin.importExcel(bytes);
    plugin.dispose();
    await expect(cancelled).rejects.toThrow("cancelled");
    expect(replaceSheets).toHaveBeenCalledTimes(1);
  });
});
