import { describe, expect, it, vi } from "vitest";
import { exportSheetToExcel, exportWorkbookToExcel } from "../../../../src/core/toolbar/exportExcel";
import type { ExcelSheetExportSource } from "../../../../src/core/toolbar/exportExcel";
import type { CyExcelCore } from "../../../../src/core/CyExcelCore";
import { matrixToWorkbookSnapshot } from "../../../../src/core/sheet/import";
import { WorkbookModel } from "../../../../src/core/sheet/WorkbookModel";
import type { DataValidationRule } from "../../../../src/core/sheet/validations";
import { readXlsxEntries } from "../../../support/xlsx";

const range = { startRow: 0, startColumn: 0, endRow: 1, endColumn: 0 };
async function files(sources: readonly ExcelSheetExportSource[], includeHeaders = true) {
  const blob = await exportWorkbookToExcel(sources, { download: false, includeHeaders });
  expect(blob).toBeInstanceOf(Blob);
  return readXlsxEntries(new Uint8Array(await (blob as Blob).arrayBuffer()));
}
function source(name = "Sheet1"): ExcelSheetExportSource {
  return { name, snapshot: matrixToWorkbookSnapshot([["first", 1], ["last", false]]) };
}

describe("Excel export", () => {
  it("exports malformed formulas as errors without emitting broken formula XML", async () => {
    const model = WorkbookModel.fromSnapshot(source().snapshot);
    model.sheet().setRecord(0, 0, { f: '=SUM(', v: '#ERROR!', fe: '#ERROR!' });
    const onWarning = vi.fn();
    const blob = await exportWorkbookToExcel([{ snapshot: model.toSnapshot() }], { download: false, onWarning });
    const xml = readXlsxEntries(new Uint8Array(await (blob as Blob).arrayBuffer())).get('xl/worksheets/sheet1.xml')!;
    expect(xml).toContain('<c r="A1" t="e"><v>#VALUE!</v></c>');
    expect(xml).not.toContain('<f>');
    expect(onWarning).toHaveBeenCalledWith(expect.stringContaining('invalid formulas'));
  });
  it("exports formula text, typed cached results and shifts references with custom headers", async () => {
    const model = WorkbookModel.fromSnapshot(source().snapshot);
    model.sheet().setRecord(0, 0, { f: '=B1*2', v: 2 });
    model.sheet().setRecord(1, 0, { f: '=CONCATENATE("A1","x")', v: "A1x" });
    model.sheet().setRecord(0, 1, { f: '=TRUE', v: true });
    model.sheet().setRecord(1, 1, { f: '=1/0', v: '#DIV/0!', fe: '#DIV/0!' });
    const result = await files([{ snapshot: model.toSnapshot(), columns: [{ title: "Formula" }] }]);
    const xml = result.get("xl/worksheets/sheet1.xml")!;
    expect(xml).toContain('<f>B2*2</f><v>2</v>');
    expect(xml).toContain('t="str"><f>CONCATENATE(&quot;A1&quot;,&quot;x&quot;)</f><v>A1x</v>');
    expect(xml).toContain('t="b"><f>TRUE</f><v>1</v>');
    expect(xml).toContain('t="e"><f>1/0</f><v>#DIV/0!</v>');
    expect(result.get("xl/workbook.xml")).toContain('<calcPr fullCalcOnLoad="1"/>');
  });
  it("keeps every data row with custom headers and preserves scalar types and escaping", async () => {
    const input = { ...source(), columns: [{ title: "Name & <label>" }, { title: "Value" }] };
    const xml = (await files([input])).get("xl/worksheets/sheet1.xml")!;
    expect(xml).toContain('r="A1"');
    expect(xml).toContain("Name &amp; &lt;label&gt;");
    expect(xml).toContain('<c r="A2" t="inlineStr"><is><t xml:space="preserve">first</t>');
    expect(xml).toContain('<c r="A3" t="inlineStr"><is><t xml:space="preserve">last</t>');
    expect(xml).toContain('<c r="B2"><v>1</v></c>');
    expect(xml).toContain('<c r="B3" t="b"><v>0</v></c>');
    const withoutHeaders = (await files([input], false)).get("xl/worksheets/sheet1.xml")!;
    expect(withoutHeaders).toContain('<c r="A2" t="inlineStr"><is><t xml:space="preserve">last</t>');
    expect(withoutHeaders).not.toContain("label");
  });

  it("writes ordered sheets, names, relationships and a shared style table", async () => {
    const first = source("Sales & totals");
    const second = source("Inventory");
    const result = await files([first, second]);
    expect(result.get("xl/workbook.xml")).toContain('<sheet name="Sales &amp; totals" sheetId="1" r:id="rId1"/>');
    expect(result.get("xl/workbook.xml")).toContain('<sheet name="Inventory" sheetId="2" r:id="rId2"/>');
    expect(result.get("xl/_rels/workbook.xml.rels")).toContain('Target="worksheets/sheet2.xml"');
    expect(result.get("[Content_Types].xml")).toContain('PartName="/xl/worksheets/sheet2.xml"');
    expect(result.get("xl/worksheets/sheet2.xml")).toContain("last");
    expect(result.has("xl/styles.xml")).toBe(true);
  });

  it("normalizes illegal, long and duplicate sheet names without losing sheets", async () => {
    const onWarning = vi.fn();
    const blob = await exportWorkbookToExcel([source("a/b"), source("A_B"), source("x".repeat(40)), source("''")], { download: false, onWarning });
    const xml = readXlsxEntries(new Uint8Array(await (blob as Blob).arrayBuffer())).get("xl/workbook.xml")!;
    expect(xml).toContain('name="a_b"');
    expect(xml).toContain('name="A_B (1)"');
    expect(xml).toContain(`name="${"x".repeat(31)}"`);
    expect(xml).toContain('name="Sheet"');
    expect(onWarning).toHaveBeenCalledTimes(4);
  });

  it("shifts merges and every validation range when headers are inserted", async () => {
    const model = WorkbookModel.fromSnapshot(source().snapshot);
    model.sheet().setMerges([{ ...range, endColumn: 1 }]);
    model.sheet().setValidations([{ id: "whole", ranges: [range, { ...range, startColumn: 1, endColumn: 1 }],
      criteria: { type: "whole", operator: "between", value1: 1, value2: 10 }, allowBlank: false,
      errorStyle: "warning", errorTitle: "Check & retry", errorMessage: "1 < value < 10" }]);
    const input = { snapshot: model.toSnapshot(), columns: [{ title: "Name" }, { title: "Amount" }] };
    const xml = (await files([input])).get("xl/worksheets/sheet1.xml")!;
    expect(xml).toContain('<mergeCell ref="A2:B3"/>');
    expect(xml).toContain('sqref="A2:A3 B2:B3"');
    expect(xml).toContain('allowBlank="0"');
    expect(xml).toContain('errorStyle="warning"');
    expect(xml).toContain('errorTitle="Check &amp; retry"');
    expect(xml).toContain('<formula1>1</formula1><formula2>10</formula2>');
    const unshifted = (await files([input], false)).get("xl/worksheets/sheet1.xml")!;
    expect(unshifted).toContain('<mergeCell ref="A1:B2"/>');
    expect(unshifted).toContain('sqref="A1:A2 B1:B2"');
  });

  it("uses hidden named ranges for long lists and special characters across multiple sheets", async () => {
    const items = ['a,b', 'say "yes"', "line\nbreak", "z".repeat(300)];
    const model = WorkbookModel.fromSnapshot(source().snapshot);
    model.sheet().setValidations([{ id: "list", ranges: [range], criteria: { type: "list", items }, showDropDown: false }]);
    const result = await files([{ name: "_cyexcel_lists", snapshot: model.toSnapshot() }, { name: "Other", snapshot: model.toSnapshot() }]);
    expect(result.get("xl/workbook.xml")).toContain('name="_cyexcel_lists (1)" sheetId="3" r:id="rId3" state="hidden"');
    expect(result.get("xl/workbook.xml")).toContain("$A$1:$A$4");
    expect(result.get("xl/workbook.xml")).not.toContain("$A$5");
    expect(result.get("xl/worksheets/sheet1.xml")).toContain('<formula1>_cyexcel_list_1</formula1>');
    expect(result.get("xl/worksheets/sheet2.xml")).toContain('<formula1>_cyexcel_list_1</formula1>');
    expect(result.get("xl/worksheets/sheet1.xml")).toContain('showDropDown="1"');
    const helper = result.get("xl/worksheets/sheet3.xml")!;
    expect(helper).toContain("a,b");
    expect(helper).toContain("say &quot;yes&quot;");
    expect(helper).toContain("line\nbreak");
    expect(helper).toContain("z".repeat(300));
  });

  it.each(["decimal", "date", "time", "textLength"] as const)("exports %s built-in validation", async type => {
    const model = WorkbookModel.fromSnapshot(source().snapshot);
    model.sheet().setValidations([{ id: type, ranges: [range], criteria: { type, operator: "greaterThanOrEqual", value1: 0.5 } }]);
    const xml = (await files([{ snapshot: model.toSnapshot() }])).get("xl/worksheets/sheet1.xml")!;
    expect(xml).toContain(`type="${type}" operator="greaterThanOrEqual"`);
    expect(xml).toContain('<formula1>0.5</formula1>');
  });

  it("reports custom JavaScript validators instead of creating an invalid Excel formula", async () => {
    const model = WorkbookModel.fromSnapshot(source().snapshot);
    const rules: DataValidationRule[] = [{ id: "custom", ranges: [range], criteria: { type: "custom", name: "businessRule" } }];
    model.sheet().setValidations(rules);
    const onWarning = vi.fn();
    const blob = await exportWorkbookToExcel([{ name: "Business", snapshot: model.toSnapshot() }], { download: false, onWarning });
    expect(onWarning).toHaveBeenCalledWith(expect.stringContaining('custom validator "businessRule"'));
    const result = readXlsxEntries(new Uint8Array(await (blob as Blob).arrayBuffer()));
    expect(result.get("xl/worksheets/sheet1.xml")).not.toContain("<dataValidations");
  });

  it("preserves the final row's styles, column defaults, row heights and hidden metadata", async () => {
    const base = source().snapshot;
    const id = base.sheetOrder[0]!;
    const sheet = base.sheets[id]!;
    const snapshot = { ...base, styles: { column: { bold: true, numFmt: "0.00" }, last: { color: "rgb(255, 128, 0)" } },
      sheets: { [id]: { ...sheet, columnData: { 1: { s: "column", w: 145, hd: true } }, rowData: { 1: { h: 40, hd: true } },
        cellData: { ...sheet.cellData, 1: { ...sheet.cellData[1], 1: { v: 9, s: "last" } } } } } };
    const result = await files([{ snapshot, columns: [{ title: "Name", wrap: true }, { title: "Number" }] }]);
    expect(result.get("xl/worksheets/sheet1.xml")).toContain('<row r="3" ht="30" customHeight="1" hidden="1">');
    expect(result.get("xl/worksheets/sheet1.xml")).toMatch(/<c r="B3" s="\d+"><v>9<\/v>/);
    expect(result.get("xl/worksheets/sheet1.xml")).toContain('width="20.00" customWidth="1" style=');
    expect(result.get("xl/styles.xml")).toContain('rgb="FFFF8000"');
    expect(result.get("xl/styles.xml")).toContain('formatCode="0.00"');
    expect(result.get("xl/styles.xml")).toContain('wrapText="1"');
  });

  it("exports sparse sheets without traversing blank coordinate-space rows", async () => {
    const model = WorkbookModel.fromSnapshot(source().snapshot);
    model.sheet().setCell(999999, 0, "far away");
    const result = await files([{ snapshot: model.toSnapshot() }]);
    const xml = result.get("xl/worksheets/sheet1.xml")!;
    expect(xml.match(/<row /g)).toHaveLength(3);
    expect(xml).toContain('r="A1000000"');
    expect(xml).toContain("far away");
  });

  it("rejects Excel row/column overflows, including header offsets", async () => {
    const model = WorkbookModel.fromSnapshot(source().snapshot);
    model.sheet().setCell(1048575, 0, "last Excel row");
    await expect(files([{ snapshot: model.toSnapshot(), columns: [{ title: "Header" }] }])).rejects.toThrow("row limit");
    model.sheet().setCell(0, 16384, "too wide");
    await expect(files([{ snapshot: model.toSnapshot() }])).rejects.toThrow("column limit");
    await expect(files([])).rejects.toThrow("without sheets");
  });

  it("keeps the single-sheet API and defaults to the current sheet style resolver", async () => {
    const snapshot = source().snapshot;
    const fake = { toSnapshot: () => snapshot, getColumnCount: () => 2,
      getHeader: (column: number) => ({ title: column === 0 ? "Name" : "Number" }),
      getColumn: () => ({ width: 120, wrap: false }) } as unknown as CyExcelCore;
    const blob = await exportSheetToExcel(fake, { download: false }, () => ({ italic: true }));
    const result = readXlsxEntries(new Uint8Array(await (blob as Blob).arrayBuffer()));
    expect(result.get("xl/styles.xml")).toContain("<i/>");
    expect(result.get("xl/worksheets/sheet1.xml")).toContain('r="A3"');
  });
});
