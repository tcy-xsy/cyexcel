import type { Worksheet } from "exceljs";

const MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function headings(sheet: Worksheet, values: string[]): void {
  sheet.addRow(values);
  sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" }, size: 12 };
  sheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF264653" } };
  sheet.getRow(1).height = 25;
  sheet.views = [{ state: "frozen", ySplit: 1 }];
  sheet.getColumn(1).width = 25;
  sheet.getColumn(2).width = 32;
  sheet.getColumn(3).width = 55;
}

/** A real compressed XLSX fixture covering the importer's supported representations. */
export async function createExcelImportSample(): Promise<Blob> {
  const ExcelJS = await import("exceljs");
  const workbook = new (ExcelJS.default ?? ExcelJS).Workbook();
  const values = workbook.addWorksheet("Values");
  headings(values, ["Type", "Value", "Expected import"]);
  values.addRows([
    ["Text", "Hello Excel", "Plain text"],
    ["Leading zeros", "00123", "Text retains leading zeros"],
    ["Integer", 42, "Number"],
    ["Decimal", -1234.567, "Number"],
    ["Zero", 0, "Zero is not an empty cell"],
    ["Boolean true", true, "Boolean"],
    ["Boolean false", false, "False is not an empty cell"],
    ["Multiline", "First line\nSecond line", "Wrapping is applied to the whole column"],
    ["Date", new Date("2026-11-06T00:00:00Z"), "Excel date serial with a date number format"],
    ["Styled empty cell", null, "An empty cell can retain its fill and border"],
    ["Hyperlink", { text: "Example website", hyperlink: "https://example.com" }, "Display text is retained; the link is omitted"],
    ["Rich text", { richText: [{ text: "Plain " }, { text: "bold", font: { bold: true } }] }, "Text is retained; formatting of individual runs is omitted"],
    ["Error", { error: "#N/A" }, "Error text is retained"]
  ]);
  values.getCell("B3").numFmt = "@";
  values.getCell("B9").alignment = { wrapText: true };
  values.getRow(9).height = 42;
  values.getCell("B10").numFmt = "yyyy-mm-dd";
  values.getCell("B11").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFE4A3" } };
  values.getCell("B11").border = { bottom: { style: "thin", color: { argb: "FFD97706" } } };

  const formats = workbook.addWorksheet("Number formats");
  headings(formats, ["Format", "Value", "Excel number format"]);
  const numbers = [
    ["Integer", 1234, "0"],
    ["Decimal", 1234.567, "0.00"],
    ["Thousands", 1234567.89, "#,##0.00"],
    ["Currency", 1234.5, '"$"#,##0.00'],
    ["Percentage", 0.256, "0.0%"],
    ["Scientific", 1234567, "0.00E+00"],
    ["Fraction", 0.75, "# ?/?"],
    ["Date", 46332, "yyyy-mm-dd"],
    ["Time", 0.5625, "hh:mm:ss"],
    ["Date and time", 46332.5625, "yyyy-mm-dd hh:mm"],
    ["Duration", 1.5, "[h]:mm"],
    ["Zero padding", 123, "000000"]
  ] as const;
  for (const [name, value, format] of numbers) {
    const row = formats.addRow([name, value, format]);
    row.getCell(2).numFmt = format;
  }

  const formulas = workbook.addWorksheet("Formulas");
  headings(formulas, ["Function", "Expected result", "Calculated result"]);
  formulas.getColumn(5).width = 15;
  formulas.getCell("E1").value = "Input";
  const calculations = [
    ["SUM", "SUM(E2:E4)", 60],
    ["AVERAGE", "AVERAGE(E2:E4)", 20],
    ["MIN", "MIN(E2:E4)", 10],
    ["MAX", "MAX(E2:E4)", 30],
    ["COUNT", "COUNT(E2:E4)", 3],
    ["COUNTIF", 'COUNTIF(E2:E4,">15")', 2],
    ["SUMIF", 'SUMIF(E2:E4,">15",E2:E4)', 50],
    ["IF", 'IF(E2>5,"Yes","No")', "Yes"],
    ["IFERROR", "IFERROR(1/0,0)", 0],
    ["ROUND", "ROUND(1.2345,2)", 1.23],
    ["ABS", "ABS(-10)", 10],
    ["CONCATENATE", 'CONCATENATE("Hello"," Excel")', "Hello Excel"],
    ["Arithmetic", "E2+E3*2", 50],
    ["Absolute reference", "$E$2*2", 20],
    ["Cached error", "1/0", { error: "#DIV/0!" as const }]
  ] as const;
  for (const [name, formula, result] of calculations) formulas.addRow([name, result, { formula, result }]);
  formulas.addRow(["Shared formula 1", 20]);
  formulas.addRow(["Shared formula 2", 40]);
  formulas.fillFormula("C17:C18", "E2*2", [20, 40]);
  formulas.getCell("E2").value = 10;
  formulas.getCell("E3").value = 20;
  formulas.getCell("E4").value = 30;

  const styles = workbook.addWorksheet("Styles");
  headings(styles, ["Style", "Preview", "Expected import"]);
  styles.addRows([
    ["Bold", "Bold text", "Bold font"],
    ["Italic", "Italic text", "Italic font"],
    ["Underline", "Underlined text", "Underline"],
    ["Strikethrough", "Old value", "Strikethrough"],
    ["Font family and size", "Georgia 18 pt", "24 px font size"],
    ["Text color", "Orange text", "#d97706"],
    ["Fill color", "Green fill", "#d1fae5"],
    ["Theme color", "Accent color", "Workbook theme color"],
    ["Left align", "Left", "Horizontal alignment"],
    ["Center align", "Center", "Horizontal alignment"],
    ["Right align", "Right", "Horizontal alignment"],
    ["Top align", "Top", "Vertical alignment"],
    ["Middle align", "Middle", "Vertical alignment"],
    ["Bottom align", "Bottom", "Vertical alignment"],
    ["Thin border", "All four sides", "1 px solid border"],
    ["Medium border", "All four sides", "2 px solid border"],
    ["Thick border", "All four sides", "3 px solid border"],
    ["Column style", 1234.5, "Column D uses a default number format"]
  ]);
  styles.getCell("B2").font = { bold: true };
  styles.getCell("B3").font = { italic: true };
  styles.getCell("B4").font = { underline: true };
  styles.getCell("B5").font = { strike: true };
  styles.getCell("B6").font = { name: "Georgia", size: 18 };
  styles.getCell("B7").font = { color: { argb: "FFD97706" } };
  styles.getCell("B8").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD1FAE5" } };
  styles.getCell("B9").font = { color: { theme: 4 } };
  for (const [index, horizontal] of ["left", "center", "right"].entries()) styles.getCell(`B${index + 10}`).alignment = { horizontal: horizontal as "left" | "center" | "right" };
  for (const [index, vertical] of ["top", "middle", "bottom"].entries()) {
    styles.getCell(`B${index + 13}`).alignment = { vertical: vertical as "top" | "middle" | "bottom" };
    styles.getRow(index + 13).height = 40;
  }
  for (const [index, style] of (["thin", "medium", "thick"] as const).entries()) {
    const side = { style, color: { argb: "FF264653" } };
    styles.getCell(`B${index + 16}`).border = { top: side, right: side, bottom: side, left: side };
  }
  styles.getColumn(4).numFmt = "#,##0.00";
  styles.getColumn(4).width = 20;
  styles.getCell("D19").value = 1234.5;

  const layout = workbook.addWorksheet("Layout");
  headings(layout, ["Feature", "Preview", "Expected import"]);
  layout.views = [{ state: "frozen", xSplit: 1, ySplit: 1 }];
  layout.addRow(["Merged cells", "Merged B2:C3"]);
  layout.mergeCells("B2:C3");
  layout.getCell("B2").alignment = { horizontal: "center", vertical: "middle" };
  layout.getCell("B2").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDBEAFE" } };
  layout.getCell("A4").value = "Custom row height";
  layout.getCell("B4").value = "48 pt / 64 px";
  layout.getRow(4).height = 48;
  layout.getCell("A5").value = "Hidden row";
  layout.getCell("B5").value = "This row should not be visible";
  layout.getRow(5).hidden = true;
  layout.getCell("A6").value = "Hidden column";
  layout.getCell("B6").value = "Column D is hidden";
  layout.getCell("D6").value = "Hidden content is retained";
  layout.getColumn(4).hidden = true;
  layout.getCell("A7").value = "Frozen panes";
  layout.getCell("B7").value = "First row and first column";
  layout.getCell("A8").value = "Column width";
  layout.getCell("B8").value = "Column B uses width 32";
  layout.getCell("A9").value = "Multiline / wrapping";
  layout.getCell("B9").value = "Line one\nLine two\nLine three";
  layout.getCell("B9").alignment = { wrapText: true };
  layout.getRow(9).height = 60;
  layout.getCell("A10").value = "Sparse cells";
  layout.getCell("H20").value = "Retained beyond empty rows and columns";
  const hidden = workbook.addWorksheet("Hidden sheet", { state: "hidden" });
  hidden.getCell("A1").value = "Hidden worksheets are imported as visible tabs.";
  hidden.getColumn(1).width = 64;

  return new Blob([new Uint8Array(await workbook.xlsx.writeBuffer())], { type: MIME });
}
