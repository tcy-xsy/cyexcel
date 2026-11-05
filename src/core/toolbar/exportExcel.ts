import type { CyExcelCore } from "../CyExcelCore";
import type { BorderSide, CellStyle } from "../style/types";
import type { SheetSnapshot, WorkbookSnapshot } from "../sheet/types";
import type { DataValidationRule } from "../sheet/validations";
import type { CellRange } from "../types";
import { shiftFormulaRows } from "../formula/references";

export interface ExcelExportOptions {
  readonly fileName?: string;
  readonly includeHeaders?: boolean;
  /** Return the XLSX Blob instead of triggering a browser download. */
  readonly download?: boolean;
  /** Called when a feature cannot be represented in Excel (e.g. a JavaScript validator). */
  readonly onWarning?: (message: string) => void;
}

export interface ExcelExportColumn {
  readonly title?: string;
  readonly width?: number;
  readonly wrap?: boolean;
}

/** A single sheet's snapshot plus its optional presentation configuration. */
export interface ExcelSheetExportSource {
  readonly name?: string;
  readonly snapshot: WorkbookSnapshot;
  readonly columns?: readonly ExcelExportColumn[];
  readonly styleOf?: (row: number, column: number) => CellStyle | undefined;
}

interface ExcelFont {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly strike?: boolean;
  readonly family?: string;
  readonly size?: number;
  readonly color?: string;
}

interface ExcelFill {
  readonly color?: string;
}

interface ExcelBorderSide {
  readonly color?: string;
}

interface ExcelBorder {
  readonly top?: ExcelBorderSide;
  readonly left?: ExcelBorderSide;
  readonly bottom?: ExcelBorderSide;
  readonly right?: ExcelBorderSide;
}

interface ExcelAlignment {
  readonly horizontal?: CellStyle["align"];
  readonly vertical?: "top" | "center" | "bottom";
  readonly wrapText?: boolean;
}

interface ExcelCellFormat {
  readonly font?: ExcelFont;
  readonly fill?: ExcelFill;
  readonly border?: ExcelBorder;
  readonly alignment?: ExcelAlignment;
  readonly numFmt?: string;
}

const XML_DECLARATION = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";

function xmlEscape(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function columnLetter(index: number): string {
  let value = index + 1;
  let label = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

function cellReference(row: number, column: number): string {
  return `${columnLetter(column)}${row + 1}`;
}

function isDefaultColumnHeader(title: string, column: number): boolean {
  return title === columnLetter(column);
}

function styleText(value: string): string {
  return value.replace(/\r?\n/g, " ");
}

function styleKey(value: unknown): string {
  return JSON.stringify(value, (_unused, item: unknown) =>
    item === undefined ? null : item
  );
}

function cssColorToArgb(value: string): string | undefined {
  const source = value.trim();
  if (!source) return undefined;

  const hex = source.replace(/^#/, "");
  if (/^[0-9a-f]{3}$/i.test(hex)) {
    return `FF${hex.split("").map((item) => item + item).join("").toUpperCase()}`;
  }
  if (/^[0-9a-f]{4}$/i.test(hex)) {
    const [red, green, blue, alpha] = hex.split("");
    return `${alpha}${alpha}${red}${red}${green}${green}${blue}${blue}`.toUpperCase();
  }
  if (/^[0-9a-f]{6}$/i.test(hex)) {
    return `FF${hex.toUpperCase()}`;
  }
  if (/^[0-9a-f]{8}$/i.test(hex)) {
    return `${hex.slice(6)}${hex.slice(0, 6)}`.toUpperCase();
  }

  const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([\d.]+))?\s*\)$/i.exec(source);
  if (rgb) {
    return [Math.round(Number(rgb[4] ?? "1") * 255), Number(rgb[1]), Number(rgb[2]), Number(rgb[3])]
      .map(part => Math.max(0, Math.min(255, part)).toString(16).padStart(2, "0")).join("").toUpperCase();
  }
  if (typeof document === "undefined") return undefined;
  const probe = document.createElement("span");
  probe.style.color = source;
  const match = /^rgba?\((\d+), (\d+), (\d+)(?:, ([\d.]+))?\)$/i.exec(probe.style.color);
  if (!match) return undefined;
  const [, red, green, blue, alpha] = match;
  const alphaHex = Math.round(Number(alpha ?? "1") * 255)
    .toString(16)
    .padStart(2, "0");
  return `${alphaHex}${[red, green, blue].map(part => Number(part).toString(16).padStart(2, "0")).join("")}`.toUpperCase();
}

function createExcelFont(style: CellStyle): ExcelFont | undefined {
  const font: { -readonly [K in keyof ExcelFont]?: ExcelFont[K] } = {};
  if (style.bold === true) font.bold = true;
  if (style.italic === true) font.italic = true;
  if (style.underline === true) font.underline = true;
  if (style.strike === true) font.strike = true;
  if (style.fontFamily) font.family = styleText(style.fontFamily);
  if (style.fontSize !== undefined && style.fontSize > 0) {
    font.size = Math.round(style.fontSize * 0.75 * 10) / 10;
  }
  if (style.color) {
    const color = cssColorToArgb(style.color);
    if (color) font.color = color;
  }
  return Object.keys(font).length > 0 ? font : undefined;
}

function createExcelFill(style: CellStyle): ExcelFill | undefined {
  if (!style.background) return undefined;
  const color = cssColorToArgb(style.background);
  return color ? { color } : undefined;
}

function createExcelBorder(style: CellStyle): ExcelBorder | undefined {
  const border = style.border;
  if (!border) return undefined;

  const result: { -readonly [K in keyof ExcelBorder]?: ExcelBorder[K] } = {};
  const addSide = (
    name: keyof ExcelBorder,
    side: BorderSide | undefined
  ): void => {
    if (!side) return;
    const color = cssColorToArgb(side.color);
    if (color) result[name] = { color };
  };
  addSide("top", border.top);
  addSide("left", border.left);
  addSide("bottom", border.bottom);
  addSide("right", border.right);
  return Object.keys(result).length > 0 ? result : undefined;
}

function createExcelAlignment(style: CellStyle, wrap: boolean): ExcelAlignment | undefined {
  if (!style.align && !style.verticalAlign && !wrap) return undefined;

  const alignment: { -readonly [K in keyof ExcelAlignment]?: ExcelAlignment[K] } = {};
  if (style.align) alignment.horizontal = style.align;
  if (style.verticalAlign === "middle") alignment.vertical = "center";
  else if (style.verticalAlign) alignment.vertical = style.verticalAlign;
  if (wrap) alignment.wrapText = true;
  return alignment;
}

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (let index = 0; index < bytes.length; index += 1) {
    const byte = bytes[index] ?? 0;
    crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ byte) & 0xff]!;
  }
  return (crc ^ 0xffffffff) >>> 0;
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let index = 0; index < 256; index += 1) {
    let value = index;
    for (let bit = 0; bit < 8; bit += 1) {
      value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    }
    table[index] = value >>> 0;
  }
  return table;
})();

function u16(bytes: number[], value: number): void {
  bytes.push(value & 0xff, (value >>> 8) & 0xff);
}

function u32(bytes: number[], value: number): void {
  bytes.push(
    value & 0xff,
    (value >>> 8) & 0xff,
    (value >>> 16) & 0xff,
    (value >>> 24) & 0xff
  );
}

function pushBytes(target: number[], source: Uint8Array): void {
  for (const byte of source) {
    target.push(byte);
  }
}

function zipFiles(files: ReadonlyArray<{ name: string; content: string }>): Blob {
  const encoder = new TextEncoder();
  const localParts: number[] = [];
  const centralParts: number[] = [];
  const entries: Array<{ name: Uint8Array; crc: number; size: number; offset: number }> = [];

  for (const file of files) {
    const name = encoder.encode(file.name);
    const content = encoder.encode(file.content);
    const crc = crc32(content);
    const offset = localParts.length;
    entries.push({ name, crc, size: content.length, offset });

    u32(localParts, 0x04034b50);
    u16(localParts, 20);
    u16(localParts, 0x0800);
    u16(localParts, 0);
    u16(localParts, 0);
    u16(localParts, 0x21); // 1980-01-01; deterministic and valid in Excel.
    u32(localParts, crc);
    u32(localParts, content.length);
    u32(localParts, content.length);
    u16(localParts, name.length);
    u16(localParts, 0);
    pushBytes(localParts, name);
    pushBytes(localParts, content);
  }

  for (const entry of entries) {
    u32(centralParts, 0x02014b50);
    u16(centralParts, 20);
    u16(centralParts, 20);
    u16(centralParts, 0x0800);
    u16(centralParts, 0);
    u16(centralParts, 0);
    u16(centralParts, 0x21);
    u32(centralParts, entry.crc);
    u32(centralParts, entry.size);
    u32(centralParts, entry.size);
    u16(centralParts, entry.name.length);
    u16(centralParts, 0);
    u16(centralParts, 0);
    u16(centralParts, 0);
    u16(centralParts, 0);
    u32(centralParts, 0);
    u32(centralParts, entry.offset);
    pushBytes(centralParts, entry.name);
  }

  const parts = [...localParts, ...centralParts];
  const centralOffset = localParts.length;
  u32(parts, 0x06054b50);
  u16(parts, 0);
  u16(parts, 0);
  u16(parts, entries.length);
  u16(parts, entries.length);
  u32(parts, centralParts.length);
  u32(parts, centralOffset);
  u16(parts, 0);

  return new Blob([new Uint8Array(parts)], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
}

function createStylesXml(formats: readonly ExcelCellFormat[]): string {
  const fonts: ExcelFont[] = [];
  const fills: ExcelFill[] = [];
  const borders: ExcelBorder[] = [];
  const numberFormats: string[] = [];
  const fontIds = new Map<string, number>();
  const fillIds = new Map<string, number>();
  const borderIds = new Map<string, number>();
  const numberFormatIds = new Map<string, number>();

  const internFont = (font: ExcelFont): number => {
    const key = styleKey(font);
    const existing = fontIds.get(key);
    if (existing !== undefined) return existing;
    const id = fonts.push(font);
    fontIds.set(key, id);
    return id;
  };
  const internFill = (fill: ExcelFill): number => {
    const key = styleKey(fill);
    const existing = fillIds.get(key);
    if (existing !== undefined) return existing;
    const id = fills.push(fill) + 1; // OOXML reserves none and gray125 as fills 0 and 1.
    fillIds.set(key, id);
    return id;
  };
  const internBorder = (border: ExcelBorder): number => {
    const key = styleKey(border);
    const existing = borderIds.get(key);
    if (existing !== undefined) return existing;
    const id = borders.push(border);
    borderIds.set(key, id);
    return id;
  };
  const internNumFmt = (pattern: string): number => {
    const key = styleKey(pattern);
    const existing = numberFormatIds.get(key);
    if (existing !== undefined) return existing;
    const id = 164 + numberFormats.length;
    numberFormats.push(pattern);
    numberFormatIds.set(key, id);
    return id;
  };

  const cellXfs = formats.map((format) => {
    const numFmtId = format.numFmt ? internNumFmt(format.numFmt) : 0;
    const fontId = format.font ? internFont(format.font) : 0;
    const fillId = format.fill ? internFill(format.fill) : 0;
    const borderId = format.border ? internBorder(format.border) : 0;
    const alignment = format.alignment
      ? `<alignment${format.alignment.horizontal ? ` horizontal="${format.alignment.horizontal}"` : ""}${format.alignment.vertical ? ` vertical="${format.alignment.vertical}"` : ""}${format.alignment.wrapText ? ' wrapText="1"' : ""}/>`
      : "";
    return [
      `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"`,
      numFmtId ? ' applyNumberFormat="1"' : "",
      fontId ? ' applyFont="1"' : "",
      fillId ? ' applyFill="1"' : "",
      borderId ? ' applyBorder="1"' : "",
      alignment ? ' applyAlignment="1"' : "",
      alignment ? `>${alignment}</xf>` : "/>"
    ].join("");
  });

  const numberFormatXml = numberFormats.map((pattern, index) =>
    `<numFmt numFmtId="${164 + index}" formatCode="${xmlEscape(pattern)}"/>`
  );
  const fontXml = fonts.map((font) =>
    [
      "<font>",
      font.bold ? "<b/>" : "",
      font.italic ? "<i/>" : "",
      font.underline ? '<u val="single"/>' : "",
      font.strike ? "<strike/>" : "",
      font.size ? `<sz val="${font.size}"/>` : "",
      font.color ? `<color rgb="${font.color}"/>` : "",
      `<name val="${xmlEscape(font.family ?? "Calibri")}"/>`,
      "</font>"
    ].join("")
  );
  const fillXml = fills.map((fill) =>
    `<fill><patternFill patternType="solid"><fgColor rgb="${fill.color ?? "FFFFFFFF"}"/></patternFill></fill>`
  );
  const borderXml = borders.map((border) => {
    const side = (name: string, item?: ExcelBorderSide) =>
      item
        ? `<${name} style="thin"><color rgb="${item.color ?? "FF000000"}"/></${name}>`
        : `<${name}/>`;
    return [
      "<border>",
      side("left", border.left),
      side("right", border.right),
      side("top", border.top),
      side("bottom", border.bottom),
      "</border>"
    ].join("");
  });
  const cellXfXml = [
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
    ...cellXfs
  ];

  return [
    XML_DECLARATION,
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    `<numFmts count="${numberFormats.length}">${numberFormatXml.join("")}</numFmts>`,
    `<fonts count="${fonts.length + 1}"><font><sz val="11"/><name val="Calibri"/></font>${fontXml.join("")}</fonts>`,
    `<fills count="${fills.length + 2}"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>${fillXml.join("")}</fills>`,
    `<borders count="${borders.length + 1}"><border/>${borderXml.join("")}</borders>`,
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
    `<cellXfs count="${cellXfXml.length}">${cellXfXml.join("")}</cellXfs>`,
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
    "</styleSheet>"
  ].join("");
}

function createCellXml(
  value: unknown,
  rowIndex: number,
  columnIndex: number,
  styleIndex: number,
  formula?: string,
  error?: string
): string {
  const reference = cellReference(rowIndex, columnIndex);
  const style = styleIndex === 0 ? "" : ` s="${styleIndex}"`;
  if (formula !== undefined) {
    const text = formula.startsWith("=") ? formula.slice(1) : formula;
    if (!text.trim() || error === "#ERROR!") {
      return `<c r="${reference}" t="e"${style}><v>#VALUE!</v></c>`;
    }
    const type = error !== undefined ? "e" : typeof value === "boolean" ? "b" : typeof value === "string" ? "str" : undefined;
    const cached = error !== undefined
      ? (["#CYCLE!", "#ERROR!"].includes(error) ? "#REF!" : error)
      : typeof value === "boolean" ? Number(value) : value;
    return `<c r="${reference}"${type === undefined ? "" : ` t="${type}"`}${style}><f>${xmlEscape(text)}</f>${cached === undefined ? "" : `<v>${xmlEscape(String(cached))}</v>`}</c>`;
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    return `<c r="${reference}"${style}><v>${value}</v></c>`;
  }
  if (typeof value === "boolean") {
    return `<c r="${reference}" t="b"${style}><v>${value ? 1 : 0}</v></c>`;
  }
  if (value === undefined || value === null || value === "") {
    return styleIndex === 0 ? "" : `<c r="${reference}"${style}/>`;
  }
  const text = xmlEscape(String(value));
  if (String(value).length > 32767) throw new RangeError("Cell text exceeds Excel's 32767-character limit");
  return `<c r="${reference}" t="inlineStr"${style}><is><t xml:space="preserve">${text}</t></is></c>`;
}

function download(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.rel = "noopener";
  link.style.display = "none";
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

const MAX_ROWS = 1_048_576;
const MAX_COLUMNS = 16_384;
const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function rangeReference(range: CellRange, rowOffset: number): string {
  if (range.startRow < 0 || range.startColumn < 0 || range.endRow + rowOffset >= MAX_ROWS || range.endColumn >= MAX_COLUMNS) {
    throw new RangeError("Exported range exceeds Excel's row or column limit");
  }
  return `${cellReference(range.startRow + rowOffset, range.startColumn)}:${cellReference(range.endRow + rowOffset, range.endColumn)}`;
}

function uniqueSheetName(input: string, used: Set<string>): string {
  const base = [...input].map(char => "\\/?*[]:".includes(char) || char.charCodeAt(0) < 32 ? "_" : char)
    .join("").trim().replace(/^'+|'+$/g, "") || "Sheet";
  let name = base.slice(0, 31);
  let suffix = 1;
  while (used.has(name.toLowerCase()) || name.toLowerCase() === "history") {
    const tail = ` (${suffix++})`;
    name = base.slice(0, 31 - tail.length) + tail;
  }
  used.add(name.toLowerCase());
  return name;
}

function sourceSheet(source: ExcelSheetExportSource): SheetSnapshot {
  const id = source.snapshot.sheetOrder[0];
  const sheet = id === undefined ? undefined : source.snapshot.sheets[id];
  if (!sheet || source.snapshot.sheetOrder.length !== 1) {
    throw new RangeError("Each Excel export source must contain exactly one sheet");
  }
  return sheet;
}

/** Captures a mounted sheet without changing its selection, filter or viewport. */
export function excelSourceFromSheet(
  sheet: CyExcelCore,
  name?: string,
  styleOf?: (row: number, column: number) => CellStyle | undefined
): ExcelSheetExportSource {
  return {
    ...(name === undefined ? {} : { name }),
    snapshot: sheet.toSnapshot(),
    columns: Array.from({ length: sheet.getColumnCount() }, (_, column) => ({
      title: sheet.getHeader(column).title,
      width: sheet.getColumn(column).width,
      wrap: sheet.getColumn(column).wrap
    })),
    ...(styleOf === undefined ? {} : { styleOf })
  };
}

/** Exports all supplied sheets in order, sharing one workbook-level style table. */
export async function exportWorkbookToExcel(
  sources: readonly ExcelSheetExportSource[],
  options: ExcelExportOptions = {}
): Promise<void | Blob> {
  if (sources.length === 0) throw new RangeError("Cannot export a workbook without sheets");
  const usedNames = new Set<string>();
  const names = sources.map(source => uniqueSheetName(source.name ?? sourceSheet(source).name, usedNames));
  const listSheetName = uniqueSheetName("_cyexcel_lists", usedNames);
  const formats: ExcelCellFormat[] = [];
  const formatIds = new Map<string, number>();
  const listRows: string[] = [];
  const definedNames: string[] = [];
  const listNames = new Map<string, string>();
  const warn = (message: string): void => {
    if (options.onWarning) options.onWarning(message);
    else console.warn(message);
  };
  const internStyle = (style: CellStyle | undefined, wrap: boolean): number => {
    const format: { -readonly [K in keyof ExcelCellFormat]?: ExcelCellFormat[K] } = {};
    const font = style && createExcelFont(style);
    const fill = style && createExcelFill(style);
    const border = style && createExcelBorder(style);
    const alignment = createExcelAlignment(style ?? {}, wrap);
    if (font) format.font = font;
    if (fill) format.fill = fill;
    if (border) format.border = border;
    if (alignment) format.alignment = alignment;
    if (style?.numFmt) format.numFmt = styleText(style.numFmt);
    if (Object.values(format).every(value => value === undefined)) return 0;
    const key = styleKey(format);
    const existing = formatIds.get(key);
    if (existing !== undefined) return existing;
    const id = formats.push(format);
    formatIds.set(key, id);
    return id;
  };
  const validationXml = (rule: DataValidationRule, offset: number, name: string): string => {
    const criteria = rule.criteria;
    if (criteria.type === "custom") {
      warn(`Sheet "${name}": custom validator "${criteria.name}" (${rule.id}) cannot be exported to Excel.`);
      return "";
    }
    const sqref = rule.ranges.map(range => rangeReference(range, offset)).join(" ");
    if (!sqref) return "";
    let formula1: string;
    let formula2: string | undefined;
    if (criteria.type === "list") {
      // A hidden range avoids Excel's 255-character inline-list limit and
      // preserves commas, quotes and line breaks inside individual options.
      const key = JSON.stringify(criteria.items);
      const existing = listNames.get(key);
      if (existing !== undefined) {
        formula1 = existing;
      } else {
        const start = listRows.length + 1;
        for (const item of criteria.items) {
          if (item.length > 32767 || listRows.length >= MAX_ROWS) {
            throw new RangeError("Data validation list exceeds Excel's cell or row limit");
          }
          const row = listRows.length;
          listRows.push(`<row r="${row + 1}">${createCellXml(item, row, 0, 0)}</row>`);
        }
        if (listRows.length < start) {
          warn(`Sheet "${name}": empty validation list (${rule.id}) was skipped.`);
          return "";
        }
        formula1 = `_cyexcel_list_${definedNames.length + 1}`;
        listNames.set(key, formula1);
        definedNames.push(`<definedName name="${formula1}">${xmlEscape(`'${listSheetName.replace(/'/g, "''")}'!$A$${start}:$A$${listRows.length}`)}</definedName>`);
      }
    } else {
      formula1 = String(criteria.value1);
      if (criteria.value2 !== undefined) formula2 = String(criteria.value2);
    }
    return [
      `<dataValidation type="${criteria.type}"`,
      criteria.type === "list" ? ` showDropDown="${rule.showDropDown === false ? 1 : 0}"` : ` operator="${criteria.operator}"`,
      ` allowBlank="${rule.allowBlank === false ? 0 : 1}" showErrorMessage="1" errorStyle="${rule.errorStyle ?? "stop"}"`,
      rule.errorTitle ? ` errorTitle="${xmlEscape(rule.errorTitle.slice(0, 32))}"` : "",
      rule.errorMessage ? ` error="${xmlEscape(rule.errorMessage.slice(0, 255))}"` : "",
      ` sqref="${sqref}"><formula1>${xmlEscape(formula1)}</formula1>`,
      formula2 === undefined ? "" : `<formula2>${xmlEscape(formula2)}</formula2>`,
      "</dataValidation>"
    ].join("");
  };

  const worksheets = sources.map((source, index) => {
    const sheet = sourceSheet(source);
    const name = names[index] as string;
    if (name !== (source.name ?? sheet.name)) warn(`Sheet "${source.name ?? sheet.name}" was exported as "${name}" to satisfy Excel's naming rules.`);
    let columnCount = source.columns?.length ?? 0;
    for (const bucket of Object.values(sheet.cellData)) {
      for (const key of Object.keys(bucket)) columnCount = Math.max(columnCount, Number(key) + 1);
    }
    for (const key of Object.keys(sheet.columnData)) columnCount = Math.max(columnCount, Number(key) + 1);
    for (const merge of sheet.mergeData) columnCount = Math.max(columnCount, merge.endColumn + 1);
    columnCount = Math.max(columnCount, 1);
    if (columnCount > MAX_COLUMNS) throw new RangeError("Sheet exceeds Excel's column limit");
    const headers = Array.from({ length: columnCount }, (_, column) => source.columns?.[column]?.title ?? columnLetter(column));
    const offset = options.includeHeaders !== false && headers.some((title, column) => title && !isDefaultColumnHeader(title, column)) ? 1 : 0;
    const rows: string[] = [];
    let imageCells = 0;
    let invalidFormulas = 0;
    if (offset) rows.push(`<row r="1">${headers.map((title, column) => createCellXml(title, 0, column, 0)).join("")}</row>`);
    const rowKeys = new Set([...Object.keys(sheet.cellData), ...Object.keys(sheet.rowData)].map(Number));
    for (const row of [...rowKeys].sort((a, b) => a - b)) {
      if (row + offset >= MAX_ROWS) throw new RangeError("Sheet exceeds Excel's row limit after adding headers");
      const bucket = sheet.cellData[row] ?? {};
      const cells = Object.keys(bucket).map(Number).sort((a, b) => a - b).map(column => {
        const record = bucket[column]!;
        if (record.img !== undefined) imageCells += 1;
        if (record.f !== undefined && (record.fe === "#ERROR!" || !record.f.replace(/^=/, "").trim())) invalidFormulas += 1;
        const columnStyleId = sheet.columnData[column]?.s;
        const columnStyle = columnStyleId === undefined ? undefined : source.snapshot.styles[columnStyleId];
        const cellStyle = record.s === undefined ? undefined : source.snapshot.styles[record.s];
        const style = source.styleOf?.(row, column) ?? (cellStyle === undefined ? columnStyle : { ...columnStyle, ...cellStyle });
        return createCellXml(record.v, row + offset, column, internStyle(style, source.columns?.[column]?.wrap ?? false),
          record.f === undefined ? undefined : shiftFormulaRows(record.f, offset), record.fe);
      }).join("");
      const meta = sheet.rowData[row];
      rows.push(`<row r="${row + offset + 1}"${meta?.h === undefined ? "" : ` ht="${meta.h * 0.75}" customHeight="1"`}${meta?.hd ? ' hidden="1"' : ""}>${cells}</row>`);
    }
    const columns = Array.from({ length: columnCount }, (_, column) => {
      const meta = sheet.columnData[column];
      const width = meta?.w ?? source.columns?.[column]?.width ?? sheet.defaultColumnWidth;
      const style = internStyle(meta?.s === undefined ? undefined : source.snapshot.styles[meta.s], source.columns?.[column]?.wrap ?? false);
      return `<col min="${column + 1}" max="${column + 1}" width="${Math.max(0, (width - 5) / 7).toFixed(2)}" customWidth="1"${style ? ` style="${style}"` : ""}${meta?.hd ? ' hidden="1"' : ""}/>`;
    });
    const merges = sheet.mergeData.map(range => `<mergeCell ref="${rangeReference(range, offset)}"/>`);
    const validations = (sheet.dataValidations ?? []).map(rule => validationXml(rule, offset, name)).filter(Boolean);
    if (imageCells) warn(`Sheet "${name}": ${imageCells} cell images were omitted; image export is not supported.`);
    if (invalidFormulas) warn(`Sheet "${name}": ${invalidFormulas} invalid formulas were exported as #VALUE! errors.`);
    return [
      XML_DECLARATION, `<worksheet xmlns="${MAIN_NS}">`,
      `<sheetFormatPr defaultRowHeight="${sheet.defaultRowHeight * 0.75}"/>`,
      `<cols>${columns.join("")}</cols><sheetData>${rows.join("")}</sheetData>`,
      merges.length ? `<mergeCells count="${merges.length}">${merges.join("")}</mergeCells>` : "",
      validations.length ? `<dataValidations count="${validations.length}">${validations.join("")}</dataValidations>` : "",
      "</worksheet>"
    ].join("");
  });
  const exportedNames = [...names];
  if (listRows.length) {
    exportedNames.push(listSheetName);
    worksheets.push(`${XML_DECLARATION}<worksheet xmlns="${MAIN_NS}"><sheetData>${listRows.join("")}</sheetData></worksheet>`);
  }
  const files = [
    { name: "[Content_Types].xml", content: `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${worksheets.map((_, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>` },
    { name: "_rels/.rels", content: `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>` },
    { name: "xl/workbook.xml", content: `${XML_DECLARATION}<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><sheets>${exportedNames.map((name, index) => `<sheet name="${xmlEscape(name)}" sheetId="${index + 1}" r:id="rId${index + 1}"${index >= names.length ? ' state="hidden"' : ""}/>`).join("")}</sheets>${definedNames.length ? `<definedNames>${definedNames.join("")}</definedNames>` : ""}<calcPr fullCalcOnLoad="1"/></workbook>` },
    { name: "xl/_rels/workbook.xml.rels", content: `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${worksheets.map((_, index) => `<Relationship Id="rId${index + 1}" Type="${REL_NS}/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`).join("")}<Relationship Id="rId${worksheets.length + 1}" Type="${REL_NS}/styles" Target="styles.xml"/></Relationships>` },
    { name: "xl/styles.xml", content: createStylesXml(formats) },
    ...worksheets.map((content, index) => ({ name: `xl/worksheets/sheet${index + 1}.xml`, content }))
  ];
  const blob = zipFiles(files);
  if (options.download === false) return blob;
  download(blob, options.fileName ?? `cyexcel-export-${Date.now()}.xlsx`);
}

/** Exports one sheet, including its merges and built-in validation rules. */
export async function exportSheetToExcel(
  sheet: CyExcelCore,
  options: ExcelExportOptions = {},
  styleOf?: (row: number, column: number) => CellStyle | undefined
): Promise<void | Blob> {
  return exportWorkbookToExcel([excelSourceFromSheet(sheet, undefined, styleOf)], options);
}
