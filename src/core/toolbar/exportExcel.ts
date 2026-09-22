import type { CyExcelCore } from "../CyExcelCore";
import type { BorderSide, CellStyle } from "../style/types";

export interface ExcelExportOptions {
  readonly fileName?: string;
  readonly includeHeaders?: boolean;
  /** Return the XLSX Blob instead of triggering a browser download. */
  readonly download?: boolean;
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

  const probe = document.createElement("span");
  probe.style.color = source;
  const match = /^rgba?\((\d+), (\d+), (\d+)(?:, ([\d.]+))?\)$/i.exec(probe.style.color);
  if (!match) return undefined;
  const [, red, green, blue, alpha] = match;
  const alphaHex = Math.round(Number(alpha ?? "1") * 255)
    .toString(16)
    .padStart(2, "0");
  return `${alphaHex}${red?.padStart(2, "0")}${green?.padStart(2, "0")}${blue?.padStart(2, "0")}`.toUpperCase();
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
      side("top", border.top),
      side("left", border.left),
      side("bottom", border.bottom),
      side("right", border.right),
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
    "</styleSheet>"
  ].join("");
}

function createCellXml(
  value: unknown,
  rowIndex: number,
  columnIndex: number,
  styleIndex: number
): string {
  const reference = cellReference(rowIndex, columnIndex);
  const style = styleIndex === 0 ? "" : ` s="${styleIndex}"`;
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

/** Exports the current sheet's used area as a small, dependency-free XLSX file. */
export async function exportSheetToExcel(
  sheet: CyExcelCore,
  options: ExcelExportOptions = {},
  styleOf?: (row: number, column: number) => CellStyle | undefined
): Promise<void | Blob> {
  const rowCount = Math.max(sheet.getUsedRowCount(), 1);
  const columnCount = sheet.getColumnCount();
  if (rowCount <= 0 || columnCount <= 0) {
    throw new RangeError("Cannot export an empty coordinate space");
  }

  const headers = Array.from({ length: columnCount }, (_unused, column) =>
    sheet.getHeader(column)
  );
  const hasCustomHeaders = options.includeHeaders !== false &&
    headers.some((header, column) =>
      header.title.length > 0 && !isDefaultColumnHeader(header.title, column)
    );
  const firstDataRow = hasCustomHeaders ? 1 : 0;

  const formats: ExcelCellFormat[] = [];
  const formatIds = new Map<string, number>();
  const styleIndexes: number[][] = [];
  for (let row = 0; row < rowCount; row += 1) {
    const sourceRow = hasCustomHeaders ? row - 1 : row;
    const rowStyles: number[] = [];
    for (let column = 0; column < columnCount; column += 1) {
      const style = sourceRow >= 0 ? styleOf?.(sourceRow, column) : undefined;
      if (!style) {
        rowStyles.push(0);
        continue;
      }

      const format: { -readonly [K in keyof ExcelCellFormat]?: ExcelCellFormat[K] } = {};
      const font = createExcelFont(style);
      const fill = createExcelFill(style);
      const border = createExcelBorder(style);
      const alignment = createExcelAlignment(style, sheet.getColumn(column).wrap);
      const numFmt = styleText(style.numFmt ?? "");
      if (font) format.font = font;
      if (fill) format.fill = fill;
      if (border) format.border = border;
      if (alignment) format.alignment = alignment;
      if (numFmt) format.numFmt = numFmt;

      if (Object.keys(format).length === 0) {
        rowStyles.push(0);
        continue;
      }
      const key = styleKey(format);
      const existing = formatIds.get(key);
      if (existing !== undefined) {
        rowStyles.push(existing);
        continue;
      }
      const id = formats.push(format);
      formatIds.set(key, id);
      rowStyles.push(id);
    }
    styleIndexes.push(rowStyles);
  }

  const rows: string[] = [];
  if (hasCustomHeaders) {
    const cells = headers
      .map((header, column) => createCellXml(header.title, 0, column, 0))
      .join("");
    rows.push(`<row r="1">${cells}</row>`);
  }
  for (let row = firstDataRow; row < rowCount; row += 1) {
    const sourceRow = hasCustomHeaders ? row - 1 : row;
    const cells = Array.from({ length: columnCount }, (_unused, column) =>
      createCellXml(
        sheet.getCell(sourceRow, column),
        row,
        column,
        styleIndexes[row]![column]!
      )
    ).join("");
    rows.push(`<row r="${row + 1}">${cells}</row>`);
  }

  const columns = Array.from({ length: columnCount }, (_unused, column) => {
    const width = Math.max(0, (sheet.getColumn(column).width - 5) / 7);
    return `<col min="${column + 1}" max="${column + 1}" width="${width.toFixed(2)}" customWidth="1"/>`;
  });
  const files = [
    {
      name: "[Content_Types].xml",
      content: [
        XML_DECLARATION,
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
        "</Types>"
      ].join("")
    },
    {
      name: "_rels/.rels",
      content: [
        XML_DECLARATION,
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
        "</Relationships>"
      ].join("")
    },
    {
      name: "xl/workbook.xml",
      content: [
        XML_DECLARATION,
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>',
        "</workbook>"
      ].join("")
    },
    {
      name: "xl/_rels/workbook.xml.rels",
      content: [
        XML_DECLARATION,
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>',
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
        "</Relationships>"
      ].join("")
    },
    { name: "xl/styles.xml", content: createStylesXml(formats) },
    {
      name: "xl/worksheets/sheet1.xml",
      content: [
        XML_DECLARATION,
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        `<cols>${columns.join("")}</cols>`,
        `<sheetData>${rows.join("")}</sheetData>`,
        "</worksheet>"
      ].join("")
    }
  ];

  const blob = zipFiles(files);
  if (options.download === false) {
    return blob;
  }

  download(blob, options.fileName ?? `cyexcel-export-${Date.now()}.xlsx`);
}
