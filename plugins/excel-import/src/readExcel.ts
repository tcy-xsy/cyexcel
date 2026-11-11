import type { Cell, CellValue as ExcelValue, Color, RowModel, Style, Workbook } from "exceljs";
import type { CyExcelCore, WorkbookSheetOptions, WorkbookSnapshot } from "cyexcel";
import type { CellRange, CellRecord, SheetSnapshot } from "cyexcel/core";

type CellStyle = NonNullable<ReturnType<CyExcelCore["getCellStyle"]>>;

export type ExcelImportInput = Blob | ArrayBuffer | Uint8Array;

export interface ExcelImportOptions {
  /** Cancellation is checked before applying any data to a workbook. */
  readonly signal?: AbortSignal;
  /** Maximum compressed file size in bytes. Defaults to 25 MiB. */
  readonly maxFileSize?: number;
  /** Maximum populated/styled cells across all sheets. Defaults to 200,000. */
  readonly maxCells?: number;
  readonly onWarning?: (message: string) => void;
}

export interface ExcelImportResult {
  /** One single-sheet snapshot per tab, suitable for WorkbookHandle.replaceSheets(). */
  readonly sheets: readonly WorkbookSheetOptions[];
  readonly warnings: readonly string[];
}

const DEFAULT_THEME = ["FFFFFF", "000000", "EEECE1", "1F497D", "4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646", "0000FF", "800080"];

function themeColors(workbook: Workbook): readonly string[] {
  const theme = (workbook.model.themes as unknown as Record<string, string> | undefined)?.theme1;
  if (!theme) return DEFAULT_THEME;
  return ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"].map((key, index) => {
    const block = new RegExp(`<(?:\\w+:)?${key}>([\\s\\S]*?)</(?:\\w+:)?${key}>`).exec(theme)?.[1];
    return block && /(?:lastClr|val)="([A-Fa-f0-9]{6})"/.exec(block)?.[1] || DEFAULT_THEME[index]!;
  });
}

function cssColor(color: (Partial<Color> & { tint?: number }) | undefined, theme: readonly string[]): string | undefined {
  if (!color) return undefined;
  const argb = color.argb;
  const hex = argb ? argb.slice(-6) : color.theme === undefined ? undefined : theme[color.theme];
  if (!hex || !/^[a-f\d]{6}$/i.test(hex)) return undefined;
  const tint = color.tint ?? 0;
  const rgb = [0, 2, 4].map(offset => {
    const value = Number.parseInt(hex.slice(offset, offset + 2), 16);
    return Math.round(tint < 0 ? value * (1 + tint) : value + (255 - value) * tint).toString(16).padStart(2, "0");
  }).join("");
  return `#${rgb}`;
}

function cellStyle(source: Partial<Style>, theme: readonly string[]): CellStyle {
  const style: { -readonly [K in keyof CellStyle]?: CellStyle[K] } = {};
  const font = source.font;
  if (font) {
    if (font.bold !== undefined) style.bold = font.bold;
    if (font.italic !== undefined) style.italic = font.italic;
    if (font.underline !== undefined) style.underline = !!font.underline;
    if (font.strike !== undefined) style.strike = font.strike;
    if (font.name) style.fontFamily = font.name;
    if (font.size) style.fontSize = font.size / 0.75;
    const color = cssColor(font.color, theme);
    if (color) style.color = color;
  }
  if (source.fill?.type === "pattern" && source.fill.pattern === "solid") {
    const color = cssColor(source.fill.fgColor, theme);
    if (color) style.background = color;
  }
  const horizontal = source.alignment?.horizontal;
  if (horizontal === "left" || horizontal === "center" || horizontal === "right") style.align = horizontal;
  const vertical = source.alignment?.vertical;
  if (vertical === "middle" || vertical === "top" || vertical === "bottom") style.verticalAlign = vertical;
  if (source.numFmt) style.numFmt = source.numFmt;
  if (source.border) {
    const border: { -readonly [K in keyof NonNullable<CellStyle["border"]>]?: NonNullable<CellStyle["border"]>[K] } = {};
    for (const side of ["top", "right", "bottom", "left"] as const) {
      const sourceSide = source.border[side];
      if (sourceSide?.style) border[side] = {
        color: cssColor(sourceSide.color, theme) ?? "#000000",
        width: sourceSide.style === "thick" ? 3 : sourceSide.style === "medium" || sourceSide.style === "double" ? 2 : 1,
        style: "solid"
      };
    }
    if (Object.keys(border).length) style.border = border;
  }
  return style;
}

function scalar(value: ExcelValue): CellRecord["v"] {
  if (value === null || value === undefined) return undefined;
  if (value instanceof Date) return value.getTime() / 86400000 + 25569;
  if (typeof value !== "object") return value;
  if ("richText" in value) return value.richText.map(part => part.text).join("");
  if ("text" in value) return value.text;
  if ("error" in value) return value.error;
  if ("result" in value) return scalar(value.result);
  return undefined;
}

function recordOf(cell: Cell): CellRecord {
  const formula = cell.formula;
  const value = cell.value;
  const result = formula && value && typeof value === "object" && "result" in value ? value.result : value;
  const v = scalar(result);
  return {
    v,
    ...(typeof v === "string" ? { t: 1 as const } : typeof v === "boolean" ? { t: 3 as const } : {}),
    ...(formula ? { f: `=${formula}` } : {}),
    ...(formula && result && typeof result === "object" && "error" in result ? { fe: result.error } : {})
  };
}

function rangeOf(reference: string): CellRange {
  const match = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i.exec(reference);
  if (!match) throw new Error(`Invalid Excel range: ${reference}`);
  const column = (name: string): number => [...name.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
  return { startRow: Number(match[2]) - 1, startColumn: column(match[1]!), endRow: Number(match[4] ?? match[2]) - 1, endColumn: column(match[3] ?? match[1]!) };
}

/** Parses XLSX locally. No workbook is modified until the caller applies the result. */
export async function readExcelWorkbook(input: ExcelImportInput, options: ExcelImportOptions = {}): Promise<ExcelImportResult> {
  options.signal?.throwIfAborted();
  const maxFileSize = options.maxFileSize ?? 25 * 1024 * 1024;
  const maxCells = options.maxCells ?? 200000;
  for (const [name, limit] of Object.entries({ maxFileSize, maxCells })) {
    if (!Number.isSafeInteger(limit) || limit <= 0) throw new RangeError(`${name} must be a positive integer`);
  }
  const size = input instanceof Uint8Array ? input.byteLength : input instanceof ArrayBuffer ? input.byteLength : input.size;
  if (size > maxFileSize) throw new RangeError(`Excel file exceeds the ${maxFileSize}-byte limit`);
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input instanceof ArrayBuffer ? input : await input.arrayBuffer());
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) throw new Error("Only unencrypted .xlsx files are supported");
  const ExcelJS = await import("exceljs");
  const workbook = new (ExcelJS.default ?? ExcelJS).Workbook();
  try {
    // ExcelJS accepts browser ArrayBuffers despite its Node-oriented Buffer declaration.
    await workbook.xlsx.load(bytes as unknown as Parameters<typeof workbook.xlsx.load>[0]);
  } catch {
    throw new Error("Unable to read this Excel file. Select a valid, unencrypted .xlsx file.");
  }
  if (!workbook.worksheets.length) throw new Error("The Excel file contains no worksheets");
  options.signal?.throwIfAborted();
  const warnings = new Set<string>();
  const warn = (message: string): void => { warnings.add(message); };
  const theme = themeColors(workbook);
  let cellCount = 0;
  const sheets = workbook.worksheets.map(worksheet => {
    const model = worksheet.model as typeof worksheet.model & { rows: readonly RowModel[] };
    if (worksheet.state !== "visible") warn("Hidden worksheets are imported as visible tabs.");
    if (worksheet.getImages().length) warn("Images are not imported.");
    const validations = (worksheet as unknown as { dataValidations: { model: Record<string, unknown> } }).dataValidations.model;
    if (Object.keys(validations).length) warn("Data validation rules are not imported.");
    if (model.media?.length) warn("Drawing objects are not imported.");
    const id = `excel-sheet-${worksheet.id}`;
    const cellData: Record<number, Record<number, CellRecord>> = {};
    const rowData: Record<number, SheetSnapshot["rowData"][number]> = {};
    const columnData: Record<number, SheetSnapshot["columnData"][number]> = {};
    const styles: Record<string, CellStyle> = {};
    const styleIds = new Map<string, string>();
    const internStyle = (source: Partial<Style>): string | undefined => {
      const style = cellStyle(source, theme);
      if (!Object.keys(style).length) return undefined;
      const key = JSON.stringify(style);
      let styleId = styleIds.get(key);
      if (!styleId) {
        styleId = `excel-style-${styleIds.size}`;
        styleIds.set(key, styleId);
        styles[styleId] = style;
      }
      return styleId;
    };
    const wrappedColumns = new Set<number>();
    worksheet.columns?.forEach((column, index) => {
      const s = internStyle(column.style ?? {});
      columnData[index] = {
        ...(column.width === undefined ? {} : { w: Math.max(1, column.width * 7 + 5) }),
        ...(column.hidden ? { hd: true } : {}),
        ...(s === undefined ? {} : { s })
      };
      if (column.alignment?.wrapText) wrappedColumns.add(index);
    });
    // Model rows/cells are sparse and include styled blanks; eachRow() skips blank rows.
    for (const rowModel of model.rows) {
      const row = worksheet.getRow(rowModel.number);
      const r = rowModel.number - 1;
      if (row.height !== undefined || row.hidden) rowData[r] = {
        ...(row.height === undefined || row.height <= 0 ? {} : { h: row.height / 0.75, hp: "manual" }),
        ...(row.hidden || row.height === 0 ? { hd: true } : {})
      };
      for (const cellModel of rowModel.cells) {
        const cell = worksheet.getCell(cellModel.address as unknown as string);
        if (++cellCount > maxCells) throw new RangeError(`Excel file exceeds the ${maxCells}-cell import limit`);
        const c = Number(cell.col) - 1;
        const record = recordOf(cell);
        const s = internStyle(cell.style);
        if (cell.alignment?.wrapText) wrappedColumns.add(c);
        if (cell.formula?.includes("!")) warn("Cross-sheet formulas are preserved, but the current formula engine cannot calculate them.");
        if (cell.value && typeof cell.value === "object" && "hyperlink" in cell.value) warn("Hyperlinks are imported as display text.");
        if (cell.value && typeof cell.value === "object" && "richText" in cell.value) warn("Rich text is imported as plain text.");
        if (cell.isMerged && cell.master.address !== cell.address) {
          // ExcelJS exposes the master's value on merged children. Only retain their own styles.
          if (s) (cellData[r] ??= {})[c] = { v: undefined, s };
        } else if (record.v !== undefined || record.f || s) {
          (cellData[r] ??= {})[c] = { ...record, ...(s === undefined ? {} : { s }) };
        }
      }
    }
    const mergeData = model.merges.map(rangeOf);
    const freeze = worksheet.views?.find(view => view.state === "frozen");
    const sheet: SheetSnapshot = {
      id, name: worksheet.name, blankRows: 100, blankColumns: 26,
      defaultRowHeight: (worksheet.properties.defaultRowHeight || 15) / 0.75,
      defaultColumnWidth: (worksheet.properties.defaultColWidth ?? 12) * 7 + 5,
      cellData, rowData, columnData, mergeData,
      ...(freeze?.state === "frozen" ? { freeze: { rows: freeze.ySplit ?? 0, columns: freeze.xSplit ?? 0 } } : {})
    };
    let maxColumn = worksheet.columnCount;
    for (const range of mergeData) maxColumn = Math.max(maxColumn, range.endColumn + 1);
    // Supply fresh column configs so host-specific headers, renderers and editors do not leak into imported tabs.
    const columns = Array.from({ length: Math.max(26, maxColumn + 26) }, (_, index) => wrappedColumns.has(index) ? { wrap: true } : {});
    const snapshot: WorkbookSnapshot = { version: 1, sheetOrder: [id], sheets: { [id]: sheet }, styles, resources: [] };
    return { name: worksheet.name, snapshot, columns, columnLetterFallback: true } satisfies WorkbookSheetOptions;
  });
  for (const warning of warnings) options.onWarning?.(warning);
  return { sheets, warnings: [...warnings] };
}
