import { CyExcelError, CyExcelErrorCode } from "./errors";
import type { CyExcelCoreOptions } from "./CyExcelCore";

export interface CyExcelCoreDataDimensions {
  readonly rowCount: number;
  readonly columnCount: number;
}

function invalidOption(field: string, value: unknown, detail?: string): CyExcelError {
  const rendered = typeof value === "string" ? JSON.stringify(value) : String(value);
  return new CyExcelError(
    CyExcelErrorCode.InvalidOptions,
    `Invalid CyExcelCoreOptions.${field}: ${rendered}${detail ? ` (${detail})` : ""}`
  );
}

function isNonNegativeSafeInteger(value: number): boolean {
  return Number.isSafeInteger(value) && value >= 0;
}

function isNonNegativeFinite(value: number): boolean {
  return Number.isFinite(value) && value >= 0;
}

function isPositiveFinite(value: number): boolean {
  return Number.isFinite(value) && value > 0;
}

function validateOptionalBoolean(
  field: string,
  value: boolean | undefined
): void {
  if (value !== undefined && typeof value !== "boolean") {
    throw invalidOption(field, value, "expected a boolean");
  }
}

function validateOptionalNonNegativeInteger(
  field: string,
  value: number | undefined
): void {
  if (value !== undefined && !isNonNegativeSafeInteger(value)) {
    throw invalidOption(field, value, "expected a non-negative safe integer");
  }
}

function validateOptionalNonNegativeFinite(
  field: string,
  value: number | undefined
): void {
  if (value !== undefined && !isNonNegativeFinite(value)) {
    throw invalidOption(field, value, "expected a non-negative finite number");
  }
}

function validateColumnGeometry(options: CyExcelCoreOptions): void {
  for (const [index, column] of (options.columns ?? []).entries()) {
    validateOptionalNonNegativeFinite(`columns[${index}].width`, column.width);
    validateOptionalNonNegativeFinite(`columns[${index}].minWidth`, column.minWidth);
    if (
      column.maxWidth !== undefined &&
      column.maxWidth !== Number.POSITIVE_INFINITY &&
      !isNonNegativeFinite(column.maxWidth)
    ) {
      throw invalidOption(
        `columns[${index}].maxWidth`,
        column.maxWidth,
        "expected a non-negative finite number or positive infinity"
      );
    }
    if (
      column.minWidth !== undefined &&
      column.maxWidth !== undefined &&
      column.maxWidth < column.minWidth
    ) {
      throw invalidOption(
        `columns[${index}].maxWidth`,
        column.maxWidth,
        `must be greater than or equal to minWidth ${column.minWidth}`
      );
    }
  }
}

function validateRowHeights(options: CyExcelCoreOptions): void {
  if (options.rowHeights === undefined) {
    return;
  }
  const candidate = options.rowHeights as unknown as {
    readonly get?: unknown;
    readonly [Symbol.iterator]?: unknown;
  };
  if (
    typeof candidate !== "object" ||
    candidate === null ||
    typeof candidate.get !== "function" ||
    typeof candidate[Symbol.iterator] !== "function"
  ) {
    throw invalidOption("rowHeights", options.rowHeights, "expected a ReadonlyMap");
  }
  for (const [row, height] of options.rowHeights) {
    if (!isNonNegativeSafeInteger(row)) {
      throw invalidOption(
        "rowHeights key",
        row,
        "expected a non-negative safe integer row index"
      );
    }
    if (!isPositiveFinite(height)) {
      throw invalidOption("rowHeights value", height, "expected a positive finite height");
    }
  }
}

/**
 * `data` 是唯一的数据入口了：`dataSource` 在 ADR 0004 阶段 4B 删除，数据层
 * 不再可替换（§3.2），所以这里也没有「鸭子类型校验一个外来对象」这件事，
 * 只剩下形状检查。
 */
function validateDataShape(options: CyExcelCoreOptions): void {
  if (options.data !== undefined && !Array.isArray(options.data)) {
    throw invalidOption("data", options.data, "expected an array");
  }
}

function validatePlugins(options: CyExcelCoreOptions): void {
  for (const [index, plugin] of (options.plugins ?? []).entries()) {
    if (typeof plugin !== "object" || plugin === null) {
      throw invalidOption(`plugins[${index}]`, plugin, "expected a plugin object");
    }
    if (typeof plugin.name !== "string") {
      throw invalidOption(`plugins[${index}].name`, plugin.name, "expected a string");
    }
    if (typeof plugin.setup !== "function") {
      throw invalidOption(`plugins[${index}].setup`, plugin.setup, "expected a function");
    }
  }
}

/**
 * Pure, integration-boundary validation for options Core reads directly.
 * It deliberately runs before DOM construction and plugin setup. Semantic
 * column/header/type validation stays with the owners that resolve it.
 */
export function validateCyExcelCoreOptions(options: CyExcelCoreOptions): void {
  if (typeof options !== "object" || options === null) {
    throw invalidOption("options", options, "expected an options object");
  }
  if (options.columns !== undefined && !Array.isArray(options.columns)) {
    throw invalidOption("columns", options.columns, "expected an array");
  }
  if (options.headerGroups !== undefined && !Array.isArray(options.headerGroups)) {
    throw invalidOption("headerGroups", options.headerGroups, "expected an array");
  }
  if (options.plugins !== undefined && !Array.isArray(options.plugins)) {
    throw invalidOption("plugins", options.plugins, "expected an array");
  }
  if (
    options.freeze !== undefined &&
    (typeof options.freeze !== "object" || options.freeze === null)
  ) {
    throw invalidOption("freeze", options.freeze, "expected an object");
  }
  if (
    options.columnTypes !== undefined &&
    (typeof options.columnTypes !== "object" ||
      options.columnTypes === null ||
      Array.isArray(options.columnTypes))
  ) {
    throw invalidOption("columnTypes", options.columnTypes, "expected a record");
  }
  if (
    options.clipboard !== undefined &&
    (typeof options.clipboard !== "object" || options.clipboard === null)
  ) {
    throw invalidOption("clipboard", options.clipboard, "expected an object");
  }

  validateDataShape(options);
  validatePlugins(options);
  validateColumnGeometry(options);

  if (options.rowHeight !== undefined && !isPositiveFinite(options.rowHeight)) {
    throw invalidOption("rowHeight", options.rowHeight, "expected a positive finite height");
  }
  validateRowHeights(options);
  validateOptionalNonNegativeFinite("headerHeight", options.headerHeight);
  validateOptionalNonNegativeFinite("rowHeaderWidth", options.rowHeaderWidth);
  validateOptionalNonNegativeInteger("overscanRows", options.overscanRows);
  validateOptionalNonNegativeInteger("overscanColumns", options.overscanColumns);
  validateOptionalNonNegativeInteger("freeze.rows", options.freeze?.rows);
  validateOptionalNonNegativeInteger("freeze.columns", options.freeze?.columns);

  if (
    options.scrollCommitPolicy !== undefined &&
    options.scrollCommitPolicy !== "commit" &&
    options.scrollCommitPolicy !== "cancel"
  ) {
    throw invalidOption("scrollCommitPolicy", options.scrollCommitPolicy);
  }
  if (options.editorFactory !== undefined && typeof options.editorFactory !== "function") {
    throw invalidOption("editorFactory", options.editorFactory, "expected a function");
  }
  validateOptionalBoolean("columnLetterFallback", options.columnLetterFallback);
  validateOptionalBoolean("rowReorder", options.rowReorder);
  validateOptionalBoolean("showGridLines", options.showGridLines);

  if (options.clipboard?.maxCells !== undefined) {
    validateOptionalNonNegativeInteger("clipboard.maxCells", options.clipboard.maxCells);
  }
  if (options.clipboard?.maxBytes !== undefined) {
    validateOptionalNonNegativeInteger("clipboard.maxBytes", options.clipboard.maxBytes);
  }
  if (
    options.clipboard?.oversizedAction !== undefined &&
    options.clipboard.oversizedAction !== "confirm-or-download" &&
    options.clipboard.oversizedAction !== "download" &&
    options.clipboard.oversizedAction !== "block"
  ) {
    throw invalidOption("clipboard.oversizedAction", options.clipboard.oversizedAction);
  }
}

/**
 * 在建 DOM 之前校验模型报出的行列数与 `rowHeights` 的键。
 *
 * 数据层不可替换之后，前三条检查不再是「校验外来实现」而是核对导入的结果
 * ——`rowsToWorkbookSnapshot` 承诺 `rowCount === data.length`、
 * `columnCount === columns.length`，这里是那条承诺的门禁。留着而不是删掉，
 * 是因为它一旦破就会以「渲染少了一行」的形式静默出现（4B 设计 §4.2）。
 */
export function validateCyExcelCoreDataDimensions(
  options: CyExcelCoreOptions,
  dimensions: CyExcelCoreDataDimensions,
  configuredColumnCount: number
): void {
  if (!isNonNegativeSafeInteger(dimensions.rowCount)) {
    throw invalidOption(
      "sheet.rowCount()",
      dimensions.rowCount,
      "expected a non-negative safe integer"
    );
  }
  if (!isNonNegativeSafeInteger(dimensions.columnCount)) {
    throw invalidOption(
      "sheet.columnCount()",
      dimensions.columnCount,
      "expected a non-negative safe integer"
    );
  }
  if (dimensions.columnCount !== configuredColumnCount) {
    throw invalidOption(
      "sheet.columnCount()",
      dimensions.columnCount,
      `must equal columns.length ${configuredColumnCount}`
    );
  }
  for (const [row] of options.rowHeights ?? []) {
    if (row >= dimensions.rowCount) {
      throw invalidOption(
        "rowHeights key",
        row,
        `must be less than row count ${dimensions.rowCount}`
      );
    }
  }
}

/** Ensures repeated group bands cannot turn a finite band height into Infinity. */
export function validateCyExcelCoreHeaderGeometry(
  headerBandHeight: number,
  headerGroupDepth: number
): void {
  if (!isNonNegativeSafeInteger(headerGroupDepth)) {
    throw invalidOption("headerGroups depth", headerGroupDepth);
  }
  const totalHeight = headerBandHeight * (headerGroupDepth + 1);
  if (!Number.isFinite(totalHeight)) {
    throw invalidOption(
      "headerHeight",
      headerBandHeight,
      `total height over ${headerGroupDepth + 1} header bands is not finite`
    );
  }
}
