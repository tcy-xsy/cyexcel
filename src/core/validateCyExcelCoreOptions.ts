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
 * `snapshot` 是唯一的数据入口了（ADR 0004 §3.2：数据层不可替换，对象数组走
 * 独立的 `rowsToWorkbookSnapshot`）。
 *
 * 这里只查最外一层的形状，内部结构交给 `assertValidWorkbookSnapshot`——那是
 * 快照格式自己的门禁，它同时守着「从磁盘读回来的快照」这条路，不该在入口再抄
 * 一份。数组要单独挡：`Array.isArray` 之外它也是 `typeof === "object"`，而传
 * 一个行对象数组进来（旧的 `data` 用法）恰恰是最可能出现的一种错，直接落到快照
 * 校验里报出来的会是「缺少 sheetOrder」这种指不到病根的话。
 */
function validateSnapshotShape(options: CyExcelCoreOptions): void {
  if (options.snapshot === undefined) {
    return;
  }
  if (
    typeof options.snapshot !== "object" ||
    options.snapshot === null ||
    Array.isArray(options.snapshot)
  ) {
    throw invalidOption(
      "snapshot",
      options.snapshot,
      "expected a WorkbookSnapshot — build one from rows with rowsToWorkbookSnapshot()"
    );
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

  validateSnapshotShape(options);
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
 * 入口收成 `snapshot` 之后，「列数对不上」重新变成一条**调用方造得出来的**
 * 错误：快照的列数由建它的那次导入决定（`rowsToWorkbookSnapshot` 的 `fields`），
 * 与这里配的 `columns` 是两处独立的输入。对不上而不拦，后果是渲染层按
 * `columns.length` 画、存储按自己的列数存，多出来的那一列永远读不到值——静默。
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
