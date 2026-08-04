import { CyExcelError, CyExcelErrorCode } from "./errors";
import type { CellRange, CyExcelOptions } from "./types";

function isNonNegativeInteger(value: number): boolean {
  return Number.isInteger(value) && value >= 0;
}

function isNonNegativeFinite(value: number): boolean {
  return Number.isFinite(value) && value >= 0;
}

export function assertValidCoordinate(row: number, column: number): void {
  if (!isNonNegativeInteger(row) || !isNonNegativeInteger(column)) {
    throw new CyExcelError(
      CyExcelErrorCode.InvalidCoordinate,
      `Invalid cell coordinate: row=${row}, column=${column}`
    );
  }
}

export function assertValidRange(range: CellRange): void {
  const { startRow, startColumn, endRow, endColumn } = range;

  if (
    !isNonNegativeInteger(startRow) ||
    !isNonNegativeInteger(startColumn) ||
    !isNonNegativeInteger(endRow) ||
    !isNonNegativeInteger(endColumn)
  ) {
    throw new CyExcelError(
      CyExcelErrorCode.InvalidRange,
      `Invalid cell range: ${JSON.stringify(range)}`
    );
  }

  if (startRow > endRow || startColumn > endColumn) {
    throw new CyExcelError(
      CyExcelErrorCode.InvalidRange,
      `Cell range start must not be after end: ${JSON.stringify(range)}`
    );
  }
}

export function assertValidCyExcelOptions(options: CyExcelOptions): void {
  if (options.rowHeight !== undefined && !isNonNegativeFinite(options.rowHeight)) {
    throw invalidOption("rowHeight", options.rowHeight);
  }

  if (
    options.headerHeight !== undefined &&
    !isNonNegativeFinite(options.headerHeight)
  ) {
    throw invalidOption("headerHeight", options.headerHeight);
  }

  if (
    options.overscanRows !== undefined &&
    !isNonNegativeInteger(options.overscanRows)
  ) {
    throw invalidOption("overscanRows", options.overscanRows);
  }

  if (
    options.overscanColumns !== undefined &&
    !isNonNegativeInteger(options.overscanColumns)
  ) {
    throw invalidOption("overscanColumns", options.overscanColumns);
  }

  if (
    options.freeze?.rows !== undefined &&
    !isNonNegativeInteger(options.freeze.rows)
  ) {
    throw invalidOption("freeze.rows", options.freeze.rows);
  }

  if (
    options.freeze?.columns !== undefined &&
    !isNonNegativeInteger(options.freeze.columns)
  ) {
    throw invalidOption("freeze.columns", options.freeze.columns);
  }
}

function invalidOption(field: string, value: number): CyExcelError {
  return new CyExcelError(
    CyExcelErrorCode.InvalidOptions,
    `Invalid CyExcelOptions.${field}: ${value}`
  );
}
