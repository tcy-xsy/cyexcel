import { CyExcelError, CyExcelErrorCode } from "./errors";
import type { CellRange } from "./types";

function isNonNegativeInteger(value: number): boolean {
  return Number.isInteger(value) && value >= 0;
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
