import type { CellRange } from "../types";
import { computeFillValues } from "./fillValues";
import type { CellValue } from "../sheet/types";

export interface FillPlanSources {
  readonly getCell: (row: number, column: number) => CellValue | undefined;
  /**
   * Columns with `clipboard.paste === false` keep their existing values in
   * the extended region — the same semantic `ClipboardService.paste()` uses
   * for read-only columns.
   */
  readonly isPasteBlocked: (column: number) => boolean;
}

/**
 * Builds the value matrix a fill-handle result (drag or double-click
 * auto-fill) writes over `targetRange`, as one SetRangeCommand payload — a
 * single undo step, the same convention paste and the row-structure commands
 * use.
 *
 * `targetRange` always extends `sourceRange` down or right, never both (see
 * FillHandleController), so exactly one of the two extension branches runs.
 * The source region itself is always copied verbatim; only the newly covered
 * rows or columns are extrapolated by `computeFillValues()`.
 *
 * Pure: every read goes through `sources`, and nothing is committed here.
 */
export function computeFillPlan(
  sourceRange: CellRange,
  targetRange: CellRange,
  sources: FillPlanSources
): (CellValue | undefined)[][] {
  const rows = targetRange.endRow - targetRange.startRow + 1;
  const columns = targetRange.endColumn - targetRange.startColumn + 1;
  const values: (CellValue | undefined)[][] = Array.from(
    { length: rows },
    () => new Array<CellValue | undefined>(columns).fill(undefined)
  );

  for (let row = sourceRange.startRow; row <= sourceRange.endRow; row += 1) {
    for (let column = sourceRange.startColumn; column <= sourceRange.endColumn; column += 1) {
      values[row - targetRange.startRow]![column - targetRange.startColumn] =
        sources.getCell(row, column);
    }
  }

  if (targetRange.endRow > sourceRange.endRow) {
    const newRowCount = targetRange.endRow - sourceRange.endRow;
    for (let column = sourceRange.startColumn; column <= sourceRange.endColumn; column += 1) {
      const relativeColumn = column - targetRange.startColumn;
      const readOnly = sources.isPasteBlocked(column);
      const known = Array.from(
        { length: sourceRange.endRow - sourceRange.startRow + 1 },
        (_unused, index) => sources.getCell(sourceRange.startRow + index, column)
      );
      const filled = readOnly ? null : computeFillValues(known, newRowCount);
      for (let index = 0; index < newRowCount; index += 1) {
        const row = sourceRange.endRow + 1 + index;
        values[row - targetRange.startRow]![relativeColumn] = readOnly
          ? sources.getCell(row, column)
          : filled![index];
      }
    }
  } else if (targetRange.endColumn > sourceRange.endColumn) {
    const newColumnCount = targetRange.endColumn - sourceRange.endColumn;
    for (let row = sourceRange.startRow; row <= sourceRange.endRow; row += 1) {
      const known = Array.from(
        { length: sourceRange.endColumn - sourceRange.startColumn + 1 },
        (_unused, index) => sources.getCell(row, sourceRange.startColumn + index)
      );
      const filled = computeFillValues(known, newColumnCount);
      for (let index = 0; index < newColumnCount; index += 1) {
        const column = sourceRange.endColumn + 1 + index;
        const relativeColumn = column - targetRange.startColumn;
        values[row - targetRange.startRow]![relativeColumn] = sources.isPasteBlocked(column)
          ? sources.getCell(row, column)
          : filled[index];
      }
    }
  }

  return values;
}

export interface AutoFillDownBounds {
  readonly columnCount: number;
  readonly rowCount: number;
}

/**
 * Excel's double-click-fill-handle reach: how far down the selection should
 * extend to match the length of the nearest non-empty neighbor column (left,
 * or right if the selection is already at column 0). Scans down from the
 * selection's own first row until an empty cell ends the neighbor's run.
 *
 * Returns `null` when there is no neighbor column to follow, or when the run
 * does not reach past the selection's existing last row — both meaning
 * "nothing to fill".
 */
export function computeAutoFillDownTargetRow(
  range: CellRange,
  bounds: AutoFillDownBounds,
  getCell: (row: number, column: number) => CellValue | undefined
): number | null {
  const neighborColumn =
    range.startColumn > 0 ? range.startColumn - 1 : range.startColumn + 1;
  if (neighborColumn < 0 || neighborColumn >= bounds.columnCount) {
    return null;
  }

  let targetRow = range.endRow;
  for (let row = range.startRow; row < bounds.rowCount; row += 1) {
    const value = getCell(row, neighborColumn);
    if (value === null || value === undefined || value === "") {
      break;
    }
    targetRow = row;
  }
  return targetRow > range.endRow ? targetRow : null;
}
