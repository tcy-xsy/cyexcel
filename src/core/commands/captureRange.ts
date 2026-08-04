import type { TableStore } from "../data";
import type { CellRange } from "../types";

/**
 * Reads the current values inside a range without touching anything outside
 * it — used by SetRangeCommand/ClearRangeCommand to capture just enough
 * state for undo, never a full-table snapshot.
 */
export function captureRange(store: TableStore, range: CellRange): unknown[][] {
  const rows: unknown[][] = [];

  for (let row = range.startRow; row <= range.endRow; row += 1) {
    const rowValues: unknown[] = [];
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      rowValues.push(store.getCell(row, column));
    }
    rows.push(rowValues);
  }

  return rows;
}
