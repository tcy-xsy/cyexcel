import type { CellFormatContext, RuntimeColumn } from "../cells";
import type { CellRange } from "../types";
import type { CopyMode } from "./types";
import type { CellValue } from "../sheet/types";

export interface SerializeContext {
  readonly getCellValue: (row: number, column: number) => CellValue | undefined;
  readonly getColumn: (column: number) => RuntimeColumn;
  readonly findMergeAt?: (row: number, column: number) => CellRange | undefined;
}

/**
 * `CellValue` 只有 string | number | boolean，所以这里不需要"stringify 会抛"
 * 的兜底——那条 try/catch 是给复杂值准备的，随复杂值一起没了。
 */
function structuredCellText(value: CellValue | undefined, column: RuntimeColumn): string {
  return JSON.stringify({ type: column.typeName, value });
}

/**
 * Turns a range into a `string[][]` ready for tsv.ts to join — one call
 * per cell to the column's own `format()`/`serialize()` (both are closures
 * that can't run anywhere but the main thread, see ClipboardWorkerClient),
 * never a full-range stringify. A column whose `clipboard.copy` policy is
 * `false` gets an empty cell but keeps its position in the grid, so a
 * copied range still pastes back column-aligned.
 */
export function buildCellGrid(range: CellRange, mode: CopyMode, ctx: SerializeContext): string[][] {
  const grid: string[][] = [];

  for (let row = range.startRow; row <= range.endRow; row += 1) {
    const rowCells: string[] = [];

    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const runtimeColumn = ctx.getColumn(column);
      if (runtimeColumn.clipboard.copy === false) {
        rowCells.push("");
        continue;
      }

      const merge = ctx.findMergeAt?.(row, column);
      const isCoveredCell = merge !== undefined && (merge.startRow !== row || merge.startColumn !== column);
      if (isCoveredCell) {
        rowCells.push("");
        continue;
      }

      const value = ctx.getCellValue(row, column);
      const formatContext: CellFormatContext = { row, column: runtimeColumn };

      if (mode === "display") {
        rowCells.push(runtimeColumn.format(value, formatContext));
      } else if (mode === "raw") {
        rowCells.push(
          runtimeColumn.serialize ? runtimeColumn.serialize(value, formatContext) : String(value ?? "")
        );
      } else {
        rowCells.push(structuredCellText(value, runtimeColumn));
      }
    }

    grid.push(rowCells);
  }

  return grid;
}
