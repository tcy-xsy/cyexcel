import type { CellFormatContext, RuntimeColumn } from "../cells";
import type { CellRange } from "../types";
import type { CopyMode } from "./types";

export interface SerializeContext {
  readonly getCellValue: (row: number, column: number) => unknown;
  readonly getColumn: (column: number) => RuntimeColumn;
  readonly findMergeAt?: (row: number, column: number) => CellRange | undefined;
}

/**
 * Same "degrade instead of throw" precedent as the built-in array/json
 * `parse()` (see doc/progress.md's editor-pluggability patch) — a
 * structured-mode copy shouldn't blow up on a value JSON.stringify can't
 * handle (e.g. one containing a BigInt).
 */
function structuredCellText(value: unknown, column: RuntimeColumn): string {
  try {
    const serialized = JSON.stringify({ type: column.typeName, value });
    return serialized ?? "";
  } catch {
    return "";
  }
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
