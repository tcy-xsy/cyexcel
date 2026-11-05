import type { RuntimeColumn } from "../cells";
import type { CellRange } from "../types";
import type { CopyPlan } from "./types";

// Both estimateCellBytes() and the plan() loop below only ever look at a
// bounded prefix of an array/object-valued cell — full-length inspection
// (or worse, JSON.stringify) would turn "plan a copy" into the same long
// task it's meant to let the caller avoid (see the 100k-item single-cell
// no-long-task exit criterion).
const SAMPLE_LIMIT = 1000;

export interface CopyPlannerOptions {
  readonly getCellValue: (row: number, column: number) => unknown;
  readonly getColumn: (column: number) => Pick<RuntimeColumn, "clipboard">;
  readonly maxCells: number;
  readonly maxBytes: number;
  readonly findMergeAt?: (row: number, column: number) => CellRange | undefined;
}

function estimateCellBytes(value: unknown): number {
  if (value === null || value === undefined) {
    return 0;
  }
  if (typeof value === "string") {
    return value.length;
  }
  if (typeof value === "number" || typeof value === "boolean") {
    return 8;
  }
  if (value instanceof Date) {
    return 24;
  }
  if (Array.isArray(value)) {
    return Math.min(value.length, SAMPLE_LIMIT) * 8 + 2;
  }
  if (typeof value === "object") {
    return Math.min(Object.keys(value).length, SAMPLE_LIMIT) * 16 + 2;
  }
  return String(value).length;
}

function isComplexValue(value: unknown): boolean {
  return Array.isArray(value) || (typeof value === "object" && value !== null && !(value instanceof Date));
}

/**
 * Builds a CopyPlan without constructing any full serialized string —
 * `doc/architecture.md` §9's explicit constraint. Columns whose
 * `clipboard.copy` policy is `false` are skipped entirely (their cells
 * never read, matching how serializers.ts leaves them blank in the
 * output).
 */
export class CopyPlanner {
  readonly #getCellValue: CopyPlannerOptions["getCellValue"];
  readonly #getColumn: CopyPlannerOptions["getColumn"];
  readonly #maxCells: number;
  readonly #maxBytes: number;
  readonly #findMergeAt: CopyPlannerOptions["findMergeAt"];

  constructor(options: CopyPlannerOptions) {
    this.#getCellValue = options.getCellValue;
    this.#getColumn = options.getColumn;
    this.#maxCells = options.maxCells;
    this.#maxBytes = options.maxBytes;
    this.#findMergeAt = options.findMergeAt;
  }

  plan(range: CellRange): CopyPlan {
    const rows = range.endRow - range.startRow + 1;
    const columns = range.endColumn - range.startColumn + 1;
    const cells = rows * columns;

    if (cells > this.#maxCells) {
      return { range, rows, columns, cells, estimatedBytes: 0, containsComplexValues: false, recommendedMode: "download" };
    }

    let estimatedBytes = 0;
    let containsComplexValues = false;

    scan: for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let column = range.startColumn; column <= range.endColumn; column += 1) {
        if (this.#getColumn(column).clipboard.copy === false) {
          continue;
        }
        const merge = this.#findMergeAt?.(row, column);
        if (merge && (merge.startRow !== row || merge.startColumn !== column)) {
          continue;
        }
        const value = this.#getCellValue(row, column);
        estimatedBytes += estimateCellBytes(value);
        if (!containsComplexValues && isComplexValue(value)) {
          containsComplexValues = true;
        }
        if (estimatedBytes > this.#maxBytes) {
          break scan;
        }
      }
    }

    const oversized = cells > this.#maxCells || estimatedBytes > this.#maxBytes;
    const recommendedMode: CopyPlan["recommendedMode"] = oversized
      ? "download"
      : containsComplexValues
        ? "raw"
        : "display";

    return { range, rows, columns, cells, estimatedBytes, containsComplexValues, recommendedMode };
  }
}
