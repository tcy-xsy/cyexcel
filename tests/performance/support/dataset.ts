import type { CyExcelColumnConfig } from "../../../src/core/CyExcelCore";
import { rowsToWorkbookSnapshot } from "../../../src/core/sheet";
import type { ImportRow, WorkbookSnapshot } from "../../../src/core/sheet";

const REGIONS = ["NA", "EU", "APAC", "LATAM", "MEA"];
const STATUSES = ["active", "pending", "archived"];

/**
 * Deterministic synthetic columns/rows for the Step 15 benchmark fixtures.
 * Values are computed from (row, column) only — no shared mutable pool
 * state — so generation stays cheap even at 1,000,000 rows, matching the
 * "100 万行 x 50 列" benchmark scale (doc/implementation-steps.md's Step 15
 * table). Modeled on playground/playground.js's buildRow()/buildRows(),
 * which already exercises 1,000,000 rows via the playground's own row-count
 * selector, but generalized to an arbitrary column count instead of a fixed
 * hand-written column list.
 */
export function buildBenchmarkColumns(columnCount: number): CyExcelColumnConfig[] {
  return Array.from({ length: columnCount }, (_unused, column) => ({
    id: `c${column}`,
    field: `c${column}`,
    width: 100
  }));
}

function buildCellValue(row: number, column: number): unknown {
  switch (column % 5) {
    case 0:
      return `Row ${row}`;
    case 1:
      return ((row * 7 + column * 13) % 100000) / 100;
    case 2:
      return REGIONS[(row + column) % REGIONS.length];
    case 3:
      return STATUSES[(row + column) % STATUSES.length];
    default:
      return (row + column) % 3 === 0;
  }
}

export function buildBenchmarkRow(row: number, columnCount: number): ImportRow {
  const data: Record<string, unknown> = {};
  for (let column = 0; column < columnCount; column += 1) {
    data[`c${column}`] = buildCellValue(row, column);
  }
  return data;
}

export function buildBenchmarkRows(rowCount: number, columnCount: number): ImportRow[] {
  return Array.from({ length: rowCount }, (_unused, row) => buildBenchmarkRow(row, columnCount));
}

/** 第 i 个字段就是第 i 列——导入的 `fields` 参数（ADR 0004 §3.2、§4.5）。 */
export function buildBenchmarkFields(columnCount: number): string[] {
  return Array.from({ length: columnCount }, (_unused, column) => `c${column}`);
}

/**
 * 基准夹具的数据入口：对象数组一次性导进稀疏矩阵（`CyExcelCoreOptions.snapshot`）。
 *
 * 中间那个行对象数组**不留引用**，交给这里就是为了让它在导完之后立刻可回收——
 * 100 万行 × 50 列的基准里它本身就是几百 MB，与矩阵同时活着毫无必要。夹具因此
 * 应当调这个函数，而不是自己 `buildBenchmarkRows()` 再 `rowsToWorkbookSnapshot()`
 * ——后者会把数组钉在夹具的作用域里直到测量结束。
 */
export function buildBenchmarkSnapshot(rowCount: number, columnCount: number): WorkbookSnapshot {
  return rowsToWorkbookSnapshot(
    buildBenchmarkRows(rowCount, columnCount),
    buildBenchmarkFields(columnCount)
  );
}
