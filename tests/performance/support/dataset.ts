import type { CyExcelColumnConfig } from "../../../src/core/CyExcelCore";
import type { RowData } from "../../../src/core/types";

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

export function buildBenchmarkRow(row: number, columnCount: number): RowData {
  const data: Record<string, unknown> = {};
  for (let column = 0; column < columnCount; column += 1) {
    data[`c${column}`] = buildCellValue(row, column);
  }
  return data;
}

export function buildBenchmarkRows(rowCount: number, columnCount: number): RowData[] {
  return Array.from({ length: rowCount }, (_unused, row) => buildBenchmarkRow(row, columnCount));
}
