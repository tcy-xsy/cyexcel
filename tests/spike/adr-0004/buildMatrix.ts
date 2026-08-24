import type { CellMatrix, CellRecord } from "./types";

const REGIONS = ["NA", "EU", "APAC", "LATAM", "MEA"];
const STATUSES = ["active", "pending", "archived"];

/** 与 tests/performance/support/dataset.ts 的 buildCellValue 同一分布。 */
function buildCellRecord(row: number, column: number): CellRecord {
  switch (column % 5) {
    case 0:
      return { v: `Row ${row}`, t: 1 };
    case 1:
      return { v: ((row * 7 + column * 13) % 100000) / 100, t: 2 };
    case 2:
      return { v: REGIONS[(row + column) % REGIONS.length], t: 1 };
    case 3:
      return { v: STATUSES[(row + column) % STATUSES.length], t: 1 };
    default:
      return { v: (row + column) % 3 === 0, t: 3 };
  }
}

export function fillBenchmarkMatrix(matrix: CellMatrix, rowCount: number, columnCount: number): void {
  for (let row = 0; row < rowCount; row += 1) {
    for (let column = 0; column < columnCount; column += 1) {
      matrix.set(row, column, buildCellRecord(row, column));
    }
  }
}
