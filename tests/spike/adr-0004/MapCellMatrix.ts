import type { CellMatrix, CellRecord } from "./types";

/** 候选 A：外层 Map<行号, Map<列号, 记录>>。键是真正的数字。 */
export class MapCellMatrix implements CellMatrix {
  readonly rowCount: number;
  readonly columnCount: number;
  readonly #rows = new Map<number, Map<number, CellRecord>>();

  constructor(rowCount: number, columnCount: number) {
    this.rowCount = rowCount;
    this.columnCount = columnCount;
  }

  get(row: number, column: number): CellRecord | undefined {
    return this.#rows.get(row)?.get(column);
  }

  set(row: number, column: number, record: CellRecord): void {
    let cells = this.#rows.get(row);
    if (!cells) {
      cells = new Map<number, CellRecord>();
      this.#rows.set(row, cells);
    }
    cells.set(column, record);
  }

  delete(row: number, column: number): void {
    const cells = this.#rows.get(row);
    if (!cells) {
      return;
    }
    cells.delete(column);
    if (cells.size === 0) {
      this.#rows.delete(row);
    }
  }

  occupiedCount(): number {
    let total = 0;
    for (const cells of this.#rows.values()) {
      total += cells.size;
    }
    return total;
  }

  /**
   * 逐"行桶"搬运——外层是 Map，行的整条 Map 可以整体换位，不必逐格拷贝。
   * 这是记录矩阵相对 Univer `ObjectMatrix` 逐格深拷贝的一个潜在优势，
   * Q3 要测的正是它到底值多少。
   */
  reorderRows(order: Int32Array): void {
    const next = new Map<number, Map<number, CellRecord>>();
    for (let newRow = 0; newRow < order.length; newRow += 1) {
      const oldRow = order[newRow] as number;
      const cells = this.#rows.get(oldRow);
      if (cells) {
        next.set(newRow, cells);
      }
    }
    this.#rows.clear();
    for (const [row, cells] of next) {
      this.#rows.set(row, cells);
    }
  }
}
