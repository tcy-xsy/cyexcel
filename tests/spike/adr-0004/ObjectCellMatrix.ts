import type { CellMatrix, CellRecord } from "./types";

type RowBucket = Record<number, CellRecord>;

/** 候选 B：{[行号]: {[列号]: 记录}}，即 Univer `cellData` 的实际形态。 */
export class ObjectCellMatrix implements CellMatrix {
  readonly rowCount: number;
  readonly columnCount: number;
  #rows: Record<number, RowBucket> = Object.create(null) as Record<number, RowBucket>;

  constructor(rowCount: number, columnCount: number) {
    this.rowCount = rowCount;
    this.columnCount = columnCount;
  }

  get(row: number, column: number): CellRecord | undefined {
    return this.#rows[row]?.[column];
  }

  set(row: number, column: number, record: CellRecord): void {
    let bucket = this.#rows[row];
    if (!bucket) {
      bucket = Object.create(null) as RowBucket;
      this.#rows[row] = bucket;
    }
    bucket[column] = record;
  }

  delete(row: number, column: number): void {
    const bucket = this.#rows[row];
    if (!bucket) {
      return;
    }
    delete bucket[column];
    if (Object.keys(bucket).length === 0) {
      delete this.#rows[row];
    }
  }

  occupiedCount(): number {
    let total = 0;
    for (const key of Object.keys(this.#rows)) {
      total += Object.keys(this.#rows[Number(key)] as RowBucket).length;
    }
    return total;
  }

  reorderRows(order: Int32Array): void {
    const next = Object.create(null) as Record<number, RowBucket>;
    for (let newRow = 0; newRow < order.length; newRow += 1) {
      const bucket = this.#rows[order[newRow] as number];
      if (bucket) {
        next[newRow] = bucket;
      }
    }
    this.#rows = next;
  }
}
