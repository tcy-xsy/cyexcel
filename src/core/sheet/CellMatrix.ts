import { assertValidRowOrder } from "../layout/rowOrder";
import type { CellRecord } from "./types";

type RowBucket = Record<number, CellRecord>;

/**
 * 稀疏记录矩阵：外层按行号索引，内层是该行的格子桶。
 *
 * 用嵌套普通对象而非嵌套 `Map`——前置 Spike 实测嵌套对象 35.8 B/占用格、
 * `Map` 46.2，且读路径快 2.4–2.6 倍（`doc/adr-0004-spike-findings.md`）。
 * 行桶用 `Object.create(null)`，避免数字键与原型成员冲突。
 *
 * `maxOccupiedRow` / `maxOccupiedColumn` 增量维护：它们进入了每一次坐标计算
 * （spec §3.3 后果 ②），不能每次扫键。代价放在删除上——**只有删掉的正好是当前
 * 最大值时才扫一次**回退，平时 O(1)。删除低频、坐标查询高频。
 */
export class CellMatrix {
  #rows: Record<number, RowBucket> = Object.create(null) as Record<number, RowBucket>;
  #occupied = 0;
  #maxRow: number | undefined = undefined;
  #maxColumn: number | undefined = undefined;

  get(row: number, column: number): CellRecord | undefined {
    return this.#rows[row]?.[column];
  }

  set(row: number, column: number, record: CellRecord): void {
    let bucket = this.#rows[row];
    if (!bucket) {
      bucket = Object.create(null) as RowBucket;
      this.#rows[row] = bucket;
    }
    if (bucket[column] === undefined) {
      this.#occupied += 1;
    }
    bucket[column] = record;
    if (this.#maxRow === undefined || row > this.#maxRow) {
      this.#maxRow = row;
    }
    if (this.#maxColumn === undefined || column > this.#maxColumn) {
      this.#maxColumn = column;
    }
  }

  delete(row: number, column: number): void {
    const bucket = this.#rows[row];
    if (!bucket || bucket[column] === undefined) {
      return;
    }
    delete bucket[column];
    this.#occupied -= 1;
    const bucketEmpty = Object.keys(bucket).length === 0;
    if (bucketEmpty) {
      delete this.#rows[row];
    }
    // 只有删掉的正好是当前最大值时才需要回退扫描
    if (bucketEmpty && row === this.#maxRow) {
      this.#maxRow = this.#largestRow();
    }
    if (column === this.#maxColumn) {
      this.#maxColumn = this.#largestColumn();
    }
  }

  occupiedCount(): number {
    return this.#occupied;
  }

  maxOccupiedRow(): number | undefined {
    return this.#maxRow;
  }

  maxOccupiedColumn(): number | undefined {
    return this.#maxColumn;
  }

  /** 遍历被占用的行及其桶。列操作需要它——否则会退化成稠密扫描（ADR §3.1）。 */
  *rows(): Generator<[number, Readonly<RowBucket>]> {
    for (const key of Object.keys(this.#rows)) {
      const row = Number(key);
      const bucket = this.#rows[row];
      if (bucket) {
        yield [row, bucket];
      }
    }
  }

  /** `order[newPosition] = oldPosition`，仓库既有约定（ADR 0003）。 */
  reorderRows(order: Int32Array): void {
    assertValidRowOrder(order, order.length);
    if (this.#maxRow !== undefined && this.#maxRow >= order.length) {
      throw new RangeError(
        `occupied row ${this.#maxRow} >= order length ${order.length}, would be silently dropped`
      );
    }
    const next = Object.create(null) as Record<number, RowBucket>;
    let maxRow: number | undefined = undefined;
    for (let newRow = 0; newRow < order.length; newRow += 1) {
      const bucket = this.#rows[order[newRow] as number];
      if (bucket) {
        next[newRow] = bucket;
        maxRow = newRow;
      }
    }
    this.#rows = next;
    this.#maxRow = maxRow;
    // 列不受行重排影响，#maxColumn 与 #occupied 保持不变
  }

  #largestRow(): number | undefined {
    let largest: number | undefined = undefined;
    for (const key of Object.keys(this.#rows)) {
      const row = Number(key);
      if (largest === undefined || row > largest) {
        largest = row;
      }
    }
    return largest;
  }

  #largestColumn(): number | undefined {
    let largest: number | undefined = undefined;
    for (const rowKey of Object.keys(this.#rows)) {
      const bucket = this.#rows[Number(rowKey)] as RowBucket;
      for (const columnKey of Object.keys(bucket)) {
        const column = Number(columnKey);
        if (largest === undefined || column > largest) {
          largest = column;
        }
      }
    }
    return largest;
  }
}
