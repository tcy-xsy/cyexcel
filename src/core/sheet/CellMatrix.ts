import { assertValidRowOrder } from "../layout/rowOrder";
import { ROW_PLAN_HOLE, assertValidRowPlan } from "./rowPlan";
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

  /**
   * 行轴唯一的搬运原语（ADR §3.1）。`plan[newRow]` 是该行的旧下标，
   * `ROW_PLAN_HOLE` 表示新空行；没有被提到的旧行被丢弃。
   *
   * insert / delete / move / reorder 全部经由这里，各自只负责算出 plan。
   */
  applyRowPlan(plan: Int32Array): void {
    const currentExtent = this.#maxRow === undefined ? 0 : this.#maxRow + 1;
    const extent = Math.max(currentExtent, plan.length);
    assertValidRowPlan(plan, extent);
    const next = Object.create(null) as Record<number, RowBucket>;
    let maxRow: number | undefined = undefined;
    let occupied = 0;
    for (let newRow = 0; newRow < plan.length; newRow += 1) {
      const source = plan[newRow] as number;
      if (source === ROW_PLAN_HOLE) {
        continue;
      }
      const bucket = this.#rows[source];
      if (bucket) {
        next[newRow] = bucket;
        maxRow = newRow;
        occupied += Object.keys(bucket).length;
      }
    }
    this.#rows = next;
    this.#maxRow = maxRow;
    this.#occupied = occupied;
    // 行可以被丢弃，所以最宽的那一行可能没了——列的最大值必须重算
    this.#maxColumn = this.#largestColumn();
  }

  /**
   * `order[newPosition] = oldPosition`，仓库既有约定（ADR 0003）。
   * 这是 `applyRowPlan` 的一个更严格的特例：必须是严格排列，不许有洞、
   * 不许丢行。
   */
  reorderRows(order: Int32Array): void {
    assertValidRowOrder(order, order.length);
    if (this.#maxRow !== undefined && this.#maxRow >= order.length) {
      throw new RangeError(
        `occupied row ${this.#maxRow} >= order length ${order.length}, would be silently dropped`
      );
    }
    this.applyRowPlan(order);
  }

  /**
   * 列操作按行桶逐个重键。**不要按列数扫**——绝大多数列是空的，
   * `rowCount × columnCount` 的稠密扫描正是 ADR §3.1 点名要避免的。
   */
  #remapColumns(remap: (column: number) => number | undefined): void {
    const next = Object.create(null) as Record<number, RowBucket>;
    let maxRow: number | undefined = undefined;
    let maxColumn: number | undefined = undefined;
    let occupied = 0;
    for (const key of Object.keys(this.#rows)) {
      const row = Number(key);
      const bucket = this.#rows[row];
      if (!bucket) {
        continue;
      }
      let rebuilt: RowBucket | undefined = undefined;
      for (const columnKey of Object.keys(bucket)) {
        const column = Number(columnKey);
        const target = remap(column);
        if (target === undefined) {
          continue;
        }
        if (!rebuilt) {
          rebuilt = Object.create(null) as RowBucket;
        }
        rebuilt[target] = bucket[column] as CellRecord;
        occupied += 1;
        if (maxColumn === undefined || target > maxColumn) {
          maxColumn = target;
        }
      }
      if (rebuilt) {
        next[row] = rebuilt;
        if (maxRow === undefined || row > maxRow) {
          maxRow = row;
        }
      }
    }
    this.#rows = next;
    this.#maxRow = maxRow;
    this.#maxColumn = maxColumn;
    this.#occupied = occupied;
  }

  insertColumns(start: number, count: number): void {
    this.#remapColumns((column) => (column < start ? column : column + count));
  }

  deleteColumns(start: number, count: number): void {
    this.#remapColumns((column) => {
      if (column < start) {
        return column;
      }
      if (column < start + count) {
        return undefined;
      }
      return column - count;
    });
  }

  moveColumns(start: number, count: number, target: number): void {
    const insertAt = target > start ? target - count : target;
    this.#remapColumns((column) => {
      if (column >= start && column < start + count) {
        return insertAt + (column - start);
      }
      const withoutBlock = column > start ? column - count : column;
      return withoutBlock >= insertAt ? withoutBlock + count : withoutBlock;
    });
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
