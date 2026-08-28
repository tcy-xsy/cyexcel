import { assertValidRowOrder } from "../layout/rowOrder";
import {
  buildDeleteColumnsShift,
  buildInsertColumnsShift,
  buildMoveColumnsShift,
  type ColumnShift
} from "./columnShift";
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
  /**
   * 与另一个矩阵按引用共享着行桶的那些**行号**——写它们之前必须先私有化一份。
   * 空集（`undefined`）是常态：没克隆过的矩阵一个共享桶都没有，那条路上不多一次
   * 判断以外的任何开销。
   */
  #shared: Set<number> | undefined = undefined;

  /**
   * 克隆：外层映射复制一层，**行桶按引用共享**，两边都记下"这些行是共享的"，谁先
   * 写谁先私有化。
   *
   * **为什么共享桶是安全的、而且是这个类既有的做法。** `applyRowPlan` 早就在按引用
   * 搬桶（`next[newRow] = bucket`），`#remapColumns` 则一律重建新桶。所以全类只有
   * `set` / `delete` 两处会**就地改**一个桶，写时复制的守卫只要加在那两处。
   *
   * **代价是 O(占用行数)，不是 O(格子数)**：复制外层键、外加两个同样大小的 `Set`。
   * 这正是暂存层选"克隆影子模型"而不是"覆盖层 + 累积下标重映射"的前提——后者要把
   * 7 个结构操作的语义在覆盖层里再实现一遍，而这条分支已经三次栽在"同一套语义两份
   * 实现"上。
   *
   * **标记必须两边都打。** 共享是对称的：克隆之后原矩阵手里的桶同样是共享的，只标
   * 克隆那一边，往**原**矩阵上写就会改到影子看得见的桶。代价是双方各自的第一次写
   * 都会各拷一份（一共两份，其中一份"多余"）——这个多余换掉的是一整套引用计数。
   *
   * **值（`CellRecord`）不拷贝**，按 `types.ts` 顶部那条浅共享规则共享引用。
   */
  clone(): CellMatrix {
    const copy = new CellMatrix();
    copy.#rows = Object.assign(Object.create(null) as Record<number, RowBucket>, this.#rows);
    copy.#occupied = this.#occupied;
    copy.#maxRow = this.#maxRow;
    copy.#maxColumn = this.#maxColumn;
    const shared = new Set<number>();
    for (const key of Object.keys(this.#rows)) {
      shared.add(Number(key));
    }
    // 两个独立的 `Set`：双方各自私有化各自的桶，标记随之分叉。这里整个替换掉原来
    // 那个而不是并进去——新集合覆盖了当前每一行，是旧集合（去掉已消失的行之后）
    // 的超集。
    copy.#shared = shared;
    this.#shared = new Set(shared);
    return copy;
  }

  /**
   * 要就地改 `row` 这一桶之前调用。它若是共享的，先换成一份私有拷贝。
   *
   * 桶里的 `CellRecord` 仍按引用共享——拷的是容器，不是格子。
   */
  #privatize(row: number, bucket: RowBucket): RowBucket {
    if (this.#shared === undefined || !this.#shared.has(row)) {
      return bucket;
    }
    const copy = Object.create(null) as RowBucket;
    for (const key of Object.keys(bucket)) {
      copy[Number(key)] = bucket[Number(key)] as CellRecord;
    }
    this.#rows[row] = copy;
    this.#shared.delete(row);
    return copy;
  }

  get(row: number, column: number): CellRecord | undefined {
    return this.#rows[row]?.[column];
  }

  set(row: number, column: number, record: CellRecord): void {
    let bucket = this.#rows[row];
    if (!bucket) {
      bucket = Object.create(null) as RowBucket;
      this.#rows[row] = bucket;
    } else {
      bucket = this.#privatize(row, bucket);
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
    const existing = this.#rows[row];
    if (!existing || existing[column] === undefined) {
      return;
    }
    const bucket = this.#privatize(row, existing);
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
   *
   * `sourceExtent` is optional and used when the source data has knowledge of rows
   * beyond what the matrix contains (e.g., rows with metadata but no cells).
   * Defaults to the matrix's own extent if not provided.
   */
  applyRowPlan(plan: Int32Array, sourceExtent?: number): void {
    const currentExtent = this.#maxRow === undefined ? 0 : this.#maxRow + 1;
    const extent = Math.max(currentExtent, plan.length, sourceExtent ?? 0);
    assertValidRowPlan(plan, extent);
    const next = Object.create(null) as Record<number, RowBucket>;
    // 共享标记是**按行号**记的，而这里正是在换行号：不跟着搬，一个共享桶换个位置
    // 就成了"没标记"，下一次写就直接改在影子看得见的那个桶上。这条不跟着搬的话，
    // 现象是撤销之后影子里凭空多出几个格子——查起来会一路查到暂存层去。
    const nextShared = this.#shared === undefined ? undefined : new Set<number>();
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
        if (this.#shared?.has(source)) {
          nextShared?.add(newRow);
        }
        maxRow = newRow;
        occupied += Object.keys(bucket).length;
      }
    }
    this.#rows = next;
    this.#shared = nextShared;
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
  #remapColumns(remap: ColumnShift): void {
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
    // 每一个留下来的桶都是上面新建的 `rebuilt`，与任何别的矩阵都不共享了
    this.#shared = undefined;
    this.#maxRow = maxRow;
    this.#maxColumn = maxColumn;
    this.#occupied = occupied;
  }

  insertColumns(start: number, count: number): void {
    this.#remapColumns(buildInsertColumnsShift(start, count));
  }

  deleteColumns(start: number, count: number): void {
    this.#remapColumns(buildDeleteColumnsShift(start, count));
  }

  moveColumns(start: number, count: number, target: number): void {
    this.#remapColumns(buildMoveColumnsShift(start, count, target));
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
