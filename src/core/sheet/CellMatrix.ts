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

/** Sparse cells with a persistent, paged row index and copy-on-write row buckets. */
export class CellMatrix {
  #rows = new RowIndex();
  #owned = new WeakSet<RowBucket>();
  #occupied = 0;
  #formulaCount = 0;
  #maxRow: number | undefined = undefined;
  #maxColumn: number | undefined = undefined;
  #rowDirty = false;
  #columnDirty = false;

  /** Snapshots share index pages and buckets; subsequent writes detach only their path. */
  clone(): CellMatrix {
    const copy = new CellMatrix();
    copy.#rows = this.#rows.clone();
    copy.#occupied = this.#occupied;
    copy.#formulaCount = this.#formulaCount;
    copy.#maxRow = this.#maxRow;
    copy.#maxColumn = this.#maxColumn;
    copy.#rowDirty = this.#rowDirty;
    copy.#columnDirty = this.#columnDirty;
    this.#owned = new WeakSet();
    return copy;
  }

  #privatize(row: number, bucket: RowBucket): RowBucket {
    if (this.#owned.has(bucket)) return bucket;
    const copy = Object.assign(Object.create(null) as RowBucket, bucket);
    this.#rows.set(row, copy);
    this.#owned.add(copy);
    return copy;
  }

  get(row: number, column: number): CellRecord | undefined {
    return this.#rows.get(row)?.[column];
  }

  set(row: number, column: number, record: CellRecord): void {
    const previous = this.get(row, column);
    this.#formulaCount += Number(record.f !== undefined) - Number(previous?.f !== undefined);
    let bucket = this.#rows.get(row);
    if (!bucket) {
      bucket = Object.create(null) as RowBucket;
      this.#rows.set(row, bucket);
      this.#owned.add(bucket);
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
    const existing = this.#rows.get(row);
    if (!existing || existing[column] === undefined) {
      return;
    }
    const bucket = this.#privatize(row, existing);
    if (existing[column]?.f !== undefined) this.#formulaCount -= 1;
    delete bucket[column];
    this.#occupied -= 1;
    const bucketEmpty = Object.keys(bucket).length === 0;
    if (bucketEmpty) {
      this.#rows.delete(row);
    }
    // Defer boundary scans until queried, so batch clears scan at most once.
    if (bucketEmpty && row === this.#maxRow) {
      this.#rowDirty = true;
    }
    if (column === this.#maxColumn) {
      this.#columnDirty = true;
    }
  }

  occupiedCount(): number {
    return this.#occupied;
  }

  formulaCount(): number {
    return this.#formulaCount;
  }

  maxOccupiedRow(): number | undefined {
    if (this.#rowDirty) {
      this.#maxRow = this.#largestRow();
      this.#rowDirty = false;
    }
    return this.#maxRow;
  }

  maxOccupiedColumn(): number | undefined {
    if (this.#columnDirty) {
      this.#maxColumn = this.#largestColumn();
      this.#columnDirty = false;
    }
    return this.#maxColumn;
  }

  /** 遍历被占用的行及其桶。列操作需要它——否则会退化成稠密扫描（ADR §3.1）。 */
  *rows(): Generator<[number, Readonly<RowBucket>]> {
    yield* this.#rows.entries();
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
    const currentMax = this.maxOccupiedRow();
    const currentExtent = currentMax === undefined ? 0 : currentMax + 1;
    const extent = Math.max(currentExtent, plan.length, sourceExtent ?? 0);
    assertValidRowPlan(plan, extent);
    const next = new RowIndex();
    let maxRow: number | undefined = undefined;
    let occupied = 0;
    let formulas = 0;
    for (let newRow = 0; newRow < plan.length; newRow += 1) {
      const source = plan[newRow] as number;
      if (source === ROW_PLAN_HOLE) {
        continue;
      }
      const bucket = this.#rows.get(source);
      if (bucket) {
        next.set(newRow, bucket);
        maxRow = newRow;
        occupied += Object.keys(bucket).length;
        for (const record of Object.values(bucket)) if (record.f !== undefined) formulas += 1;
      }
    }
    this.#rows = next;
    this.#rowDirty = false;
    this.#maxRow = maxRow;
    this.#occupied = occupied;
    this.#formulaCount = formulas;
    // 行可以被丢弃，所以最宽的那一行可能没了——列的最大值必须重算
    this.#columnDirty = true;
  }

  /**
   * `order[newPosition] = oldPosition`，仓库既有约定（ADR 0003）。
   * 这是 `applyRowPlan` 的一个更严格的特例：必须是严格排列，不许有洞、
   * 不许丢行。
   */
  reorderRows(order: Int32Array): void {
    assertValidRowOrder(order, order.length);
    this.maxOccupiedRow();
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
    const next = new RowIndex();
    let maxRow: number | undefined = undefined;
    let maxColumn: number | undefined = undefined;
    let occupied = 0;
    let formulas = 0;
    for (const [row, bucket] of this.#rows.entries()) {
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
        if (bucket[column]?.f !== undefined) formulas += 1;
        if (maxColumn === undefined || target > maxColumn) {
          maxColumn = target;
        }
      }
      if (rebuilt) {
        next.set(row, rebuilt);
        this.#owned.add(rebuilt);
        if (maxRow === undefined || row > maxRow) {
          maxRow = row;
        }
      }
    }
    this.#rows = next;
    this.#rowDirty = false;
    this.#columnDirty = false;
    this.#maxRow = maxRow;
    this.#maxColumn = maxColumn;
    this.#occupied = occupied;
    this.#formulaCount = formulas;
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
    for (const [row] of this.#rows.entries()) {
      if (largest === undefined || row > largest) {
        largest = row;
      }
    }
    return largest;
  }

  #largestColumn(): number | undefined {
    let largest: number | undefined = undefined;
    for (const [, bucket] of this.#rows.entries()) {
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

const PAGE_SIZE = 1024;
type IndexNode = { owner: object; entries: Map<number, IndexNode | RowBucket> };

/** Pages bound the amount copied per edit, independently of the occupied row count. */
class RowIndex {
  #owner: object = {};
  #root: IndexNode = { owner: this.#owner, entries: new Map() };
  #level = 0;

  clone(): RowIndex {
    const copy = new RowIndex();
    copy.#root = this.#root;
    copy.#level = this.#level;
    this.#owner = {};
    return copy;
  }

  get(row: number): RowBucket | undefined {
    if (row >= PAGE_SIZE ** (this.#level + 1)) return undefined;
    let node = this.#root;
    for (let level = this.#level; level > 0; level -= 1) {
      const child = node.entries.get(Math.floor(row / PAGE_SIZE ** level) % PAGE_SIZE);
      if (!child) return undefined;
      node = child as IndexNode;
    }
    return node.entries.get(row % PAGE_SIZE) as RowBucket | undefined;
  }

  #detach(node: IndexNode): IndexNode {
    return node.owner === this.#owner
      ? node
      : { owner: this.#owner, entries: new Map(node.entries) };
  }

  set(row: number, bucket: RowBucket): void {
    while (row >= PAGE_SIZE ** (this.#level + 1)) {
      this.#root = { owner: this.#owner, entries: new Map([[0, this.#root]]) };
      this.#level += 1;
    }
    this.#root = this.#detach(this.#root);
    let node = this.#root;
    for (let level = this.#level; level > 0; level -= 1) {
      const key = Math.floor(row / PAGE_SIZE ** level) % PAGE_SIZE;
      const old = node.entries.get(key) as IndexNode | undefined;
      const child = old ? this.#detach(old) : { owner: this.#owner, entries: new Map() };
      node.entries.set(key, child);
      node = child;
    }
    node.entries.set(row % PAGE_SIZE, bucket);
  }

  delete(row: number): void {
    this.#root = this.#detach(this.#root);
    let node = this.#root;
    const path: [IndexNode, number][] = [];
    for (let level = this.#level; level > 0; level -= 1) {
      const key = Math.floor(row / PAGE_SIZE ** level) % PAGE_SIZE;
      const old = node.entries.get(key) as IndexNode | undefined;
      if (!old) return;
      const child = this.#detach(old);
      node.entries.set(key, child);
      path.push([node, key]);
      node = child;
    }
    node.entries.delete(row % PAGE_SIZE);
    for (let index = path.length - 1; index >= 0 && node.entries.size === 0; index -= 1) {
      const [parent, key] = path[index] as [IndexNode, number];
      parent.entries.delete(key);
      node = parent;
    }
  }

  *entries(): Generator<[number, RowBucket]> {
    yield* this.#walk(this.#root, this.#level, 0);
  }

  *#walk(node: IndexNode, level: number, offset: number): Generator<[number, RowBucket]> {
    for (const key of Array.from(node.entries.keys()).sort((a, b) => a - b)) {
      const value = node.entries.get(key) as IndexNode | RowBucket;
      if (!value) continue;
      const row = offset + key * PAGE_SIZE ** level;
      if (level === 0) yield [row, value as RowBucket];
      else yield* this.#walk(value as IndexNode, level - 1, row);
    }
  }
}
