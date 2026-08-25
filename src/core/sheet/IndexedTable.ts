import { remapByColumnShift, remapByRowPlan } from "./indexShift";
import type { ColumnShift } from "./columnShift";
import { CyExcelError, CyExcelErrorCode } from "../errors";

export type IndexAxis = "row" | "column";

/**
 * 一张按行或按列下标索引的稀疏侧表。
 *
 * **为什么侧���要自带搬运。** ADR §9 不变量 4 要求每张按下标索引的侧表为 7 个
 * 结构操作提供重键。阶段二的教训是这条纪律靠人记不住：行元数据在 insert /
 * delete / move / reorder 上各漏了一次，四次都是**静默丢数据**。把存储和重键
 * 绑在同一个对象上，"忘了搬这张表"就不再是一个可能的状态——登记处（见
 * `IndexedTableRegistry`）遍历时它自己会搬。
 *
 * `T` 只受存储约束，本类不解释它。
 */
export class IndexedTable<T> {
  #data: Record<number, T>;
  #extent: number;

  constructor(
    readonly name: string,
    readonly indexedBy: IndexAxis,
    initial?: Readonly<Record<number, T>>
  ) {
    // 拷一层：调用方交进来的对象之后不该被我们改动
    this.#data = initial ? { ...initial } : {};
    this.#extent = 0;
    for (const key of Object.keys(this.#data)) {
      const index = Number(key) + 1;
      if (index > this.#extent) {
        this.#extent = index;
      }
    }
  }

  read(): Readonly<Record<number, T>> {
    return this.#data;
  }

  get(index: number): T | undefined {
    return this.#data[index];
  }

  set(index: number, value: T): void {
    this.#data[index] = value;
    if (index + 1 > this.#extent) {
      this.#extent = index + 1;
    }
  }

  delete(index: number): void {
    if (!(index in this.#data)) {
      return;
    }
    delete this.#data[index];
    // 删掉的正好是最大的那个才需要重算，与 CellMatrix 同样的取舍
    if (index + 1 === this.#extent) {
      this.#extent = this.#recomputeExtent();
    }
  }

  /** 最大占用下标加一；空表为 0。不是键的个数。 */
  extent(): number {
    return this.#extent;
  }

  applyRowPlan(plan: Int32Array): void {
    this.#assertAxis("row");
    this.#data = remapByRowPlan(this.#data, plan);
    this.#extent = this.#recomputeExtent();
  }

  applyColumnShift(shift: ColumnShift): void {
    this.#assertAxis("column");
    this.#data = remapByColumnShift(this.#data, shift);
    this.#extent = this.#recomputeExtent();
  }

  #assertAxis(expected: IndexAxis): void {
    if (this.indexedBy !== expected) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidOptions,
        `table "${this.name}" is indexed by ${this.indexedBy}, cannot take a ${expected} rekey`
      );
    }
  }

  #recomputeExtent(): number {
    let extent = 0;
    for (const key of Object.keys(this.#data)) {
      const index = Number(key) + 1;
      if (index > extent) {
        extent = index;
      }
    }
    return extent;
  }
}
