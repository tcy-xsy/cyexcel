import { remapByColumnShift, remapByRowPlan } from "./indexShift";
import type { ColumnShift } from "./columnShift";
import { assertCanonicalIndexKeys, assertIndexIsNonNegativeInteger } from "./indexKey";
import { CyExcelError, CyExcelErrorCode } from "../errors";

export type IndexAxis = "row" | "column";

/**
 * 一张按行或按列下标索引的稀疏侧表。
 *
 * **为什么侧表要自带搬运。** ADR §9 不变量 4 要求每张按下标索引的侧表为 7 个
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

  /**
   * `initial` 的键在**展开之前**就要校验。
   *
   * 拥有存储的类才是该拒绝它搬不动的键的类：一个落在非规范键上的条目，`extent()`
   * 算不进去（`Number("-1") + 1 === 0`），于是排列永远不会点名它，第一次结构操作
   * 就把它静默丢掉——这正是 ADR §9 不变量 4 存在的那一类缺陷。校验此前只活在
   * `WorkbookModel.addResource` 那条路上，而本类的构造函数是公开导出的，
   * `SheetModel.registerTable()` 又照单全收，等于登记处自己重新生产了它要根除的缺陷。
   */
  constructor(
    readonly name: string,
    readonly indexedBy: IndexAxis,
    initial?: Readonly<Record<number, T>>
  ) {
    if (initial) {
      assertCanonicalIndexKeys(initial, `table "${name}" initial data`);
    }
    // 拷一层：调用方交进来的对象之后不该被我们改动
    this.#data = initial ? { ...initial } : {};
    this.#extent = this.#recomputeExtent();
  }

  /**
   * **返回的是活对象，不是快照。** 之后任何 `set` / `delete` / 重键都会就地改在
   * 它上面——拿着它跨过一次写入，再读就是新值。
   *
   * 不在这里拷一份是有意的：`serializeSheet` 与 `tableToResource` 各自都会展开
   * 一层，在这里再拷等于每条快照路径拷两遍，而少拷这一遍正是这个模型要省的开销。
   * 需要一份稳定视图的调用方自己 `{ ...table.read() }`。
   *
   * 返回类型上的 `Readonly` 挡的是"往里写"，挡不住"它自己会变"——这是两件事，
   * 而 `SheetModel.registerTable()` 把这个句柄交到了插件作者手里（spec §7.2），
   * 所以这条必须写在这里而不是靠人猜。
   */
  read(): Readonly<Record<number, T>> {
    return this.#data;
  }

  /**
   * 克隆：名字与轴照抄，条目**拷一层**（构造函数本来就拷，见上）。
   *
   * 这里没有 `CellMatrix` 那样的写时复制，是因为量级不同：侧表只在被自定义过的
   * 行/列上有条目，条目数远小于格子数，拷一层就够便宜；给它加一套 COW 反而是拿
   * 复杂度换一个量不出来的收益。
   *
   * **值按引用共享**（`types.ts` 顶部那条浅共享规则），插件载荷不深拷。
   */
  clone(): IndexedTable<T> {
    return new IndexedTable<T>(this.name, this.indexedBy, this.#data);
  }

  get(index: number): T | undefined {
    return this.#data[index];
  }

  /**
   * 下标必须是非负整数，理由同构造函数：负数或小数下标搬不动，会在第一次结构
   * 操作里静默消失。校验放在这里而不是各个调用方，是因为拥有存储的类才是该拒绝
   * 它搬不动的键的类——`SheetModel.setRowMeta` 之类的门面都委托到这里。
   */
  set(index: number, value: T): void {
    assertIndexIsNonNegativeInteger(index, `table "${this.name}" index`);
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
