import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { ColumnShift } from "./columnShift";
import type { IndexedTable } from "./IndexedTable";

/**
 * 一张表的全部按下标索引的侧表。
 *
 * **它存在的理由是让 ADR §9 不变量 4 变成构造上的事实。** 那条不变量要求每张
 * 侧表为 7 个结构操作提供重键；阶段二里这条纪律被漏了四次，每次都是静默丢数据。
 * 登记处把"有哪些表"和"怎么搬"收到一处：结构操作遍历登记处，新表登记即被搬，
 * 没有"忘了加进去"这个状态可言。
 *
 * `rowData` / `columnData` 不是特例，是这里的头两个成员；spec §7.2 的插件资源
 * 走同一条路。
 */
export class IndexedTableRegistry {
  readonly #tables = new Map<string, IndexedTable<unknown>>();

  register<T>(table: IndexedTable<T>): IndexedTable<T> {
    if (this.#tables.has(table.name)) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidOptions,
        `an index-addressed table named "${table.name}" is already registered`
      );
    }
    this.#tables.set(table.name, table as IndexedTable<unknown>);
    return table;
  }

  has(name: string): boolean {
    return this.#tables.has(name);
  }

  table(name: string): IndexedTable<unknown> | undefined {
    return this.#tables.get(name);
  }

  names(): readonly string[] {
    return [...this.#tables.keys()];
  }

  /**
   * 所有**行**表的占用范围并集。
   *
   * 必须是并集：一张表在第 9 行有数据、另一张只到第 2 行，按 3 去算排列会把
   * 第 9 行那条静默丢掉。这正是阶段二 `#occupiedRowExtent()` 修过的那个洞，
   * 区别是现在它对任意多张表成立，而不是对硬编码的两张。
   */
  rowExtent(): number {
    let extent = 0;
    for (const table of this.#tables.values()) {
      if (table.indexedBy === "row" && table.extent() > extent) {
        extent = table.extent();
      }
    }
    return extent;
  }

  applyRowPlan(plan: Int32Array): void {
    for (const table of this.#tables.values()) {
      if (table.indexedBy === "row") {
        table.applyRowPlan(plan);
      }
    }
  }

  applyColumnShift(shift: ColumnShift): void {
    for (const table of this.#tables.values()) {
      if (table.indexedBy === "column") {
        table.applyColumnShift(shift);
      }
    }
  }
}
