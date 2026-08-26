import type { CellRange } from "../types";
import type { CellRecord, ColumnMeta, RowMeta } from "./types";

/** 一行的格子：列号 -> 记录。稀疏，未占用的列不出现（spec §11 不变量 1）。 */
export type RowCells = Readonly<Record<number, CellRecord>>;

/** 一列的格子：**行号** -> 记录。稀疏，未占用的行不出现。 */
export type ColumnCells = Readonly<Record<number, CellRecord>>;

/**
 * 被删除的那几行/列在**每一张插件侧表**里的条目，按**表名**索引。
 *
 * 每个数组与被恢复的行数（`restore-rows` 的 `rows.length`）或列数
 * （`restore-columns` 的 `count`）等长、**逐位置对应**：第 `i` 项就是
 * `start + i` 那一行/列的条目，`undefined` 表示那一行/列在这张表里本来就没有
 * 条目（稀疏表里"有没有这个键"就是"有没有这条条目"本身）。
 *
 * **为什么必须有这个字段。** spec §11 不变量 8：声明 `indexedBy: "row" | "column"`
 * 的资源，其重键**与逆操作**由核心提供并纳入同一 transaction（§7.2）。删除会让
 * 登记处丢掉落在被删区间里的条目（`IndexedTableRegistry.applyRowPlan` →
 * `remapByRowPlan`：新序列只按行计划逐位重建，被删的那几个下标根本不在计划里；
 * 列轴的 `remapByColumnShift` 则是 `remap` 返回 `undefined` 即丢弃）。重键那一半
 * 早就实现了，逆操作这一半在任务三之前没有落脚点——于是任务三只能拒绝求逆。
 *
 * **值是 `unknown`。** 侧表的载荷由插件定义，核心不解释它（`IndexedTable<T>` 的
 * `T` 同样只受存储约束）。写成 `unknown | undefined` 是多余的：`unknown` 本身就
 * 含 `undefined`，TypeScript 会把那个联合折叠成 `unknown`，反而让"`undefined`
 * 有特殊含义"这件事看起来像是类型保证的，其实不是——它由上面那段话保证。
 *
 * **没有插件侧表时是 `{}`，不是 `undefined`**：省得每个消费者都先判一次。
 *
 * **核心自己的 `rowData` / `columnData` 不进这里**：它们已经由 `meta` 字段承载，
 * 重复携带会在恢复时写两遍。
 */
export type SideTableEntries = Readonly<Record<string, readonly unknown[]>>;

/**
 * 本模型自己的变更类型。
 *
 * **为什么不复用 `src/core/types.ts` 的 `DataMutation`：** 那个类型今天由
 * `ArrayDataSource` 实现，带着对象数组时代的两个成员——`insert-row-objects`
 * （按字段名的行对象）与 `ColumnFieldMapping`（列带 `id`/`field`，而 ADR §3.8.2
 * 已判定列不再有身份）。改它等于提前开始替换，而替换必须原子（ADR §3.9）。
 * 两者在替换那一刻汇合。
 *
 * **三处实质差异：**
 * 1. 格子载 `CellRecord` 而非裸值——只带值的逆会丢掉 `t`/`s`/`f`/`si`；
 * 2. 没有 `insert-row-objects`；
 * 3. 列操作不带字段映射，插列只需要个数。
 *
 * `restore-*` 那几个成员**只作为逆变更出现**：删除的逆必须把原样的记录、元数据
 * 与**每一张插件侧表在被删区间里的条目**（`tables`，spec §11 不变量 8）带回来，
 * 光有 `start`/`count` 是复原不了的。
 */
export type SheetMutation =
  | { readonly type: "set-cell"; readonly row: number; readonly column: number; readonly record: CellRecord | undefined }
  | { readonly type: "set-range"; readonly range: CellRange; readonly records: readonly (readonly (CellRecord | undefined)[])[] }
  | { readonly type: "insert-rows"; readonly start: number; readonly rows: readonly RowCells[] }
  | { readonly type: "delete-rows"; readonly start: number; readonly count: number }
  | { readonly type: "restore-rows"; readonly start: number; readonly rows: readonly RowCells[]; readonly meta: readonly (RowMeta | undefined)[]; readonly tables: SideTableEntries }
  | { readonly type: "move-rows"; readonly start: number; readonly count: number; readonly target: number }
  | { readonly type: "reorder-rows"; readonly order: Int32Array }
  | { readonly type: "insert-columns"; readonly start: number; readonly count: number }
  | { readonly type: "delete-columns"; readonly start: number; readonly count: number }
  | { readonly type: "restore-columns"; readonly start: number; readonly count: number; readonly cells: readonly ColumnCells[]; readonly meta: readonly (ColumnMeta | undefined)[]; readonly tables: SideTableEntries }
  | { readonly type: "move-columns"; readonly start: number; readonly count: number; readonly target: number };

const STRUCTURAL = new Set<SheetMutation["type"]>([
  "insert-rows",
  "delete-rows",
  "restore-rows",
  "move-rows",
  "reorder-rows",
  "insert-columns",
  "delete-columns",
  "restore-columns",
  "move-columns"
]);

/**
 * 结构变更会移动下标，因此必须带着每一张按下标索引的侧表一起走
 * （ADR §9 不变量 4）。格子变更不移动下标。4B 的暂存层靠这个区分决定要不要
 * 重键已暂存的内容。
 */
export function isStructuralSheetMutation(mutation: SheetMutation): boolean {
  return STRUCTURAL.has(mutation.type);
}
