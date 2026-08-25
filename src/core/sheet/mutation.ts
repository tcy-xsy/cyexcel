import type { CellRange } from "../types";
import type { CellRecord, ColumnMeta, RowMeta } from "./types";

/** 一行的格子：列号 -> 记录。稀疏，未占用的列不出现（spec §11 不变量 1）。 */
export type RowCells = Readonly<Record<number, CellRecord>>;

/** 一列的格子：**行号** -> 记录。稀疏，未占用的行不出现。 */
export type ColumnCells = Readonly<Record<number, CellRecord>>;

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
 * `restore-*` 那几个成员**只作为逆变更出现**：删除的逆必须把原样的记录与元数据
 * 带回来，光有 `start`/`count` 是复原不了的。
 */
export type SheetMutation =
  | { readonly type: "set-cell"; readonly row: number; readonly column: number; readonly record: CellRecord | undefined }
  | { readonly type: "set-range"; readonly range: CellRange; readonly records: readonly (readonly (CellRecord | undefined)[])[] }
  | { readonly type: "insert-rows"; readonly start: number; readonly rows: readonly RowCells[] }
  | { readonly type: "delete-rows"; readonly start: number; readonly count: number }
  | { readonly type: "restore-rows"; readonly start: number; readonly rows: readonly RowCells[]; readonly meta: readonly (RowMeta | undefined)[] }
  | { readonly type: "move-rows"; readonly start: number; readonly count: number; readonly target: number }
  | { readonly type: "reorder-rows"; readonly order: Int32Array }
  | { readonly type: "insert-columns"; readonly start: number; readonly count: number }
  | { readonly type: "delete-columns"; readonly start: number; readonly count: number }
  | { readonly type: "restore-columns"; readonly start: number; readonly count: number; readonly cells: readonly ColumnCells[]; readonly meta: readonly (ColumnMeta | undefined)[] }
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
