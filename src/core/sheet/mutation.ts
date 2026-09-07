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
 *
 * **别名规则与 `CellRecord` 同一条**（`types.ts` 顶部那段规则的形态 2 与形态 3）：
 * 这里装的是插件载荷对象**本身的引用**，不是拷贝——捕获走的是 `IndexedTable.get()`，
 * 那是一次裸读。于是一条撤销记录会**别名到活的插件数据**：插件若在删除之后原地改了
 * 那个载荷对象，历史栈里这条记录跟着变。这是本仓库的既定规则（容器复制一层、值按
 * 引用共享），不是这里的疏漏；写在这儿是因为读者会先落到这个声明上。
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
 *
 * **`set-row-meta` / `set-column-meta` 是 4B 加的，理由不是"顺手补齐"。** 行高权威
 * 从 `LayoutEngine` 迁进 `rowData` 之后（spec §5.1），改行高就是在改 `SheetModel`
 * 内部的状态。若这条路径不走 `SheetMutation`，那么"同一个事务里既改行高又做结构
 * 变更"就成了两个参与者写同一份状态：结构那一半在影子模型上暂存，行高那一半直接
 * 写活模型，`MutationCoordinator` 的"prepare 期不得触碰活状态"当场破掉。列宽/列
 * 隐藏是同一根轴上的同一件事，所以两个成员一起加。
 */
export type SheetMutation =
  | { readonly type: "set-cell"; readonly row: number; readonly column: number; readonly record: CellRecord | undefined }
  | { readonly type: "set-range"; readonly range: CellRange; readonly records: readonly (readonly (CellRecord | undefined)[])[] }
  | { readonly type: "insert-rows"; readonly start: number; readonly rows: readonly RowCells[] }
  | { readonly type: "delete-rows"; readonly start: number; readonly count: number }
  | {
      readonly type: "restore-rows";
      readonly start: number;
      readonly rows: readonly RowCells[];
      readonly meta: readonly (RowMeta | undefined)[];
      readonly tables: SideTableEntries;
      readonly mergesBefore: readonly CellRange[];
    }
  | { readonly type: "move-rows"; readonly start: number; readonly count: number; readonly target: number }
  | { readonly type: "reorder-rows"; readonly order: Int32Array }
  | { readonly type: "insert-columns"; readonly start: number; readonly count: number }
  | { readonly type: "delete-columns"; readonly start: number; readonly count: number }
  | {
      readonly type: "restore-columns";
      readonly start: number;
      readonly count: number;
      readonly cells: readonly ColumnCells[];
      readonly meta: readonly (ColumnMeta | undefined)[];
      readonly tables: SideTableEntries;
      readonly mergesBefore: readonly CellRange[];
    }
  | { readonly type: "move-columns"; readonly start: number; readonly count: number; readonly target: number }
  // `meta: undefined` 表示"这一行/列原先在 `rowData` / `columnData` 里没有条目"，
  // 施加下去是**删条目**，不是写一个空的 `{}`——与 `set-cell` 的
  // `record: undefined` 同一口径（spec §11 不变量 1：不存空记录，稀疏性依赖这点）。
  | { readonly type: "set-row-meta"; readonly row: number; readonly meta: RowMeta | undefined }
  | { readonly type: "set-column-meta"; readonly column: number; readonly meta: ColumnMeta | undefined }
  // 合并区是整体替换的哑变更，与 `set-row-meta` / `set-column-meta` 同一类
  // （见下面 `isStructuralSheetMutation` 的口径）：不移动任何下标，只是换一整份
  // 状态。求逆同样是"读现在的整份 `merges()`"，见 `invert.ts`。
  | { readonly type: "set-merges"; readonly merges: readonly CellRange[] };

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
 * （ADR §9 不变量 4）。格子变更与元数据变更（`set-row-meta` / `set-column-meta`）
 * 都只改一个既有下标上的内容，不移动下标。4B 的暂存层靠这个区分决定要不要重键
 * 已暂存的内容。
 */
export function isStructuralSheetMutation(mutation: SheetMutation): boolean {
  return STRUCTURAL.has(mutation.type);
}
