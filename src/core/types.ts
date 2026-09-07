import type { Mutation, TransactionReadView } from "./mutations";

export interface CellRange {
  readonly startRow: number;
  readonly startColumn: number;
  readonly endRow: number;
  readonly endColumn: number;
}

/*
 * `RowData` 随 `CyExcelCoreOptions.data` 一起在 4B（T7）删除：入口只吃快照了。
 * 对象数组的那一头搬去了导入器，叫 `sheet/import.ts` 的 `ImportRow`——它是导入
 * 这件事的参数类型，不是核心的数据形状（ADR 0004 §3.2）。
 *
 * `DataSource` / `DataMutation` / `ColumnFieldMapping` / `CellDescriptorSnapshot` /
 * `PreparedDataSourceMutation` / `DataSourceTransaction` 在 4B（T5）整体删除。
 *
 * 它们描述的是"数据在一个可替换的外部对象里，核心通过接口去改它"这套模型，而
 * ADR 0004 §3.2 判定调用方不能替换数据层：存储就是 `SheetModel`，变更就是
 * `SheetMutation`，暂存就是 `SheetTransaction`。留一个接口壳子当过渡层是白做的工
 * （§0 决定 1）。对应物：
 *
 * - `DataSource` → `src/core/sheet/SheetModel.ts`
 * - `DataMutation` → `src/core/sheet/mutation.ts` 的 `SheetMutation`
 * - `DataSourceTransaction` / `PreparedDataSourceMutation` →
 *   `src/core/sheet/SheetTransaction.ts`
 * - `ColumnFieldMapping` → 没有对应物：列不再有身份（§3.8.2），字段名是导入那一刻
 *   的参数（`sheet/import.ts`）
 * - `CellDescriptorSnapshot` → 没有对应物：它是"就地改调用方数组"时代用来复原属性
 *   描述符的，稀疏矩阵里删格就是删键
 */

/**
 * Minimal shape shared by every Command; CommandManager (Step 3) will add
 * fields as they become necessary, not replace this one.
 */
export interface CommandContext {
  readonly transactionId: string;
  readonly model: TransactionReadView;
  stage(mutation: Mutation): void;
}

export type Patch =
  | { readonly type: "cell-value"; readonly row: number; readonly column: number }
  | { readonly type: "cell-style"; readonly row: number; readonly column: number }
  | { readonly type: "range-value"; readonly range: CellRange }
  | {
      readonly type: "row-structure";
      readonly operation: "insert" | "delete" | "move";
      readonly start: number;
      readonly count: number;
      readonly target?: number;
    }
  | {
      readonly type: "column-structure";
      readonly operation: "insert" | "delete" | "move";
      readonly start: number;
      readonly count: number;
      readonly target?: number;
    }
  | { readonly type: "column-layout"; readonly column: number }
  /**
   * A whole-sheet "merges changed" marker, carrying no fields — like
   * `column-layout`, it does not describe a value change or an index shift,
   * so downstream treats it as a full repaint rather than an incremental
   * range (doc/merge-cells-plan.md §Task 5).
   */
  | { readonly type: "merge-layout" }
  /**
   * One atomic row-height mutation. `rows` names every changed row while
   * `range` starts at the earliest one because its geometry shift affects
   * every row below it.
   */
  | {
      readonly type: "row-layout";
      readonly rows: readonly [number, ...number[]];
      readonly range: CellRange;
    }
  | { readonly type: "full-reset" }
  /**
   * A whole-axis row permutation (`order[newPosition] = oldPosition`).
   *
   * Deliberately its own variant rather than a fourth `row-structure`
   * operation: `remapIndex()` and every consumer of `row-structure` assume
   * `start`/`count` semantics that a permutation does not have, so widening
   * that variant would compile while silently remapping indices wrongly.
   */
  | { readonly type: "row-reorder"; readonly order: Int32Array };

export interface Command<TType extends string = string, TParams = unknown> {
  readonly type: TType;
  readonly params: TParams;
}

export type Disposer = () => void;

/**
 * Shared vocabulary for the Hook system. Concrete hook names, execution
 * semantics (series/bail/waterfall/parallel) and contexts are defined in
 * Step 11's HookManager, not here.
 */
export type HookHandler<TContext = unknown> = (context: TContext) => void;

/**
 * Shared vocabulary for the Event system. Concrete event names and payloads
 * are defined per module as each one lands.
 */
export type EventListener<TPayload = unknown> = (payload: TPayload) => void;
