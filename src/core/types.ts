import type { ModelRevision } from "./model";
import type { Mutation, TransactionReadView } from "./mutations";

export interface CellRange {
  readonly startRow: number;
  readonly startColumn: number;
  readonly endRow: number;
  readonly endColumn: number;
}

export type RowData = Record<string, unknown>;

export interface DataSource {
  getRowCount(): number;
  getColumnCount(): number;
  getCell(row: number, column: number): unknown;
  setCell(row: number, column: number, value: unknown): void;
  getRows?(start: number, end: number): RowData[] | Promise<RowData[]>;
  /**
   * Optional structural mutation capability (Step 12) — a remote/computed
   * DataSource is free to leave all three unimplemented; TableStore checks
   * for their presence and throws a clear capability error instead of
   * assuming they exist. `rows` uses the same "one array of values per
   * row, in column order" convention as `setRange()`, not
   * `RowData` — field-name mapping is ArrayDataSource's own concern, never
   * TableStore's or a Command's.
   */
  insertRows?(start: number, rows: readonly (readonly unknown[])[]): void;
  deleteRows?(start: number, count: number): void;
  moveRows?(start: number, count: number, target: number): void;
  /**
   * Column-axis counterpart to insertRows/deleteRows/moveRows (Step 17) —
   * same optional-capability contract. `columns` is the same `{id, field}`
   * shape `ArrayDataSourceColumn` uses; declared inline (not imported from
   * `./data/ArrayDataSource`) since this interface stays implementation-
   * agnostic — TypeScript's structural typing makes the two interchangeable
   * without either needing to reference the other.
   */
  insertColumns?(start: number, columns: readonly ColumnFieldMapping[]): void;
  deleteColumns?(start: number, count: number): void;
  moveColumns?(start: number, count: number, target: number): void;
  /**
   * Optional whole-axis row permutation (Step 22 sorting). `order` follows the
   * repo-wide convention `order[newPosition] = oldPosition`. Kept separate from
   * `moveRows()` because a sort is not expressible as a cheap sequence of
   * moves, and because a DataSource may reasonably support one and not the
   * other.
   */
  reorderRows?(order: Int32Array): void;
  /** Optional strong transaction adapter. Mutable TableStore operations require it. */
  beginTransaction?(baseRevision: ModelRevision): DataSourceTransaction;
}

export interface ColumnFieldMapping {
  readonly id: string;
  readonly field: string;
}

export interface CellDescriptorSnapshot {
  readonly row: number;
  readonly column: number;
  readonly descriptor: PropertyDescriptor | undefined;
}

export type DataMutation =
  | { readonly type: "set-cell"; readonly row: number; readonly column: number; readonly value: unknown }
  | { readonly type: "set-range"; readonly range: CellRange; readonly values: readonly (readonly unknown[])[] }
  | { readonly type: "restore-cells"; readonly cells: readonly CellDescriptorSnapshot[] }
  /**
   * ADR 0004 spike: the record-native counterpart of `restore-cells`. A data
   * source that owns its cells restores them by putting the whole record
   * back, with no need to reproduce a property descriptor on somebody else's
   * object. `record` is `unknown` rather than a concrete cell-record type
   * because this module must not depend on a data source's internal shape —
   * the implementation narrows it.
   */
  | {
      readonly type: "restore-records";
      readonly cells: readonly {
        readonly row: number;
        readonly column: number;
        /** `undefined` means the cell was unoccupied: the inverse deletes it rather than writing `undefined`. */
        readonly record: unknown;
      }[];
    }
  | { readonly type: "insert-rows"; readonly start: number; readonly rows: readonly (readonly unknown[])[] }
  | { readonly type: "insert-row-objects"; readonly start: number; readonly rows: readonly RowData[] }
  | { readonly type: "delete-rows"; readonly start: number; readonly count: number }
  | { readonly type: "move-rows"; readonly start: number; readonly count: number; readonly target: number }
  | { readonly type: "reorder-rows"; readonly order: Int32Array }
  | { readonly type: "insert-columns"; readonly start: number; readonly columns: readonly ColumnFieldMapping[] }
  | { readonly type: "delete-columns"; readonly start: number; readonly count: number }
  | { readonly type: "move-columns"; readonly start: number; readonly count: number; readonly target: number };

export interface PreparedDataSourceMutation {
  readonly forward: DataMutation;
  readonly inverse: DataMutation;
  readonly estimatedBytes: number;
  commit(): void;
  rollback(): void;
  dispose(): void;
}

/** Source-owned scratch projection used while a Command transaction is planned. */
export interface DataSourceTransaction extends TransactionReadView {
  prepareMutation(mutation: DataMutation): PreparedDataSourceMutation;
  dispose(): void;
}

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
