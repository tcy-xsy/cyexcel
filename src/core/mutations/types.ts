import type { ModelRevision } from "../model";
import type { ChangeSet } from "../data/changes";
import type { SheetMutation } from "../sheet/mutation";
import type { SheetModel } from "../sheet/SheetModel";
import type { PreparedSheetMutation } from "../sheet/SheetTransaction";
import type { ColumnMeta, RowMeta } from "../sheet/types";
import type { CellRange, Patch } from "../types";

export type MutationCoordinatorState =
  | "idle"
  | "preparing"
  | "prepared"
  | "committing"
  | "faulted"
  | "disposed";

/** A replayable, fully declarative unit handed to the commit coordinator. */
export interface Mutation {
  readonly type: string;
  readonly changes: readonly Patch[];
  prepare(context: MutationPrepareContext): PreparedMutationParticipant;
}

export interface MutationPrepareContext {
  readonly baseRevision: ModelRevision;
  readonly model: TransactionReadView;
  /**
   * 把一条变更暂存进本次事务的影子模型，拿回一个可 commit / rollback 的句柄。
   *
   * `source` 是**归属校验**用的：一个协调器一次只服务一张表，交进来的模型不是那
   * 一张就说明调用方拿错了对象——这种错早一步抛出来，比在影子上悄悄写另一张表好。
   */
  prepareData(source: SheetModel, mutation: SheetMutation): PreparedSheetMutation;
  /**
   * 读影子模型上的行 / 列元数据，`source` 的归属校验与 `prepareData` 同一套。
   *
   * 存在的理由是行高的权威在 `rowData`（ADR 0004 阶段 4B T8）：一条只改高度的
   * 变更要写回整条 `RowMeta`，就得先看见这一行现有的 `hd`/`hp`，而看的必须是影子
   * ——同一个事务里的前一条变更已经改过它了（设计稿 §4.3）。
   *
   * 没有并进 `TransactionReadView`：那个接口是公开 API（`check:public-api` 盯着），
   * 而且有别的实现方；这里沿用 `SheetTransaction.getRecord` 的做法，在读视图之外多开。
   */
  readRowMeta(source: SheetModel, row: number): RowMeta | undefined;
  readColumnMeta(source: SheetModel, column: number): ColumnMeta | undefined;
  readonly roots: MutationRootDraftRegistry;
}

/** Shared transaction-local roots let independent mutations compose on one owner. */
export interface MutationRootDraftRegistry {
  get<TRoot>(owner: object, fallback: () => TRoot): TRoot;
  set<TRoot>(owner: object, root: TRoot): void;
}

export interface MutationPlanningResources {
  readonly model: TransactionReadView;
  readonly prepareData: MutationPrepareContext["prepareData"];
  readonly readRowMeta: MutationPrepareContext["readRowMeta"];
  readonly readColumnMeta: MutationPrepareContext["readColumnMeta"];
  dispose(): void;
}

/**
 * A participant prepared without touching live state. rollback() must be
 * idempotent and must also handle a commit() call that started and threw.
 */
export interface PreparedMutationParticipant {
  /** Explicit live-state contract; Patch presence is not an effect signal. */
  readonly hasEffect: boolean;
  readonly forward: Mutation;
  readonly inverse: Mutation;
  readonly estimatedBytes: number;
  commit(): void;
  rollback(): void;
  dispose(): void;
}

export interface PreparedTransaction {
  readonly transactionId: string;
  readonly baseRevision: ModelRevision;
  readonly preview: TransactionReadView;
  readonly hasEffect: boolean;
  readonly participants: readonly PreparedMutationParticipant[];
  readonly changes: readonly Patch[];
  readonly estimatedHistoryBytes: number;
  dispose(): void;
}

export interface CommitReceipt {
  readonly transactionId: string;
  readonly noop: boolean;
  readonly beforeRevision: ModelRevision;
  readonly afterRevision: ModelRevision;
  readonly forward: readonly Mutation[];
  readonly inverse: readonly Mutation[];
  readonly patches: readonly Patch[];
  readonly historyBytes: number;
  readonly changeSet: ChangeSet;
}

export interface TransactionReadView {
  readonly baseRevision: ModelRevision;
  getRowCount(): number;
  getColumnCount(): number;
  getCell(row: number, column: number): unknown;
  /**
   * 当前（暂存后）的合并区表。`merge-cells`/`unmerge-cells` 的 Command handler 在
   * plan 期靠它读到"目前有哪些合并区"，去做重叠/吸收判断（`../sheet/merges` 的
   * `resolveMergeCells`）——这条读法与 `getRowCount`/`getColumnCount` 同一个理由，
   * 不查 dispose：答的是"这次事务规划完的那个状态"，不是一个会过期的数。
   */
  merges(): readonly CellRange[];
}

/** Planning session whose preview advances after every staged mutation. */
export interface MutationPrepareSession {
  readonly transactionId: string;
  readonly baseRevision: ModelRevision;
  readonly model: TransactionReadView;
  stage(mutation: Mutation): void;
  seal(): PreparedTransaction;
  discard(): void;
}

export interface StateMutationOptions {
  readonly type: string;
  readonly changes?: readonly Patch[];
  readonly inverseChanges?: readonly Patch[];
  readonly apply: () => void;
  readonly revert: () => void;
  readonly estimatedBytes?: number;
}
