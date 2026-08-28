import type { ModelRevision } from "../model";
import type { ChangeSet } from "../data/changes";
import type { SheetMutation } from "../sheet/mutation";
import type { SheetModel } from "../sheet/SheetModel";
import type { PreparedSheetMutation } from "../sheet/SheetTransaction";
import type { Patch } from "../types";

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
