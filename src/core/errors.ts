/**
 * Stable error codes. Callers must branch on `code`, not on `message` text —
 * message wording may change between releases without being a breaking
 * change. New codes are appended as later steps need them; existing codes
 * are never renumbered or reused for a different meaning.
 */
export const CyExcelErrorCode = {
  InvalidCoordinate: "invalid-coordinate",
  InvalidRange: "invalid-range",
  InvalidOptions: "invalid-options",
  ColumnNotFound: "column-not-found",
  // `readonly-data-source` 与 `structural-modification-not-supported` 在 ADR 0004
  // 阶段 4B 一起作废：它唯一的抛出点是 `ArrayDataSource.setCell()` 里的
  // `options.readonly`，而那个类连同可替换的数据层一起删掉了（§3.2）。表级只读
  // 是另一件事（`setReadonly()` 至今没有实现），真做的时候按"码只增不改"的规矩
  // 加一个新码，不复用这个——那样才不会让"数据源只读"和"这张表被锁了"混在一个码里。
  CommandCancelled: "command-cancelled",
  AsyncRendererNotSupported: "async-renderer-not-supported",
  UnknownCellType: "unknown-cell-type",
  UnknownHook: "unknown-hook",
  AsyncHookNotSupported: "async-hook-not-supported",
  // `structural-modification-not-supported` 在 ADR 0004 阶段 4B 作废并删除：它答的是
  // "这个 DataSource 实现了 insertRows 吗"，而数据层不再可替换（§3.2），`SheetModel`
  // 七个结构操作全都在。这里不留一个永远不会被抛出的码——留着会让调用方以为还需要
  // 分支处理它。码是**删除**不是保留占位，因为它从未与别的码共用过语义。
  MutationBlockedByReadLease: "mutation-blocked-by-read-lease",
  MutationConflict: "mutation-conflict",
  MutationNotSupported: "mutation-not-supported",
  InvalidMutation: "invalid-mutation",
  ReentrantMutation: "reentrant-mutation",
  MutationRuntimeFaulted: "mutation-runtime-faulted",
  MutationRuntimeDisposed: "mutation-runtime-disposed",
  PluginRuntimeNotReady: "plugin-runtime-not-ready",
  PluginDisposed: "plugin-disposed",
  MergeOverlapsExistingMerge: "merge-overlaps-existing-merge",
  MergeSpansMoveBoundary: "merge-spans-move-boundary"
} as const;

export type CyExcelErrorCode =
  (typeof CyExcelErrorCode)[keyof typeof CyExcelErrorCode];

export class CyExcelError extends Error {
  readonly code: CyExcelErrorCode;

  constructor(code: CyExcelErrorCode, message: string) {
    super(message);
    this.name = "CyExcelError";
    this.code = code;
  }
}
