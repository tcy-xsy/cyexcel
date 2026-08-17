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
  ReadonlyDataSource: "readonly-data-source",
  CommandCancelled: "command-cancelled",
  AsyncRendererNotSupported: "async-renderer-not-supported",
  UnknownCellType: "unknown-cell-type",
  UnknownHook: "unknown-hook",
  AsyncHookNotSupported: "async-hook-not-supported",
  StructuralModificationNotSupported: "structural-modification-not-supported",
  HeaderGroupsBlockColumnStructureChange: "header-groups-block-column-structure-change",
  MutationBlockedByReadLease: "mutation-blocked-by-read-lease",
  MutationConflict: "mutation-conflict",
  MutationNotSupported: "mutation-not-supported",
  ReentrantMutation: "reentrant-mutation",
  MutationRuntimeFaulted: "mutation-runtime-faulted",
  MutationRuntimeDisposed: "mutation-runtime-disposed",
  PluginRuntimeNotReady: "plugin-runtime-not-ready",
  PluginDisposed: "plugin-disposed"
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
