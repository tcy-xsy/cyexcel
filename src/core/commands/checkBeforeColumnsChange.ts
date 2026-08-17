import { CyExcelError, CyExcelErrorCode } from "../errors";
import { createCancellableContext, type CancellableContext, type HookManager } from "../hooks";

export interface BeforeColumnsChangeContext extends CancellableContext {
  readonly operation: "insert" | "delete" | "move";
  readonly start: number;
  readonly count: number;
  readonly target?: number;
}

/**
 * Column-axis counterpart to checkBeforeRowsChange.ts — same reasoning:
 * throwing here (before touching Store/LayoutEngine) means
 * CommandManager.transaction()'s existing rollback path needs nothing
 * extra to guarantee "cancelled before commit means no partial commit".
 */
export function checkBeforeColumnsChange(
  hooks: HookManager,
  change: Omit<BeforeColumnsChangeContext, keyof CancellableContext>
): void {
  const context = createCancellableContext(change);
  hooks.runBail("beforeColumnsChange", context);
  if (context.cancelled) {
    throw new CyExcelError(
      CyExcelErrorCode.CommandCancelled,
      context.reason ?? `Column ${change.operation} was cancelled before commit`
    );
  }
}
