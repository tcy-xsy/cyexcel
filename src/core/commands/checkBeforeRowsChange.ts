import { CyExcelError, CyExcelErrorCode } from "../errors";
import { createCancellableContext, type CancellableContext, type HookManager } from "../hooks";

export interface BeforeRowsChangeContext extends CancellableContext {
  readonly operation: "insert" | "delete" | "move";
  readonly start: number;
  readonly count: number;
  readonly target?: number;
}

/**
 * Shared by InsertRowsCommand/DeleteRowsCommand/MoveRowsCommand — each
 * calls this at the very start of execute(), before touching the Store.
 * Throwing here (rather than returning a boolean) means the store is never
 * written to, so CommandManager.transaction()'s existing rollback path
 * needs nothing extra to guarantee "cancelled before commit means no
 * partial commit" for a multi-command transaction.
 */
export function checkBeforeRowsChange(
  hooks: HookManager,
  change: Omit<BeforeRowsChangeContext, keyof CancellableContext>
): void {
  const context = createCancellableContext(change);
  hooks.runBail("beforeRowsChange", context);
  if (context.cancelled) {
    throw new CyExcelError(
      CyExcelErrorCode.CommandCancelled,
      context.reason ?? `Row ${change.operation} was cancelled before commit`
    );
  }
}
