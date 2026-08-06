import { mergePatches, type TableStore } from "../data";
import { CyExcelError, CyExcelErrorCode } from "../errors";
import { HistoryManager, type HistoryEntry } from "../history";
import {
  BEFORE_COMMAND_HOOK,
  BEFORE_TRANSACTION_COMMIT_HOOK,
  HookManager,
  createCancellableContext,
  type CancellableContext,
  type HookRegistry
} from "../hooks";
import type { Command, CommandContext, Patch } from "../types";

/** Context for the `beforeCommand` Hook — fires once per Command, before it executes. */
export interface BeforeCommandContext extends CancellableContext {
  readonly command: Command;
  readonly transactionId: string;
}

/** Context for the `beforeTransactionCommit` Hook — fires once per transaction, after every Command in it has executed, before the History entry is written. */
export interface BeforeTransactionCommitContext extends CancellableContext {
  readonly commands: readonly Command[];
  readonly patches: readonly Patch[];
  readonly transactionId: string;
}

/** Used when no shared HookManager is injected — just enough of the registry for this module's own two hook names, defined from the same constants CyExcelCore's production registry uses (see src/core/hooks/registry.ts). */
const STANDALONE_HOOK_REGISTRY: HookRegistry = {
  beforeCommand: BEFORE_COMMAND_HOOK,
  beforeTransactionCommit: BEFORE_TRANSACTION_COMMIT_HOOK
};

export interface CommandManagerOptions {
  readonly historyLimit?: number;
  readonly historyByteLimit?: number;
  /**
   * Shares a HookManager with the rest of a CyExcelCore instance so
   * `sheet.hooks.tap("beforeCommand", ...)` actually participates in real
   * command execution. Defaults to a private HookManager scoped to just
   * `beforeCommand`/`beforeTransactionCommit`, so this class stays fully
   * standalone-constructible for tests that don't care about Hooks.
   */
  readonly hooks?: HookManager;
  /**
   * Sugar for `hooks.tap("beforeCommand", ...)`, kept for backward
   * compatibility with Step 3's original shape — checked before every
   * single command executes (not once per transaction). Returning `false`
   * cancels the whole transaction: already-executed commands in it are
   * rolled back, nothing is written, and no History entry is recorded.
   */
  readonly beforeCommit?: (command: Command, context: CommandContext) => boolean;
}

interface TransactionFrame {
  readonly transactionId: string;
  readonly commands: Command[];
  readonly patches: Patch[];
}

/**
 * Sole place that turns Commands into Store writes. Mouse, keyboard, paste,
 * and the public API are all expected to funnel through this one instance.
 */
export class CommandManager {
  readonly #store: TableStore;
  readonly #history: HistoryManager;
  readonly #hooks: HookManager;
  #frame: TransactionFrame | null = null;
  #nextTransactionId = 1;

  constructor(store: TableStore, options: CommandManagerOptions = {}) {
    this.#store = store;
    this.#history = new HistoryManager({
      maxEntries: options.historyLimit,
      maxBytes: options.historyByteLimit
    });
    this.#hooks = options.hooks ?? new HookManager({ registry: STANDALONE_HOOK_REGISTRY });

    if (options.beforeCommit) {
      const beforeCommit = options.beforeCommit;
      this.#hooks.tap(
        "beforeCommand",
        { name: "legacy-beforeCommit-option" },
        (context: BeforeCommandContext) => {
          if (!beforeCommit(context.command, { transactionId: context.transactionId })) {
            context.cancel();
          }
        }
      );
    }
  }

  get store(): TableStore {
    return this.#store;
  }

  get historySize(): number {
    return this.#history.size;
  }

  canUndo(): boolean {
    return this.#history.canUndo();
  }

  canRedo(): boolean {
    return this.#history.canRedo();
  }

  clearHistory(): void {
    this.#history.clear();
  }

  /** Runs a single command as its own one-command transaction. */
  execute(command: Command): readonly Patch[] {
    return this.transaction(() => this.#runCommand(command));
  }

  /**
   * Groups every command executed inside `fn` into one History entry. A
   * nested `transaction()` call (while already inside one) reuses the
   * outer frame instead of starting a new one — only the outermost call
   * commits or rolls back. `fn` must be synchronous.
   */
  transaction(fn: () => void): readonly Patch[] {
    const isOutermost = this.#frame === null;
    const frame: TransactionFrame =
      this.#frame ??
      (this.#frame = {
        transactionId: this.#createTransactionId(),
        commands: [],
        patches: []
      });

    try {
      fn();
    } catch (error) {
      if (isOutermost) {
        this.#rollback(frame);
        this.#frame = null;
      }
      throw error;
    }

    if (!isOutermost) {
      return frame.patches;
    }

    this.#frame = null;

    if (frame.commands.length === 0) {
      return [];
    }

    const commitContext = createCancellableContext<Omit<BeforeTransactionCommitContext, keyof CancellableContext>>({
      commands: frame.commands,
      patches: frame.patches,
      transactionId: frame.transactionId
    });
    this.#hooks.runBail("beforeTransactionCommit", commitContext);
    if (commitContext.cancelled) {
      this.#rollback(frame);
      throw new CyExcelError(
        CyExcelErrorCode.CommandCancelled,
        commitContext.reason ?? "Transaction was cancelled before commit"
      );
    }

    const mergedPatches = mergePatches(frame.patches);
    this.#history.push({ commands: frame.commands, patches: mergedPatches });
    return mergedPatches;
  }

  undo(): readonly Patch[] | undefined {
    const entry = this.#history.undo();
    return entry ? this.#applyUndo(entry) : undefined;
  }

  redo(): readonly Patch[] | undefined {
    const entry = this.#history.redo();
    return entry ? this.#applyRedo(entry) : undefined;
  }

  #runCommand(command: Command): void {
    // Only ever invoked as the `fn` passed to transaction(), which sets
    // #frame before calling fn() — never null in practice.
    const frame = this.#frame as TransactionFrame;

    const beforeCommandContext = createCancellableContext<Omit<BeforeCommandContext, keyof CancellableContext>>({
      command,
      transactionId: frame.transactionId
    });
    this.#hooks.runBail("beforeCommand", beforeCommandContext);
    if (beforeCommandContext.cancelled) {
      throw new CyExcelError(
        CyExcelErrorCode.CommandCancelled,
        beforeCommandContext.reason ?? `Command "${command.type}" was cancelled before commit`
      );
    }

    const context: CommandContext = { transactionId: frame.transactionId };
    const patches = command.execute(context);
    frame.commands.push(command);
    frame.patches.push(...(Array.isArray(patches) ? patches : [patches]));
  }

  #rollback(frame: TransactionFrame): void {
    const context: CommandContext = { transactionId: frame.transactionId };
    for (let index = frame.commands.length - 1; index >= 0; index -= 1) {
      frame.commands[index]?.undo?.(context);
    }
  }

  #applyUndo(entry: HistoryEntry): readonly Patch[] {
    const context: CommandContext = { transactionId: this.#createTransactionId() };
    const patches: Patch[] = [];
    for (let index = entry.commands.length - 1; index >= 0; index -= 1) {
      const result = entry.commands[index]?.undo?.(context);
      if (result) {
        patches.push(...(Array.isArray(result) ? result : [result]));
      }
    }
    return mergePatches(patches);
  }

  #applyRedo(entry: HistoryEntry): readonly Patch[] {
    const context: CommandContext = { transactionId: this.#createTransactionId() };
    const patches: Patch[] = [];
    for (const command of entry.commands) {
      const result = command.execute(context);
      patches.push(...(Array.isArray(result) ? result : [result]));
    }
    return mergePatches(patches);
  }

  #createTransactionId(): string {
    const id = `t-${this.#nextTransactionId}`;
    this.#nextTransactionId += 1;
    return id;
  }
}
