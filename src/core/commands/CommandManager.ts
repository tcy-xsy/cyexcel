import { mergePatches, type TableStore } from "../data";
import { CyExcelError, CyExcelErrorCode } from "../errors";
import { HistoryManager, type HistoryEntry } from "../history";
import type { Command, CommandContext, Patch } from "../types";

export interface CommandManagerOptions {
  readonly historyLimit?: number;
  readonly historyByteLimit?: number;
  /**
   * Minimal synchronous cancellation point, checked before every single
   * command executes (not once per transaction). Returning `false` cancels
   * the whole transaction: already-executed commands in it are rolled back,
   * nothing is written, and no History entry is recorded. This is a plain
   * constructor option, not a registrable/stackable Hook API — the real
   * HookManager lands in Step 11 and will subsume this.
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
  readonly #beforeCommit:
    | ((command: Command, context: CommandContext) => boolean)
    | undefined;
  #frame: TransactionFrame | null = null;
  #nextTransactionId = 1;

  constructor(store: TableStore, options: CommandManagerOptions = {}) {
    this.#store = store;
    this.#history = new HistoryManager({
      maxEntries: options.historyLimit,
      maxBytes: options.historyByteLimit
    });
    this.#beforeCommit = options.beforeCommit;
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

    const context: CommandContext = { transactionId: frame.transactionId };

    if (this.#beforeCommit && !this.#beforeCommit(command, context)) {
      throw new CyExcelError(
        CyExcelErrorCode.CommandCancelled,
        `Command "${command.type}" was cancelled before commit`
      );
    }

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
