import type { ChangeSet } from "../data/changes";
import { mergePatches } from "../data/patch";
import { CyExcelError, CyExcelErrorCode } from "../errors";
import { HistoryManager } from "../history";
import {
  BEFORE_COMMAND_HOOK,
  BEFORE_TRANSACTION_COMMIT_HOOK,
  HookManager,
  createCancellableContext,
  type CancellableContext,
  type HookRegistry
} from "../hooks";
import type {
  CommitReceipt,
  Mutation,
  MutationPrepareSession,
  PreparedTransaction,
  TransactionReadView
} from "../mutations";
import type { Command, CommandContext, Patch } from "../types";
import type { CommandPlanner } from "./CommandHandlerRegistry";

export interface BeforeCommandContext extends CancellableContext {
  readonly command: Command;
  readonly transactionId: string;
}

export interface BeforeTransactionCommitContext extends CancellableContext {
  readonly commands: readonly Command[];
  readonly patches: readonly Patch[];
  readonly model: TransactionReadView;
  readonly transactionId: string;
}

const STANDALONE_HOOK_REGISTRY: HookRegistry = {
  beforeCommand: BEFORE_COMMAND_HOOK,
  beforeTransactionCommit: BEFORE_TRANSACTION_COMMIT_HOOK
};

export interface CommandManagerOptions {
  readonly historyLimit?: number;
  readonly historyByteLimit?: number;
  readonly hooks?: HookManager;
}

export type CommandHistoryPolicy = "record" | "skip";

export interface CommandExecutionOptions {
  /** Defaults to `record`; nested calls inherit the active transaction policy. */
  readonly history?: CommandHistoryPolicy;
}

export interface CommandTransactionHost {
  begin(transactionId: string): MutationPrepareSession;
  commit(prepared: PreparedTransaction): CommitReceipt;
  discard(prepared: PreparedTransaction): void;
}

interface TransactionFrame {
  readonly transactionId: string;
  readonly history: CommandHistoryPolicy;
  readonly commands: Command[];
  readonly session: MutationPrepareSession;
}

export class CommandManager {
  readonly #host: CommandTransactionHost;
  readonly #planner: CommandPlanner;
  readonly #history: HistoryManager;
  readonly #hooks: HookManager;
  #frame: TransactionFrame | null = null;
  #acceptingCommands = false;
  #planningCommand = false;
  #lastChangeSet: ChangeSet | undefined;
  #nextTransactionId = 1;

  constructor(
    host: CommandTransactionHost,
    planner: CommandPlanner,
    options: CommandManagerOptions = {}
  ) {
    this.#host = host;
    this.#planner = planner;
    this.#history = new HistoryManager({
      maxEntries: options.historyLimit,
      maxBytes: options.historyByteLimit
    });
    this.#hooks = options.hooks ?? new HookManager({ registry: STANDALONE_HOOK_REGISTRY });
  }

  get historySize(): number {
    return this.#history.size;
  }

  get lastChangeSet(): ChangeSet | undefined {
    return this.#lastChangeSet;
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

  execute(command: Command, options: CommandExecutionOptions = {}): readonly Patch[] {
    return this.transaction(() => this.#planCommand(command), options);
  }

  transaction(fn: () => void, options: CommandExecutionOptions = {}): readonly Patch[] {
    if (this.#frame && !this.#acceptingCommands) {
      throw new CyExcelError(CyExcelErrorCode.ReentrantMutation, "Cannot enter a transaction while it is preparing or committing");
    }
    const isOutermost = this.#frame === null;
    if (
      !isOutermost &&
      options.history !== undefined &&
      options.history !== this.#frame?.history
    ) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidOptions,
        `Nested command history policy "${options.history}" conflicts with active transaction policy "${this.#frame?.history}"`
      );
    }
    const transactionId = this.#frame?.transactionId ?? this.#createTransactionId();
    const frame = this.#frame ?? {
      transactionId,
      history: options.history ?? "record",
      commands: [],
      session: this.#host.begin(transactionId)
    };
    if (isOutermost) {
      this.#frame = frame;
      this.#acceptingCommands = true;
    }

    try {
      fn();
      if (!isOutermost) {
        return [];
      }
      this.#acceptingCommands = false;
      if (frame.commands.length === 0) {
        frame.session.discard();
        this.#lastChangeSet = undefined;
        return [];
      }
      const prepared = frame.session.seal();
      return this.#hookAndCommit(frame, prepared);
    } catch (error) {
      if (isOutermost) {
        frame.session.discard();
      }
      throw error;
    } finally {
      if (isOutermost) {
        this.#frame = null;
        this.#acceptingCommands = false;
      }
    }
  }

  undo(): readonly Patch[] | undefined {
    const entry = this.#history.peekUndo();
    if (!entry) {
      this.#lastChangeSet = undefined;
      return undefined;
    }
    const replay = this.#replay(entry.inverse);
    this.#history.commitUndo(entry);
    this.#lastChangeSet = replay.changeSet;
    return mergePatches(replay.patches);
  }

  redo(): readonly Patch[] | undefined {
    const entry = this.#history.peekRedo();
    if (!entry) {
      this.#lastChangeSet = undefined;
      return undefined;
    }
    const replay = this.#replay(entry.forward);
    this.#history.commitRedo(entry);
    this.#lastChangeSet = replay.changeSet;
    return mergePatches(replay.patches);
  }

  #planCommand(command: Command): void {
    const frame = this.#frame;
    if (!frame || !this.#acceptingCommands) {
      throw new CyExcelError(CyExcelErrorCode.ReentrantMutation, "Commands can only be planned in an active transaction");
    }
    if (this.#planningCommand) {
      throw new CyExcelError(CyExcelErrorCode.ReentrantMutation, "Cannot execute a command from a planning hook");
    }
    this.#planningCommand = true;
    try {
      const beforeCommandContext = createCancellableContext<Omit<BeforeCommandContext, keyof CancellableContext>>({
        command,
        transactionId: frame.transactionId
      });
      this.#hooks.runBail("beforeCommand", beforeCommandContext);
      if (beforeCommandContext.cancelled) {
        throw new CyExcelError(
          CyExcelErrorCode.CommandCancelled,
          beforeCommandContext.reason ?? `Command "${command.type}" was cancelled before prepare`
        );
      }

      this.#planner.plan(command, this.#createContext(frame));
      frame.commands.push(command);
    } finally {
      this.#planningCommand = false;
    }
  }

  #hookAndCommit(frame: TransactionFrame, prepared: PreparedTransaction): readonly Patch[] {
    try {
      const context = createCancellableContext<Omit<BeforeTransactionCommitContext, keyof CancellableContext>>({
        commands: Object.freeze([...frame.commands]),
        patches: prepared.changes,
        model: prepared.preview,
        transactionId: frame.transactionId
      });
      this.#hooks.runBail("beforeTransactionCommit", context);
      if (context.cancelled) {
        throw new CyExcelError(
          CyExcelErrorCode.CommandCancelled,
          context.reason ?? "Transaction was cancelled before commit"
        );
      }
    } catch (error) {
      this.#host.discard(prepared);
      this.#lastChangeSet = undefined;
      throw error;
    }

    let receipt: CommitReceipt;
    try {
      receipt = this.#host.commit(prepared);
    } catch (error) {
      this.#host.discard(prepared);
      this.#lastChangeSet = undefined;
      throw error;
    }
    if (frame.history === "record" && !receipt.noop) {
      this.#history.push(receipt);
    }
    this.#lastChangeSet = receipt.changeSet;
    return mergePatches(receipt.patches);
  }

  #createContext(frame: TransactionFrame): CommandContext {
    return {
      transactionId: frame.transactionId,
      model: frame.session.model,
      stage: (mutation) => frame.session.stage(mutation)
    };
  }

  #replay(mutations: readonly Mutation[]): CommitReceipt {
    const session = this.#host.begin(this.#createTransactionId());
    try {
      for (const mutation of mutations) {
        session.stage(mutation);
      }
      const prepared = session.seal();
      try {
        return this.#host.commit(prepared);
      } catch (error) {
        this.#host.discard(prepared);
        throw error;
      }
    } catch (error) {
      session.discard();
      throw error;
    }
  }

  #createTransactionId(): string {
    return `t-${this.#nextTransactionId++}`;
  }
}
