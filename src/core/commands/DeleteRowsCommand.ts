import { BEFORE_ROWS_CHANGE_HOOK, HookManager, type HookRegistry } from "../hooks";
import type { TableStore } from "../data";
import type { Command, CommandContext, Patch } from "../types";
import { captureRange } from "./captureRange";
import { checkBeforeRowsChange } from "./checkBeforeRowsChange";

const STANDALONE_HOOK_REGISTRY: HookRegistry = { beforeRowsChange: BEFORE_ROWS_CHANGE_HOOK };

export class DeleteRowsCommand implements Command {
  readonly type = "delete-rows";

  readonly #store: TableStore;
  readonly #start: number;
  readonly #count: number;
  readonly #hooks: HookManager;
  #deletedValues: unknown[][] = [];

  constructor(store: TableStore, start: number, count: number, hooks?: HookManager) {
    this.#store = store;
    this.#start = start;
    this.#count = count;
    this.#hooks = hooks ?? new HookManager({ registry: STANDALONE_HOOK_REGISTRY });
  }

  execute(context: CommandContext): Patch | Patch[] {
    void context;
    checkBeforeRowsChange(this.#hooks, { operation: "delete", start: this.#start, count: this.#count });

    const columnCount = this.#store.getColumnCount();
    this.#deletedValues =
      columnCount === 0
        ? Array.from({ length: this.#count }, () => [])
        : captureRange(this.#store, {
            startRow: this.#start,
            endRow: this.#start + this.#count - 1,
            startColumn: 0,
            endColumn: columnCount - 1
          });

    return [...this.#store.deleteRows(this.#start, this.#count).patches];
  }

  undo(context: CommandContext): Patch | Patch[] {
    void context;
    return [...this.#store.insertRows(this.#start, this.#deletedValues).patches];
  }
}
