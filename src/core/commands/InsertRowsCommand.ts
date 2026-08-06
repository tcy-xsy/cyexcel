import { BEFORE_ROWS_CHANGE_HOOK, HookManager, type HookRegistry } from "../hooks";
import type { TableStore } from "../data";
import type { Command, CommandContext, Patch } from "../types";
import { checkBeforeRowsChange } from "./checkBeforeRowsChange";

/** Used when no shared HookManager is injected (see CommandManager/ColumnManager's identical fallback pattern). */
const STANDALONE_HOOK_REGISTRY: HookRegistry = { beforeRowsChange: BEFORE_ROWS_CHANGE_HOOK };

/** `rows` is column-ordered values, one array per new row (same convention as SetRangeCommand) — DataSource.insertRows() owns the field-name mapping. */
export class InsertRowsCommand implements Command {
  readonly type = "insert-rows";

  readonly #store: TableStore;
  readonly #start: number;
  readonly #rows: readonly (readonly unknown[])[];
  readonly #hooks: HookManager;

  constructor(
    store: TableStore,
    start: number,
    rows: readonly (readonly unknown[])[],
    hooks?: HookManager
  ) {
    this.#store = store;
    this.#start = start;
    this.#rows = rows;
    this.#hooks = hooks ?? new HookManager({ registry: STANDALONE_HOOK_REGISTRY });
  }

  execute(context: CommandContext): Patch | Patch[] {
    void context;
    checkBeforeRowsChange(this.#hooks, { operation: "insert", start: this.#start, count: this.#rows.length });
    return [...this.#store.insertRows(this.#start, this.#rows).patches];
  }

  undo(context: CommandContext): Patch | Patch[] {
    void context;
    return [...this.#store.deleteRows(this.#start, this.#rows.length).patches];
  }
}
