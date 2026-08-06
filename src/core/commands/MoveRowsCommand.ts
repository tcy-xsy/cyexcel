import { BEFORE_ROWS_CHANGE_HOOK, HookManager, type HookRegistry } from "../hooks";
import type { TableStore } from "../data";
import type { Command, CommandContext, Patch } from "../types";
import { checkBeforeRowsChange } from "./checkBeforeRowsChange";

const STANDALONE_HOOK_REGISTRY: HookRegistry = { beforeRowsChange: BEFORE_ROWS_CHANGE_HOOK };

/** `target` is the insertion index in the pre-move index space — same convention `ArrayDataSource.moveRows()`/`remapRowIndex()` use. */
export class MoveRowsCommand implements Command {
  readonly type = "move-rows";

  readonly #store: TableStore;
  readonly #start: number;
  readonly #count: number;
  readonly #target: number;
  readonly #hooks: HookManager;

  constructor(store: TableStore, start: number, count: number, target: number, hooks?: HookManager) {
    this.#store = store;
    this.#start = start;
    this.#count = count;
    this.#target = target;
    this.#hooks = hooks ?? new HookManager({ registry: STANDALONE_HOOK_REGISTRY });
  }

  execute(context: CommandContext): Patch | Patch[] {
    void context;
    checkBeforeRowsChange(this.#hooks, {
      operation: "move",
      start: this.#start,
      count: this.#count,
      target: this.#target
    });
    return [...this.#store.moveRows(this.#start, this.#count, this.#target).patches];
  }

  /**
   * The moved block now sits at `#adjustedTarget()`. Moving it back isn't
   * simply "moveRows(adjustedTarget, count, start)" in every case — that
   * only holds when the forward move went *later* (target > start, so
   * adjustedTarget >= start, and the undo is itself an "earlier" move
   * where moveRows()'s target is used as-is). When the forward move went
   * *earlier* (target <= start, so adjustedTarget <= start), undoing is a
   * "later" move, and moveRows() subtracts `count` from any target past
   * its own start — so undoing to land back exactly on `start` requires
   * passing `start + count`, not `start`. Both cases were verified by
   * hand against `[A,B,C,D,E]`-style examples; see doc/progress.md's
   * Step 12 entry.
   */
  undo(context: CommandContext): Patch | Patch[] {
    void context;
    const adjustedTarget = this.#adjustedTarget();
    const undoTarget = this.#start > adjustedTarget ? this.#start + this.#count : this.#start;
    return [...this.#store.moveRows(adjustedTarget, this.#count, undoTarget).patches];
  }

  #adjustedTarget(): number {
    return this.#target > this.#start ? this.#target - this.#count : this.#target;
  }
}
