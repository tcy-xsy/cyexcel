import type { Command } from "../types";

export interface ReorderRowsCommandParams {
  readonly order: Int32Array;
}

/**
 * Physically permutes the whole row axis. `order[newPosition] =
 * oldPosition`, the repo-wide convention — see `layout/rowOrder.ts`.
 *
 * Unlike `MoveRowsCommand` this carries no `start`/`count`/`target`: a
 * permutation touches every row at once and has no contiguous block to name.
 */
export class ReorderRowsCommand implements Command<"reorder-rows", ReorderRowsCommandParams> {
  readonly type = "reorder-rows";
  readonly params: ReorderRowsCommandParams;

  constructor(order: Int32Array) {
    // Copied for the same reason every other Command snapshots its payload:
    // a caller that keeps mutating `order` after construction must not be
    // able to change what an already-recorded history entry replays.
    this.params = Object.freeze({ order: Int32Array.from(order) });
    Object.freeze(this);
  }
}
