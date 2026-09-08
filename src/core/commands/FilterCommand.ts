import type { Command } from "../types";

export interface FilterCommandParams {
  readonly filteredRows: ReadonlySet<number>;
}

/**
 * Whole-set replacement of `filteredRows` (ADR 0003 §3.2/§3.3) — never an
 * incremental toggle, so undoing it is just "write back the previous set".
 * Always run with `{ history: "skip" }` (same mechanism as auto row height):
 * it commits to live state but never enters the undo/redo stacks and never
 * clears an existing redo entry.
 */
export class FilterCommand implements Command<"filter-rows", FilterCommandParams> {
  readonly type = "filter-rows";
  readonly params: FilterCommandParams;

  constructor(filteredRows: ReadonlySet<number>) {
    this.params = Object.freeze({ filteredRows: new Set(filteredRows) });
    Object.freeze(this);
  }
}
