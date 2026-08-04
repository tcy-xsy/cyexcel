import type { TableStore } from "../data";
import type { CellRange, Command, CommandContext, Patch } from "../types";
import { captureRange } from "./captureRange";

export class SetRangeCommand implements Command {
  readonly type = "set-range";

  readonly #store: TableStore;
  readonly #range: CellRange;
  readonly #nextValues: readonly (readonly unknown[])[];
  #previousValues: unknown[][] = [];

  constructor(
    store: TableStore,
    range: CellRange,
    values: readonly (readonly unknown[])[]
  ) {
    this.#store = store;
    this.#range = range;
    this.#nextValues = values;
  }

  execute(context: CommandContext): Patch | Patch[] {
    void context;
    this.#previousValues = captureRange(this.#store, this.#range);
    return [...this.#store.setRange(this.#range, this.#nextValues).patches];
  }

  undo(context: CommandContext): Patch | Patch[] {
    void context;
    return [...this.#store.setRange(this.#range, this.#previousValues).patches];
  }
}
