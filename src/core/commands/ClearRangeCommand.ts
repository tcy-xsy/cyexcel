import type { TableStore } from "../data";
import type { CellRange, Command, CommandContext, Patch } from "../types";
import { captureRange } from "./captureRange";

function createEmptyGrid(range: CellRange): undefined[][] {
  const rowCount = range.endRow - range.startRow + 1;
  const columnCount = range.endColumn - range.startColumn + 1;
  return Array.from({ length: rowCount }, () =>
    Array.from({ length: columnCount }, () => undefined)
  );
}

export class ClearRangeCommand implements Command {
  readonly type = "clear-range";

  readonly #store: TableStore;
  readonly #range: CellRange;
  #previousValues: unknown[][] = [];

  constructor(store: TableStore, range: CellRange) {
    this.#store = store;
    this.#range = range;
  }

  execute(context: CommandContext): Patch | Patch[] {
    void context;
    this.#previousValues = captureRange(this.#store, this.#range);
    return [...this.#store.setRange(this.#range, createEmptyGrid(this.#range)).patches];
  }

  undo(context: CommandContext): Patch | Patch[] {
    void context;
    return [...this.#store.setRange(this.#range, this.#previousValues).patches];
  }
}
