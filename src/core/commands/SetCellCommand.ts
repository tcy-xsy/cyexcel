import type { TableStore } from "../data";
import type { Command, CommandContext, Patch } from "../types";

export class SetCellCommand implements Command {
  readonly type = "set-cell";

  readonly #store: TableStore;
  readonly #row: number;
  readonly #column: number;
  readonly #nextValue: unknown;
  #previousValue: unknown;

  constructor(store: TableStore, row: number, column: number, value: unknown) {
    this.#store = store;
    this.#row = row;
    this.#column = column;
    this.#nextValue = value;
  }

  execute(context: CommandContext): Patch | Patch[] {
    void context;
    this.#previousValue = this.#store.getCell(this.#row, this.#column);
    return [...this.#store.setCell(this.#row, this.#column, this.#nextValue).patches];
  }

  undo(context: CommandContext): Patch | Patch[] {
    void context;
    return [...this.#store.setCell(this.#row, this.#column, this.#previousValue).patches];
  }
}
