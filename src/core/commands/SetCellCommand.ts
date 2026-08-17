import type { Command } from "../types";

export interface SetCellCommandParams {
  readonly row: number;
  readonly column: number;
  readonly value: unknown;
}

export class SetCellCommand implements Command<"set-cell", SetCellCommandParams> {
  readonly type = "set-cell";
  readonly params: SetCellCommandParams;

  constructor(row: number, column: number, value: unknown) {
    this.params = Object.freeze({ row, column, value });
    Object.freeze(this);
  }
}
