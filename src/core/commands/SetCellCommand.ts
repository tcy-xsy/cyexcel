import type { CellValue } from "../sheet/types";
import type { Command } from "../types";

export interface SetCellCommandParams {
  readonly row: number;
  readonly column: number;
  readonly value: CellValue | null | undefined;
}

export class SetCellCommand implements Command<"set-cell", SetCellCommandParams> {
  readonly type = "set-cell";
  readonly params: SetCellCommandParams;

  constructor(row: number, column: number, value: CellValue | null | undefined) {
    this.params = Object.freeze({ row, column, value });
    Object.freeze(this);
  }
}
