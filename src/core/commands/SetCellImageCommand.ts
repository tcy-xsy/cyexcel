import type { Command } from "../types";

export interface SetCellImageCommandParams {
  readonly row: number;
  readonly column: number;
  readonly url: string | undefined;
}

export class SetCellImageCommand implements Command<"set-cell-image", SetCellImageCommandParams> {
  readonly type = "set-cell-image";
  readonly params: SetCellImageCommandParams;

  constructor(row: number, column: number, url: string | undefined) {
    this.params = Object.freeze({ row, column, url });
    Object.freeze(this);
  }
}
