import type { Command } from "../types";

export interface ShowColumnCommandParams {
  readonly column: number;
}

export class ShowColumnCommand implements Command<"show-column", ShowColumnCommandParams> {
  readonly type = "show-column";
  readonly params: ShowColumnCommandParams;

  constructor(column: number) {
    this.params = Object.freeze({ column });
    Object.freeze(this);
  }
}
