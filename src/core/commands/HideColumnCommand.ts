import type { Command } from "../types";

export interface HideColumnCommandParams {
  readonly column: number;
}

export class HideColumnCommand implements Command<"hide-column", HideColumnCommandParams> {
  readonly type = "hide-column";
  readonly params: HideColumnCommandParams;

  constructor(column: number) {
    this.params = Object.freeze({ column });
    Object.freeze(this);
  }
}
