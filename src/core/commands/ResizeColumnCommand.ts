import type { Command } from "../types";

export interface ResizeColumnCommandParams {
  readonly column: number;
  readonly width: number;
}

export class ResizeColumnCommand implements Command<"resize-column", ResizeColumnCommandParams> {
  readonly type = "resize-column";
  readonly params: ResizeColumnCommandParams;

  constructor(column: number, width: number) {
    this.params = Object.freeze({ column, width });
    Object.freeze(this);
  }
}
