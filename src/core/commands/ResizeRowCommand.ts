import type { Command } from "../types";

export interface ResizeRowCommandParams {
  readonly row: number;
  readonly height: number;
}

export class ResizeRowCommand implements Command<"resize-row", ResizeRowCommandParams> {
  readonly type = "resize-row";
  readonly params: ResizeRowCommandParams;

  constructor(row: number, height: number) {
    this.params = Object.freeze({ row, height });
    Object.freeze(this);
  }
}
