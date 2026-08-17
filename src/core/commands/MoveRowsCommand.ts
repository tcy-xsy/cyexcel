import type { Command } from "../types";

export interface MoveRowsCommandParams {
  readonly start: number;
  readonly count: number;
  readonly target: number;
}

/** `target` is the insertion index in the pre-move index space — same convention `ArrayDataSource.moveRows()`/`remapRowIndex()` use. */
export class MoveRowsCommand implements Command<"move-rows", MoveRowsCommandParams> {
  readonly type = "move-rows";
  readonly params: MoveRowsCommandParams;

  constructor(start: number, count: number, target: number) {
    this.params = Object.freeze({ start, count, target });
    Object.freeze(this);
  }
}
