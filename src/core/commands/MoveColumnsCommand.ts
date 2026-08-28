import type { Command } from "../types";

export interface MoveColumnsCommandParams {
  readonly start: number;
  readonly count: number;
  readonly target: number;
}

/** `target` is the insertion index in the pre-move index space — same convention `SheetModel.moveColumns()`/`remapColumnIndex()` use. */
export class MoveColumnsCommand implements Command<"move-columns", MoveColumnsCommandParams> {
  readonly type = "move-columns";
  readonly params: MoveColumnsCommandParams;

  constructor(start: number, count: number, target: number) {
    this.params = Object.freeze({ start, count, target });
    Object.freeze(this);
  }
}
