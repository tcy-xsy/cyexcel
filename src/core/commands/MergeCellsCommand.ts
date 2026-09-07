import type { CellRange, Command } from "../types";

export interface MergeCellsCommandParams {
  readonly range: CellRange;
}

export class MergeCellsCommand implements Command<"merge-cells", MergeCellsCommandParams> {
  readonly type = "merge-cells";
  readonly params: MergeCellsCommandParams;

  constructor(range: CellRange) {
    this.params = Object.freeze({ range: Object.freeze({ ...range }) });
    Object.freeze(this);
  }
}
