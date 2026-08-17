import type { CellRange, Command } from "../types";

export interface ClearRangeCommandParams {
  readonly range: CellRange;
}

export class ClearRangeCommand implements Command<"clear-range", ClearRangeCommandParams> {
  readonly type = "clear-range";
  readonly params: ClearRangeCommandParams;

  constructor(range: CellRange) {
    this.params = Object.freeze({ range: Object.freeze({ ...range }) });
    Object.freeze(this);
  }
}
