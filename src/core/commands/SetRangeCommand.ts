import type { CellRange, Command } from "../types";

export interface SetRangeCommandParams {
  readonly range: CellRange;
  readonly values: readonly (readonly unknown[])[];
}

export class SetRangeCommand implements Command<"set-range", SetRangeCommandParams> {
  readonly type = "set-range";
  readonly params: SetRangeCommandParams;

  constructor(range: CellRange, values: readonly (readonly unknown[])[]) {
    this.params = Object.freeze({
      range: Object.freeze({ ...range }),
      values: Object.freeze(values.map((row) => Object.freeze([...row])))
    });
    Object.freeze(this);
  }
}
