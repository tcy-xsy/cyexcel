import type { CellValue } from "../sheet/types.js";
import type { CellRange, Command } from "../types.js";

export interface SetRangeCommandParams {
  readonly range: CellRange;
  readonly values: readonly (readonly (CellValue | null | undefined)[])[];
}

export class SetRangeCommand implements Command<"set-range", SetRangeCommandParams> {
  readonly type = "set-range";
  readonly params: SetRangeCommandParams;

  constructor(range: CellRange, values: readonly (readonly (CellValue | null | undefined)[])[]) {
    this.params = Object.freeze({
      range: Object.freeze({ ...range }),
      values: Object.freeze(values.map((row) => Object.freeze([...row])))
    });
    Object.freeze(this);
  }
}
