import type { CellRange, Command } from "../types";

export interface ClearDataValidationCommandParams {
  readonly ranges: readonly CellRange[];
}

/** 清掉 `ranges` 覆盖到的数据验证规则部分；区域被挖空的规则整条消失。 */
export class ClearDataValidationCommand implements Command<"clear-data-validation", ClearDataValidationCommandParams> {
  readonly type = "clear-data-validation";
  readonly params: ClearDataValidationCommandParams;

  constructor(ranges: readonly CellRange[]) {
    this.params = Object.freeze({ ranges: Object.freeze(ranges.map((range) => Object.freeze({ ...range }))) });
    Object.freeze(this);
  }
}
