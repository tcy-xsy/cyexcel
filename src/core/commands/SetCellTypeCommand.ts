import type { CellRange, Command } from "../types";

export interface SetCellTypeCommandParams {
  readonly range: CellRange;
  /** 类型名（注册表里的），`null` 表示回到列默认。 */
  readonly typeName: string | null;
}

/** 给一片区域设格级语义类型，覆盖列默认。 */
export class SetCellTypeCommand implements Command<"set-cell-type", SetCellTypeCommandParams> {
  readonly type = "set-cell-type";
  readonly params: SetCellTypeCommandParams;

  constructor(range: CellRange, typeName: string | null) {
    this.params = Object.freeze({ range: Object.freeze({ ...range }), typeName });
    Object.freeze(this);
  }
}
