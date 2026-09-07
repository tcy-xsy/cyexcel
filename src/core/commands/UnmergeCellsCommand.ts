import type { CellRange, Command } from "../types";

/**
 * `range` 圈定要拆散哪些合并区——落在 `range` 内、或与 `range` 相交的合并区全部
 * 拆散，语义与 Excel 一致（不要求精确等于某条已有合并区）。
 */
export interface UnmergeCellsCommandParams {
  readonly range: CellRange;
}

export class UnmergeCellsCommand implements Command<"unmerge-cells", UnmergeCellsCommandParams> {
  readonly type = "unmerge-cells";
  readonly params: UnmergeCellsCommandParams;

  constructor(range: CellRange) {
    this.params = Object.freeze({ range: Object.freeze({ ...range }) });
    Object.freeze(this);
  }
}
