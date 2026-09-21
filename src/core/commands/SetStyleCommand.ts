import type { Command } from "../types";
import type { CellRange } from "../types";
import type { CellStylePatch } from "../style/styleKey";

export interface SetStyleCommandParams {
  readonly range: CellRange;
  /** 合并语义：省略的键保留，`null` 的键清除。 */
  readonly patch: CellStylePatch;
}

/**
 * 给一片区域改样式。单格与区域是同一条命令——区域退化成 1×1 时产出的补丁自然是
 * `cell-style` 而不是 `range-style`，不需要两个命令类型。
 *
 * 整片区域是**一次事务、一次 History 记录**（Step 24 退出条件第 6 条）：
 * `CellStyleMutation` 是一个参与者，内部把每一格的写入收在一起 commit。
 */
export class SetStyleCommand implements Command<"set-style", SetStyleCommandParams> {
  readonly type = "set-style";
  readonly params: SetStyleCommandParams;

  constructor(range: CellRange, patch: CellStylePatch) {
    this.params = Object.freeze({ range: Object.freeze({ ...range }), patch: Object.freeze({ ...patch }) });
    Object.freeze(this);
  }
}
