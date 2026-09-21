import type { CellStyle } from "../style/types";
import type { ToolbarContext, ToolbarItem, ToolbarItemState } from "./types";

/**
 * 遍历选区里的每一格，判断某个样式键是不是**全都**满足。
 *
 * "全部都满足才算激活"与 Excel 一致：选区里有的格子粗、有的不粗时，加粗按钮不高亮
 * （再点一次是"全部变粗"，而不是"全部反转"）。若改成"有一个就算"，那按钮按下去到底
 * 该变粗还是变不粗就没有确定答案了。
 *
 * 大选区上这个遍历不便宜，所以**有上限**：超过 `MAX_PROBED_CELLS` 格就只看左上角那
 * 一格。工具栏状态是给人看的提示，不是不变量；为了它去扫一百万格，代价和收益不成
 * 比例。
 */
export const MAX_PROBED_CELLS = 2000;

function everyCell(
  ctx: ToolbarContext,
  predicate: (style: CellStyle | undefined) => boolean
): boolean {
  const { range } = ctx;
  const rows = range.endRow - range.startRow + 1;
  const columns = range.endColumn - range.startColumn + 1;
  if (rows * columns > MAX_PROBED_CELLS) {
    return predicate(ctx.styleOf(range.startRow, range.startColumn));
  }
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      if (!predicate(ctx.styleOf(row, column))) {
        return false;
      }
    }
  }
  return true;
}

/** 某个布尔样式键在整片选区上是否都为真。 */
export function everyCellHas(ctx: ToolbarContext, key: "bold" | "italic" | "underline" | "strike"): boolean {
  return everyCell(ctx, (style) => style?.[key] === true);
}

/**
 * 整片选区上某个键的共同值；不一致（或没设）时是 `undefined`。
 *
 * 字号下拉、对齐按钮的当前值都靠它——三格字号不同时，下拉框应当显示空白而不是随便
 * 挑一个格子的值糊弄。
 */
export function commonValue<K extends keyof CellStyle>(
  ctx: ToolbarContext,
  key: K
): CellStyle[K] | undefined {
  const { range } = ctx;
  const first = ctx.styleOf(range.startRow, range.startColumn)?.[key];
  const rows = range.endRow - range.startRow + 1;
  const columns = range.endColumn - range.startColumn + 1;
  if (rows * columns > MAX_PROBED_CELLS) {
    return first;
  }
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      if (ctx.styleOf(row, column)?.[key] !== first) {
        return undefined;
      }
    }
  }
  return first;
}

/**
 * 算出一项当前的状态。
 *
 * 项没有声明 `activated`/`disabled`/`value` 时给确定的默认值（不高亮、不禁用、
 * 无值），所以渲染层永远拿到一个完整对象，不必到处判 `undefined`。
 */
export function resolveItemState(item: ToolbarItem, ctx: ToolbarContext): ToolbarItemState {
  return {
    activated: item.activated?.(ctx) ?? false,
    disabled: item.disabled?.(ctx) ?? false,
    value: item.value?.(ctx)
  };
}
