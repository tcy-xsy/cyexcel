import type { CellStyle } from "../style/types";
import type { ToolbarContext, ToolbarItem, ToolbarItemState } from "./types";

/**
 * 全部满足才算激活，与 Excel 一致——半粗的选区按钮不亮，再点一次是"全部变粗"。
 * 超过 `MAX_PROBED_CELLS` 格只看左上角一格：状态是提示，不值得为它扫一百万格。
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

/** 选区上某个键的共同值；不一致时 `undefined`，下拉框据此显示空白而不是随便挑一格。 */
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

/** 没声明的项给确定默认值，渲染层不必到处判 `undefined`。 */
export function resolveItemState(item: ToolbarItem, ctx: ToolbarContext): ToolbarItemState {
  return {
    activated: item.activated?.(ctx) ?? false,
    disabled: item.disabled?.(ctx) ?? false,
    value: item.value?.(ctx)
  };
}
