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

/**
 * 选区上某个键的共同值；不一致时 `undefined`，下拉框据此显示空白而不是随便挑一格。
 *
 * **`fallback` 是"没设这个键时它实际等于什么"**，不是"显示什么"。给了它之后，
 * 一个没设 `fontSize` 的格子与一个显式设成 13 的格子会被当成相同——它们画出来本来
 * 就一模一样（`resolveDrawStyle` 对缺省键取的就是同一个默认值），工具栏没有理由把
 * 它们报成"不一致"。
 *
 * 这也是"整片都没设"与"选区内不一致"能分开的唯一办法：不给 fallback 时两者都落到
 * `undefined`，调用方再 `?? 默认值` 就会把真正的不一致也显示成默认值——字号框因此
 * 一直空白（`""` 匹配不到任何 option），字体框则把不一致的选区显示成"默认字体"。
 */
export function commonValue<K extends keyof CellStyle>(
  ctx: ToolbarContext,
  key: K,
  fallback?: CellStyle[K]
): CellStyle[K] | undefined {
  const { range } = ctx;
  const read = (row: number, column: number): CellStyle[K] | undefined =>
    ctx.styleOf(row, column)?.[key] ?? fallback;
  const first = read(range.startRow, range.startColumn);
  const rows = range.endRow - range.startRow + 1;
  const columns = range.endColumn - range.startColumn + 1;
  if (rows * columns > MAX_PROBED_CELLS) {
    return first;
  }
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      if (read(row, column) !== first) {
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
