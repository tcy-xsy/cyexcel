import {
  CELL_FONT,
  CELL_PADDING,
  DEFAULT_FONT_FAMILY,
  DEFAULT_FONT_SIZE,
  TEXT_COLOR
} from "../render/theme";
import type { BorderSide, CellStyle } from "./types";

/**
 * 一格绘制真正要用的参数。字段都是"画之前直接赋给 ctx"的形态——`font` 已经是拼好的
 * 串，`padding` 已经把缺省补上——所以绘制路径上没有任何条件判断与字符串拼接。
 */
export interface DrawStyle {
  readonly font: string;
  readonly color: string;
  readonly background: string | undefined;
  readonly align: "left" | "center" | "right";
  readonly verticalAlign: "top" | "middle" | "bottom";
  readonly padding: number;
  readonly underline: boolean;
  readonly strike: boolean;
  readonly border: CellStyle["border"];
}

/** 没有任何样式的格子用这一份——与 Step 24 之前的绘制行为逐字相同。 */
export const DEFAULT_DRAW_STYLE: DrawStyle = Object.freeze({
  font: CELL_FONT,
  color: TEXT_COLOR,
  background: undefined,
  align: "left",
  verticalAlign: "middle",
  padding: CELL_PADDING,
  underline: false,
  strike: false,
  border: undefined
});

/**
 * 缓存派生结果。key 是样式**对象本身**：`StyleRegistry` intern 过之后，同一个样式
 * 在全表共享同一个对象引用，所以一万个加粗格子滚过视口只会算一次。
 *
 * 用 `WeakMap` 而不是 `Map`：样式被 undo 掉、再没人引用时，这份派生跟着一起走，
 * 不需要谁来记得清理。
 */
const cache = new WeakMap<CellStyle, DrawStyle>();

/**
 * 把样式对象翻成绘制参数。
 *
 * 每格滚过视口都要调一次，所以这里的开销直接摊在滚动帧上——真正的计算只在
 * 缓存未命中时发生一次。
 */
export function resolveDrawStyle(style: CellStyle | undefined): DrawStyle {
  if (style === undefined) {
    return DEFAULT_DRAW_STYLE;
  }
  const cached = cache.get(style);
  if (cached !== undefined) {
    return cached;
  }
  const resolved = Object.freeze({
    font: composeFont(style),
    color: style.color ?? TEXT_COLOR,
    background: style.background,
    align: style.align ?? "left",
    verticalAlign: style.verticalAlign ?? "middle",
    padding: style.padding ?? CELL_PADDING,
    underline: style.underline ?? false,
    strike: style.strike ?? false,
    border: style.border
  });
  cache.set(style, resolved);
  return resolved;
}

/**
 * CSS font 简写的顺序是 `style weight size family`，不能随便换位——写成
 * `600 italic ...` 整条声明会被丢掉，表现为"这一格突然变回默认字体"。
 */
function composeFont(style: CellStyle): string {
  const size = style.fontSize ?? DEFAULT_FONT_SIZE;
  const family = style.fontFamily ?? DEFAULT_FONT_FAMILY;
  const parts: string[] = [];
  if (style.italic === true) {
    parts.push("italic");
  }
  if (style.bold === true) {
    parts.push("600");
  }
  parts.push(`${size}px ${family}`);
  return parts.join(" ");
}

/** 一条边的最终画法，`undefined` 表示这条边不画。 */
export function borderSideOf(side: BorderSide | undefined): { color: string; width: number } | undefined {
  return side === undefined ? undefined : { color: side.color, width: side.width ?? 1 };
}
