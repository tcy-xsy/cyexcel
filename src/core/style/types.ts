/**
 * 单元格样式的字段集（ADR 0004 §3.4）。
 *
 * # 为什么是这 13 个键
 *
 * 取的是"渲染层画得出来、且不牵动自动行高链路"的那些。被有意排除的三项——
 * 文字旋转、上下标、每格换行策略——共同点是画出来之后行高得跟着重算，而重算是
 * Step 21 `measurement/` 的地盘：要做就得连同那条链路一起设计，不是加个字段的事。
 * 所以**字段也不预留**，免得出现"设得了但没效果"的 API。
 *
 * # 为什么是全拼而不是 Univer 的两字母缩写
 *
 * Univer 的 `IStyleData` 用 `bl`/`it`/`cl`/`n` 这样的缩写，省的是 workbook 级
 * `styles` 表里每个样式对象一份的体积。但样式本来就已经 intern 去重过了——一万个
 * 加粗格子只存一条——省下的量抵不上可读性。xlsx 那层字段名映射 M3 反正要写，
 * 多映射几个全拼字段不增加工作量。
 */

/** 一条边。`style` 暂时只有实线，虚线/双线留到有人真要时再加。 */
export interface BorderSide {
  readonly color: string;
  /** 逻辑像素，默认 1 */
  readonly width?: number;
  readonly style?: "solid";
}

export interface CellBorder {
  readonly top?: BorderSide;
  readonly right?: BorderSide;
  readonly bottom?: BorderSide;
  readonly left?: BorderSide;
}

export interface CellStyle {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly strike?: boolean;
  readonly fontFamily?: string;
  /** px，与 `CELL_FONT` 同单位 */
  readonly fontSize?: number;
  /** CSS 颜色串，画文字用 */
  readonly color?: string;
  /** CSS 颜色串，画背景用 */
  readonly background?: string;
  readonly align?: "left" | "center" | "right";
  readonly verticalAlign?: "top" | "middle" | "bottom";
  /** 逻辑像素，覆盖 `CELL_PADDING` */
  readonly padding?: number;
  readonly border?: CellBorder;
  /** Excel 数字格式 pattern 串，交给 numfmt 解析 */
  readonly numFmt?: string;
}

/**
 * 字段的**固定顺序**，intern key 的规范化依赖它（见 `styleKey`）。
 *
 * 往里加字段必须同时想清楚两件事：加在哪个位置（追加到末尾最安全，插在中间会让
 * 所有既有快照的 key 变化——虽然 key 只是内存里的去重索引、不进快照，但测试里的
 * 期望字符串会全部要改），以及渲染层是不是真的画得出来。
 */
export const STYLE_KEYS = [
  "bold",
  "italic",
  "underline",
  "strike",
  "fontFamily",
  "fontSize",
  "color",
  "background",
  "align",
  "verticalAlign",
  "padding",
  "border",
  "numFmt"
] as const;

export type CellStyleKey = (typeof STYLE_KEYS)[number];
