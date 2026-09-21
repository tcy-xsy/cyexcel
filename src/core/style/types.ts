/**
 * 单元格样式的字段集（ADR 0004 §3.4）。
 *
 * 只收渲染层画得出来、且不牵动自动行高链路的键。文字旋转、上下标、每格换行策略
 * 有意不收，字段也不预留——它们要连同 Step 21 的行高测量一起设计。
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

/** 固定顺序，intern key 的规范化依赖它（见 `styleKey`）。加字段请追加到末尾。 */
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
