import type { CellRect } from "../layout";
import { CELL_LINE_HEIGHT, DROPDOWN_ARROW_ZONE } from "./theme";

/**
 * list 单元格下拉的几何（数据验证设计稿 §8）：画在哪、点哪里才有反应，全在这一个文件里，
 * 绘制（`ContentLayerRenderer`）和命中测试（`CyExcelCore`）共用，不会各算各的。纯函数，
 * 不碰 canvas——文字宽度由调用方量好传进来。
 *
 * 常量取自 Univer 的 `DropdownWidget`：胶囊距格子边的外边距 `MARGIN_H` / `MARGIN_V`、
 * 胶囊内边距 `PADDING_H`、箭头占位 `ICON_PLACE`。
 */

/** 显示模式，同 `DataValidationRenderMode`；这里独立声明，渲染层不依赖 `sheet/`。 */
export type DropdownMode = "text" | "arrow" | "custom";

export interface CellDropdown {
  readonly mode: DropdownMode;
  /** 当前值对应的颜色：`custom` 是胶囊底色，其余模式是单元格底色。 */
  readonly color?: string;
}

export const CAPSULE_MARGIN_H = 6;
export const CAPSULE_MARGIN_V = 3;
export const CAPSULE_PADDING_H = 4;
export const CAPSULE_ICON_PLACE = 14;
export const CAPSULE_RADIUS = 8;
/** 胶囊的高度上限：一行字加上下各 1px。格子再高胶囊也不拉长，垂直居中。 */
export const CAPSULE_HEIGHT = CELL_LINE_HEIGHT + 2;

type Point = { readonly x: number; readonly y: number };

function contains(rect: CellRect, point: Point): boolean {
  return (
    point.x >= rect.x && point.x <= rect.x + rect.width && point.y >= rect.y && point.y <= rect.y + rect.height
  );
}

/** `arrow` 模式：单元格右侧 `DROPDOWN_ARROW_ZONE` 像素宽、整格高。格子更窄时整格都算。 */
export function isInDropdownArrowZone(rect: CellRect, point: Point): boolean {
  const left = rect.x + Math.max(0, rect.width - DROPDOWN_ARROW_ZONE);
  return contains({ x: left, y: rect.y, width: rect.x + rect.width - left, height: rect.height }, point);
}

export interface CapsuleGeometry {
  /** 胶囊本身的矩形。 */
  readonly capsule: CellRect;
  /** 胶囊里文字的起点 x 与最大宽度（过长的文字裁剪到这里）。 */
  readonly textX: number;
  readonly textMaxWidth: number;
  /** 箭头中心。 */
  readonly arrowCenter: Point;
}

/**
 * `custom` 模式的胶囊：贴着格子左边、垂直居中，宽度是"内边距 + 文字 + 箭头占位 + 内边距"，
 * 最宽撑满格子（扣掉两侧外边距）。格子放不下（太窄或太矮）返回 `null`，调用方退回
 * `arrow` 模式——胶囊挤不下时硬画只会画出一个变形的东西。
 */
export function computeCapsule(rect: CellRect, textWidth: number): CapsuleGeometry | null {
  const maxWidth = rect.width - CAPSULE_MARGIN_H * 2;
  const height = Math.min(rect.height - CAPSULE_MARGIN_V * 2, CAPSULE_HEIGHT);
  if (height <= 0 || maxWidth < CAPSULE_ICON_PLACE + CAPSULE_PADDING_H * 2) {
    return null;
  }
  const width = Math.min(maxWidth, CAPSULE_PADDING_H + Math.max(0, textWidth) + CAPSULE_ICON_PLACE + CAPSULE_PADDING_H);
  const x = rect.x + CAPSULE_MARGIN_H;
  const y = rect.y + (rect.height - height) / 2;
  return {
    capsule: { x, y, width, height },
    textX: x + CAPSULE_PADDING_H,
    textMaxWidth: Math.max(0, width - CAPSULE_PADDING_H * 2 - CAPSULE_ICON_PLACE),
    arrowCenter: { x: x + width - CAPSULE_PADDING_H - CAPSULE_ICON_PLACE / 2, y: y + height / 2 }
  };
}

/**
 * 点 `point` 落没落在这一格的下拉可点区域里。`text` 模式没有箭头，恒为 `false`；`arrow`
 * 看右侧命中区；`custom` 看整个胶囊（比只认箭头那一小块好点），胶囊挤不下时同 `arrow`。
 */
export function isInDropdownHitArea(
  dropdown: CellDropdown,
  rect: CellRect,
  textWidth: number,
  point: Point
): boolean {
  switch (dropdown.mode) {
    case "text":
      return false;
    case "arrow":
      return isInDropdownArrowZone(rect, point);
    case "custom": {
      const geometry = computeCapsule(rect, textWidth);
      return geometry === null ? isInDropdownArrowZone(rect, point) : contains(geometry.capsule, point);
    }
  }
}
