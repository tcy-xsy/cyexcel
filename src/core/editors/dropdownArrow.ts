import type { CellRect } from "../layout";
import { DROPDOWN_ARROW_ZONE } from "../render/theme";

/**
 * 下拉箭头的命中区：单元格右侧 `DROPDOWN_ARROW_ZONE` 像素宽、整格高。
 * 与画布上画箭头的位置（`ContentLayerRenderer#paintDropdownArrow`）共用同一个常量，
 * 画在哪、点哪里才有反应，不会两边各算各的。
 *
 * 命中区比画出来的三角形大得多，是为了好点：Univer 的 `isHit` 也是用记录下来的
 * 箭头矩形而不是三角形本身。
 */
export function isInDropdownArrowZone(rect: CellRect, point: { readonly x: number; readonly y: number }): boolean {
  const left = rect.x + Math.max(0, rect.width - DROPDOWN_ARROW_ZONE);
  return point.x >= left && point.x <= rect.x + rect.width && point.y >= rect.y && point.y <= rect.y + rect.height;
}
