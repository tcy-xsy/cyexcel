import type { CellRect, ViewportState } from "./LayoutEngine";

/**
 * `LayoutEngine.getCellRect()` returns absolute content-space coordinates
 * (a row's cumulative offset from row 0, independent of scroll — see
 * `LayoutEngine.getRowOffset()`). Anything drawn or positioned in
 * a viewport-sized, `position: sticky`-pinned layer (the content canvas, the
 * DOM cell editor overlay, and — from Step 8 — DOMCellManager's mounted
 * cells) needs that rect shifted by the current scroll offset first, or it
 * lands at the wrong on-screen position once scrollTop/scrollLeft are
 * nonzero. Shared here so every one of those call sites translates the same
 * way instead of each reimplementing (and potentially diverging on) the
 * same two-line subtraction.
 *
 * `frozen` (Step 9) opts either axis out of that subtraction — a frozen row
 * or column never moves as its own axis scrolls, since it isn't part of
 * the scrollable band. Defaults to `{}` (neither axis frozen), so existing
 * two-argument call sites (EditorManager, and ContentLayerRenderer before
 * Step 9's freeze support) are unaffected.
 */
export function toViewportRect(
  rect: CellRect,
  viewport: ViewportState,
  frozen: { readonly row?: boolean; readonly column?: boolean } = {}
): CellRect {
  return {
    x: frozen.column ? rect.x : rect.x - viewport.scrollLeft,
    y: frozen.row ? rect.y : rect.y - viewport.scrollTop,
    width: rect.width,
    height: rect.height
  };
}
