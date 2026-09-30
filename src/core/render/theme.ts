/**
 * Shared color/font constants for the built-in canvas renderers
 * (ContentLayerRenderer/SelectionLayerRenderer/HeaderCanvasRenderer). A
 * plain constants module, not a configurable theme system — kept in one
 * place so the ~7 colors introduced together stay visually coherent
 * instead of drifting between files with independently-chosen hex values.
 */

export const GRID_LINE_COLOR = "#e2e5ea";
export const TEXT_COLOR = "#111827";
/** Auto-detected URL/email cells (see cells/linkDetection.ts) — the standard web hyperlink blue. */
export const LINK_COLOR = "#1a73e8";
/** 数据验证的"无效"角标（单元格右上角的小三角）。 */
export const INVALID_MARK_COLOR = "#fe4b4b";
export const INVALID_MARK_SIZE = 6;
/** 数据验证 list 单元格的下拉箭头：颜色、命中区宽度（单元格右侧这么宽都算点在箭头上）。 */
export const DROPDOWN_ARROW_COLOR = "#5f6368";
export const DROPDOWN_ARROW_ZONE = 18;
/** `custom` 显示模式的胶囊标签：没有选项颜色时的底色。 */
export const DROPDOWN_CAPSULE_COLOR = "#e8eaed";
export const CELL_PADDING = 4;
/**
 * 字号与字体分开导出，是为了让每格样式能覆盖其中一半（Step 24）。
 * `CELL_FONT` 由这两个常量拼出来，所以**没有样式的格子拼出的串与从前逐字相同**
 * ——Step 0-15 那些断言字体串的既有测试因此一个字都不用改。
 */
export const DEFAULT_FONT_SIZE = 13;
export const DEFAULT_FONT_FAMILY = '-apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif';
export const CELL_FONT = `${DEFAULT_FONT_SIZE}px ${DEFAULT_FONT_FAMILY}`;
/** Vertical space one wrapped line occupies — used only by wrap-enabled columns' multi-line paint/measurement; a single-line cell still centers on its own rect.height and never reads this. */
// 16px line box + 4px top/bottom padding exactly matches the default 24px
// row height, so opting a column into wrap does not make ordinary one-line
// rows grow before any content actually needs a second line.
export const CELL_LINE_HEIGHT = 16;
export const HEADER_FONT = `600 ${CELL_FONT}`;
export const HEADER_BACKGROUND_COLOR = "#f3f4f6";
export const HEADER_HOVER_BACKGROUND_COLOR = "#e5e7eb";
export const HEADER_BORDER_COLOR = "#d1d5db";
/**
 * The two sort carets are always both drawn; which one is *active* is said in
 * colour, not by swapping the mark — the same thing antd's Table does
 * (`.ant-table-column-sorter-up/down`, `.active` → `colorPrimary`). Keeping the
 * inactive caret visible is what tells a user the other direction exists.
 */
export const HEADER_SORT_INDICATOR_COLOR = "#bfbfbf";
export const FROZEN_SEPARATOR_COLOR = "#9ca3af";
export const FROZEN_SEPARATOR_THICKNESS = 2;
export const SELECTION_STROKE_COLOR = "#3b82f6";
export const SELECTION_STROKE_WIDTH = 1.5;
export const SELECTION_HANDLE_SIZE = 6;
/** Non-active regions in a multi-region selection (Ctrl/Cmd+click) — same stroke color, just faded, keeping with this project's "thin and light, not bold" chrome convention. */
export const SELECTION_INACTIVE_ALPHA = 0.5;
/** The row drag-to-reorder insertion indicator — reuses the selection blue rather than introducing a new hue, just thicker so it reads as a distinct "drop here" line rather than a selection edge. */
export const ROW_REORDER_INDICATOR_COLOR = SELECTION_STROKE_COLOR;
/** The caret for the direction currently in effect — see HEADER_SORT_INDICATOR_COLOR. */
export const HEADER_SORT_INDICATOR_ACTIVE_COLOR = SELECTION_STROKE_COLOR;
export const ROW_REORDER_INDICATOR_THICKNESS = 3;
