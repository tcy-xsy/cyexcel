/**
 * Shared color/font constants for the built-in canvas renderers
 * (ContentLayerRenderer/SelectionLayerRenderer/HeaderCanvasRenderer). A
 * plain constants module, not a configurable theme system — kept in one
 * place so the ~7 colors introduced together stay visually coherent
 * instead of drifting between files with independently-chosen hex values.
 */

export const GRID_LINE_COLOR = "#e2e5ea";
export const TEXT_COLOR = "#111827";
export const CELL_PADDING = 4;
export const CELL_FONT = '13px -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif';
export const HEADER_FONT = `600 ${CELL_FONT}`;
export const HEADER_BACKGROUND_COLOR = "#f3f4f6";
export const HEADER_HOVER_BACKGROUND_COLOR = "#e5e7eb";
export const HEADER_BORDER_COLOR = "#d1d5db";
export const FROZEN_SEPARATOR_COLOR = "#9ca3af";
export const FROZEN_SEPARATOR_THICKNESS = 2;
export const SELECTION_STROKE_COLOR = "#3b82f6";
export const SELECTION_STROKE_WIDTH = 1.5;
export const SELECTION_HANDLE_SIZE = 6;
