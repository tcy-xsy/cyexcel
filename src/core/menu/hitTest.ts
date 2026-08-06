import type { SelectionSnapshot } from "../interaction";
import type { LayoutEngine, ViewportState } from "../layout";
import type { ContextMenuContext } from "./types";

/** Same shape as interaction's HeaderLayout — kept structural here so this module doesn't need to import from "../interaction" just for a two-field type. */
export interface HeaderLayoutLike {
  readonly rowHeaderWidth: number;
  readonly columnHeaderHeight: number;
}

/**
 * Classifies a point (already relative to the content canvas's own origin,
 * same convention PointerController uses) into the five
 * doc/customization-and-hooks.md §3.6 targets. Reuses
 * `LayoutEngine.getCellAtPoint()`'s existing "past the data extent returns
 * null" semantics for "blank" instead of re-deriving that boundary here.
 */
export function classifyContextMenuTarget(
  point: { readonly x: number; readonly y: number },
  viewport: ViewportState,
  headerLayout: HeaderLayoutLike,
  layout: LayoutEngine,
  selection: SelectionSnapshot
): ContextMenuContext {
  const inRowHeader = point.x < headerLayout.rowHeaderWidth;
  const inColumnHeader = point.y < headerLayout.columnHeaderHeight;

  if (inRowHeader && inColumnHeader) {
    return { target: "corner", selection, point };
  }

  if (inRowHeader) {
    // x forced to 0 (always a valid column) — only the resolved row matters here.
    const cell = layout.getCellAtPoint({ x: 0, y: point.y - headerLayout.columnHeaderHeight }, viewport);
    return { target: "row-header", selection, point, ...(cell ? { row: cell.row } : {}) };
  }

  if (inColumnHeader) {
    // y forced to 0 (always a valid row) — only the resolved column matters here.
    const cell = layout.getCellAtPoint({ x: point.x - headerLayout.rowHeaderWidth, y: 0 }, viewport);
    return { target: "column-header", selection, point, ...(cell ? { column: cell.column } : {}) };
  }

  const gridPoint = { x: point.x - headerLayout.rowHeaderWidth, y: point.y - headerLayout.columnHeaderHeight };
  const cell = layout.getCellAtPoint(gridPoint, viewport);
  if (!cell) {
    return { target: "blank", selection, point };
  }
  return { target: "cell", selection, point, row: cell.row, column: cell.column };
}
