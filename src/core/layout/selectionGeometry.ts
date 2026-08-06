import type { CellRange } from "../types";
import type { CellRect, LayoutEngine, ViewportState, IndexRange } from "./LayoutEngine";
import { splitByFrozenBoundary } from "./frozenRegions";
import { toViewportRect } from "./viewportRect";

interface Quadrant {
  readonly rows: IndexRange;
  readonly columns: IndexRange;
  readonly frozenRow: boolean;
  readonly frozenColumn: boolean;
}

function nonEmptyQuadrants(quadrants: readonly Quadrant[]): Quadrant[] {
  return quadrants.filter((quadrant) => quadrant.rows.count > 0 && quadrant.columns.count > 0);
}

function translatedCellRect(
  layout: LayoutEngine,
  viewport: ViewportState,
  row: number,
  column: number,
  frozenRow: boolean,
  frozenColumn: boolean
): CellRect {
  return toViewportRect(layout.getCellRect(row, column), viewport, { row: frozenRow, column: frozenColumn });
}

function quadrantRect(layout: LayoutEngine, viewport: ViewportState, quadrant: Quadrant): CellRect {
  const lastRow = quadrant.rows.start + quadrant.rows.count - 1;
  const lastColumn = quadrant.columns.start + quadrant.columns.count - 1;
  const topLeft = translatedCellRect(
    layout,
    viewport,
    quadrant.rows.start,
    quadrant.columns.start,
    quadrant.frozenRow,
    quadrant.frozenColumn
  );
  const bottomRight = translatedCellRect(
    layout,
    viewport,
    lastRow,
    lastColumn,
    quadrant.frozenRow,
    quadrant.frozenColumn
  );
  return {
    x: topLeft.x,
    y: topLeft.y,
    width: bottomRight.x + bottomRight.width - topLeft.x,
    height: bottomRight.y + bottomRight.height - topLeft.y
  };
}

/**
 * On-screen rect of the selection's bottom-right-most non-empty quadrant —
 * the same up-to-four-quadrant frozen-aware split SelectionLayerRenderer
 * paints the selection box with (doc/architecture.md §6.1). Shared so the
 * fill handle's painted position (SelectionLayerRenderer) and its hit-testing
 * (FillHandleController) can never disagree about where it is.
 */
export function getSelectionBottomRightRect(
  range: CellRange,
  layout: LayoutEngine,
  viewport: ViewportState
): CellRect {
  const rowIndexRange: IndexRange = { start: range.startRow, count: range.endRow - range.startRow + 1 };
  const columnIndexRange: IndexRange = {
    start: range.startColumn,
    count: range.endColumn - range.startColumn + 1
  };

  const rowSplit = splitByFrozenBoundary(rowIndexRange, layout.frozenRowCount);
  const columnSplit = splitByFrozenBoundary(columnIndexRange, layout.frozenColumnCount);

  const quadrants = nonEmptyQuadrants([
    { rows: rowSplit.frozen, columns: columnSplit.frozen, frozenRow: true, frozenColumn: true },
    { rows: rowSplit.frozen, columns: columnSplit.scrollable, frozenRow: true, frozenColumn: false },
    { rows: rowSplit.scrollable, columns: columnSplit.frozen, frozenRow: false, frozenColumn: true },
    { rows: rowSplit.scrollable, columns: columnSplit.scrollable, frozenRow: false, frozenColumn: false }
  ]);

  const lastQuadrant = quadrants[quadrants.length - 1];
  if (!lastQuadrant) {
    throw new Error("getSelectionBottomRightRect: range has no non-empty quadrant");
  }
  return quadrantRect(layout, viewport, lastQuadrant);
}
