import type { CellRect, IndexRange, LayoutEngine, ViewportState } from "../layout";
import { getSelectionBottomRightRect, splitByFrozenBoundary, toViewportRect } from "../layout";
import type { CellRange } from "../types";
import type { SelectionSnapshot } from "../interaction";
import type { Canvas2DLike } from "./Canvas2DLike";
import { SELECTION_HANDLE_SIZE, SELECTION_STROKE_COLOR, SELECTION_STROKE_WIDTH } from "./theme";

export interface SelectionLayerRendererOptions {
  readonly ctx: Canvas2DLike;
  readonly layout: LayoutEngine;
  readonly getSelection: () => SelectionSnapshot;
  /**
   * Whether the box should actually be drawn this frame — defaults to
   * always true. The canvas is still cleared either way, so a box hidden
   * this way (rather than by the caller simply not calling
   * `paintSelection()`) correctly disappears instead of leaving a stale
   * stroke. Intended for hiding the box while the active cell is being
   * edited: the edit input's own border is the visual indicator then, and
   * without this the two were drawn on top of each other (a highlighted
   * cell that's also mid-edit showed two overlapping, misaligned borders).
   */
  readonly isVisible?: () => boolean;
  /**
   * The range a fill-handle drag would extend the selection to, or null when
   * no fill drag is in progress. Painted as a dashed preview rect on top of
   * the solid selection box. Not frozen-aware — fill dragging only extends
   * down/right into the scrollable band (see FillHandleController), so the
   * preview never needs the quadrant split the selection box itself does.
   */
  readonly getFillPreviewRange?: () => CellRange | null;
}

interface Quadrant {
  readonly rows: IndexRange;
  readonly columns: IndexRange;
  readonly frozenRow: boolean;
  readonly frozenColumn: boolean;
}

function nonEmptyQuadrants(quadrants: readonly Quadrant[]): Quadrant[] {
  return quadrants.filter((quadrant) => quadrant.rows.count > 0 && quadrant.columns.count > 0);
}

/**
 * Draws the current selection's bounding box (SelectionModel.snapshot.range
 * already collapses cell/row/column/all selection kinds into one
 * CellRange — see SelectionModel — so this never branches on `kind`).
 * Frozen-aware via the same up-to-four-quadrant split ContentLayerRenderer
 * uses (doc/architecture.md §6.1): a selection straddling the frozen
 * boundary is drawn as up to 4 separate stroked rects, each translated per
 * its own frozen-row/frozen-column status.
 */
export class SelectionLayerRenderer {
  readonly #ctx: Canvas2DLike;
  readonly #layout: LayoutEngine;
  readonly #getSelection: () => SelectionSnapshot;
  readonly #isVisible: () => boolean;
  readonly #getFillPreviewRange: () => CellRange | null;
  #width = 0;
  #height = 0;
  #viewport: ViewportState = { scrollTop: 0, scrollLeft: 0, width: 0, height: 0 };

  constructor(options: SelectionLayerRendererOptions) {
    this.#ctx = options.ctx;
    this.#layout = options.layout;
    this.#getSelection = options.getSelection;
    this.#isVisible = options.isVisible ?? (() => true);
    this.#getFillPreviewRange = options.getFillPreviewRange ?? (() => null);
  }

  setSize(width: number, height: number): void {
    this.#width = width;
    this.#height = height;
  }

  setViewport(viewport: ViewportState): void {
    this.#viewport = viewport;
  }

  paintSelection(): void {
    this.#ctx.clearRect(0, 0, this.#width, this.#height);
    if (!this.#isVisible()) {
      return;
    }

    const { range } = this.#getSelection();
    const rowIndexRange: IndexRange = { start: range.startRow, count: range.endRow - range.startRow + 1 };
    const columnIndexRange: IndexRange = {
      start: range.startColumn,
      count: range.endColumn - range.startColumn + 1
    };

    const rowSplit = splitByFrozenBoundary(rowIndexRange, this.#layout.frozenRowCount);
    const columnSplit = splitByFrozenBoundary(columnIndexRange, this.#layout.frozenColumnCount);

    const quadrants = nonEmptyQuadrants([
      { rows: rowSplit.frozen, columns: columnSplit.frozen, frozenRow: true, frozenColumn: true },
      { rows: rowSplit.frozen, columns: columnSplit.scrollable, frozenRow: true, frozenColumn: false },
      { rows: rowSplit.scrollable, columns: columnSplit.frozen, frozenRow: false, frozenColumn: true },
      {
        rows: rowSplit.scrollable,
        columns: columnSplit.scrollable,
        frozenRow: false,
        frozenColumn: false
      }
    ]);

    this.#ctx.strokeStyle = SELECTION_STROKE_COLOR;
    this.#ctx.lineWidth = SELECTION_STROKE_WIDTH;
    for (const quadrant of quadrants) {
      const rect = this.#quadrantRect(quadrant);
      this.#ctx.strokeRect(rect.x, rect.y, rect.width, rect.height);
    }

    if (quadrants.length > 0) {
      this.#paintFillHandle(getSelectionBottomRightRect(range, this.#layout, this.#viewport));
    }

    const previewRange = this.#getFillPreviewRange();
    if (previewRange) {
      this.#paintFillPreview(previewRange);
    }
  }

  /** Dashed preview of the range a fill-handle drag would extend into. */
  #paintFillPreview(range: CellRange): void {
    const rect = this.#boundingRect(range);
    this.#ctx.strokeStyle = SELECTION_STROKE_COLOR;
    this.#ctx.lineWidth = SELECTION_STROKE_WIDTH;
    this.#ctx.setLineDash([4, 3]);
    this.#ctx.strokeRect(rect.x, rect.y, rect.width, rect.height);
    this.#ctx.setLineDash([]);
  }

  #boundingRect(range: CellRange): CellRect {
    const topLeft = toViewportRect(this.#layout.getCellRect(range.startRow, range.startColumn), this.#viewport);
    const bottomRight = toViewportRect(
      this.#layout.getCellRect(range.endRow, range.endColumn),
      this.#viewport
    );
    return {
      x: topLeft.x,
      y: topLeft.y,
      width: bottomRight.x + bottomRight.width - topLeft.x,
      height: bottomRight.y + bottomRight.height - topLeft.y
    };
  }

  /**
   * Purely cosmetic fill handle at the bottom-right corner of the
   * selection, matching Excel's affordance — no drag-to-fill interaction is
   * wired to it (see doc/progress.md's "不做的事" for this patch).
   */
  #paintFillHandle(rect: CellRect): void {
    this.#ctx.fillStyle = SELECTION_STROKE_COLOR;
    this.#ctx.fillRect(
      rect.x + rect.width - SELECTION_HANDLE_SIZE / 2,
      rect.y + rect.height - SELECTION_HANDLE_SIZE / 2,
      SELECTION_HANDLE_SIZE,
      SELECTION_HANDLE_SIZE
    );
  }

  #quadrantRect(quadrant: Quadrant): CellRect {
    const lastRow = quadrant.rows.start + quadrant.rows.count - 1;
    const lastColumn = quadrant.columns.start + quadrant.columns.count - 1;
    const topLeft = this.#translatedCellRect(
      quadrant.rows.start,
      quadrant.columns.start,
      quadrant.frozenRow,
      quadrant.frozenColumn
    );
    const bottomRight = this.#translatedCellRect(
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

  #translatedCellRect(row: number, column: number, frozenRow: boolean, frozenColumn: boolean): CellRect {
    return toViewportRect(this.#layout.getCellRect(row, column), this.#viewport, {
      row: frozenRow,
      column: frozenColumn
    });
  }
}
