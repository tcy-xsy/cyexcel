import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { CellRect, IndexRange, LayoutEngine, ViewportState } from "../layout";
import { splitByFrozenBoundary, toViewportRect } from "../layout";
import type { CellRange } from "../types";
import type { Canvas2DLike } from "./Canvas2DLike";
import type { DirtyState } from "./RenderScheduler";
import {
  CELL_FONT,
  CELL_PADDING,
  FROZEN_SEPARATOR_COLOR,
  FROZEN_SEPARATOR_THICKNESS,
  GRID_LINE_COLOR,
  TEXT_COLOR
} from "./theme";

export interface CellPaintContext {
  readonly ctx: Canvas2DLike;
  readonly rect: CellRect;
  readonly row: number;
  readonly column: number;
  readonly value: unknown;
}

/**
 * Contract is synchronous (return value ignored except to detect a Promise
 * misuse — see `checkSyncRenderers`). Typed `unknown` rather than `void` so
 * that check isn't silently defeated by TypeScript's lenient void-return
 * position.
 */
export type CellPainter = (context: CellPaintContext) => unknown;

export interface ContentLayerRendererOptions {
  readonly ctx: Canvas2DLike;
  readonly layout: LayoutEngine;
  readonly getCellValue: (row: number, column: number) => unknown;
  readonly painter?: CellPainter;
  readonly onError?: (error: unknown, cell: { row: number; column: number }) => void;
  /** Throw immediately if a painter returns a thenable. Default: true. */
  readonly checkSyncRenderers?: boolean;
  /** Draw a light gridline around every cell after it's painted. Default: true. */
  readonly showGridLines?: boolean;
}

function defaultPainter(context: CellPaintContext): void {
  const { ctx, rect, value } = context;
  ctx.fillStyle = TEXT_COLOR;
  ctx.font = CELL_FONT;
  ctx.textBaseline = "middle";
  ctx.textAlign = "left";
  ctx.fillText(
    value === undefined || value === null ? "" : String(value),
    rect.x + CELL_PADDING,
    rect.y + rect.height / 2
  );
}

function isThenable(value: unknown): boolean {
  return (
    typeof value === "object" &&
    value !== null &&
    typeof (value as { then?: unknown }).then === "function"
  );
}

function intersectIndexRange(a: IndexRange, b: IndexRange): IndexRange {
  const start = Math.max(a.start, b.start);
  const end = Math.min(a.start + a.count, b.start + b.count);
  return { start, count: Math.max(0, end - start) };
}

/**
 * One of up to four rectangular row×column combinations a paint pass needs
 * to handle separately — frozen rows/columns are always shown regardless
 * of scroll, so they can't share one "visible range" with the scrollable
 * body the way a single un-frozen table could. See doc/architecture.md
 * §6.1 for why this is modeled as up to four independent regions rather
 * than one merged range.
 */
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
 * Paints cell values onto a Canvas2DLike surface, clipped to whatever is
 * currently visible per LayoutEngine + the last setViewport() call. Never
 * depends on DataSource/TableStore directly — only a getCellValue callback
 * — so it can be unit-tested without constructing a real data layer.
 */
export class ContentLayerRenderer {
  readonly #ctx: Canvas2DLike;
  readonly #layout: LayoutEngine;
  readonly #getCellValue: (row: number, column: number) => unknown;
  readonly #painter: CellPainter;
  readonly #onError:
    | ((error: unknown, cell: { row: number; column: number }) => void)
    | undefined;
  readonly #checkSyncRenderers: boolean;
  readonly #showGridLines: boolean;
  #viewport: ViewportState = { scrollTop: 0, scrollLeft: 0, width: 0, height: 0 };

  constructor(options: ContentLayerRendererOptions) {
    this.#ctx = options.ctx;
    this.#layout = options.layout;
    this.#getCellValue = options.getCellValue;
    this.#painter = options.painter ?? defaultPainter;
    this.#onError = options.onError;
    this.#checkSyncRenderers = options.checkSyncRenderers ?? true;
    this.#showGridLines = options.showGridLines ?? true;
  }

  setViewport(viewport: ViewportState): void {
    this.#viewport = viewport;
  }

  paintContent(dirty: DirtyState): void {
    if (dirty.kind === "none") {
      return;
    }

    if (dirty.kind === "full") {
      for (const quadrant of this.#currentQuadrants()) {
        this.#paintQuadrant(quadrant);
      }
    } else {
      for (const range of dirty.ranges) {
        for (const quadrant of this.#quadrantsForRange(range)) {
          this.#paintQuadrant(quadrant);
        }
      }
    }

    this.#paintFrozenSeparators();
  }

  /**
   * A heavier line marking the fixed screen boundary between the frozen
   * band and the scrollable body — always at the same pixel position
   * regardless of scroll, since that's what "frozen" means. Drawn after
   * every paint pass (not just "full") because a regions repaint can clear
   * pixels right up to that boundary without necessarily repainting it.
   */
  #paintFrozenSeparators(): void {
    const frozenRowCount = this.#layout.frozenRowCount;
    const frozenColumnCount = this.#layout.frozenColumnCount;

    if (frozenRowCount > 0) {
      const y = frozenRowCount * this.#layout.rowHeight;
      this.#ctx.fillStyle = FROZEN_SEPARATOR_COLOR;
      this.#ctx.fillRect(
        0,
        y - FROZEN_SEPARATOR_THICKNESS / 2,
        this.#viewport.width,
        FROZEN_SEPARATOR_THICKNESS
      );
    }

    if (frozenColumnCount > 0) {
      const lastFrozenColumn = frozenColumnCount - 1;
      const x =
        this.#layout.getColumnLeft(lastFrozenColumn) + this.#layout.getColumnWidth(lastFrozenColumn);
      this.#ctx.fillStyle = FROZEN_SEPARATOR_COLOR;
      this.#ctx.fillRect(
        x - FROZEN_SEPARATOR_THICKNESS / 2,
        0,
        FROZEN_SEPARATOR_THICKNESS,
        this.#viewport.height
      );
    }
  }

  /** The up-to-four regions currently on screen: frozen corner, frozen row band, frozen column band, scrollable body. */
  #currentQuadrants(): Quadrant[] {
    const frozenRows: IndexRange = { start: 0, count: this.#layout.frozenRowCount };
    const scrollableRows = this.#layout.getVisibleRowRange(this.#viewport);
    const frozenColumns: IndexRange = { start: 0, count: this.#layout.frozenColumnCount };
    const scrollableColumns = this.#layout.getVisibleColumnRange(this.#viewport);

    return nonEmptyQuadrants([
      { rows: frozenRows, columns: frozenColumns, frozenRow: true, frozenColumn: true },
      { rows: frozenRows, columns: scrollableColumns, frozenRow: true, frozenColumn: false },
      { rows: scrollableRows, columns: frozenColumns, frozenRow: false, frozenColumn: true },
      { rows: scrollableRows, columns: scrollableColumns, frozenRow: false, frozenColumn: false }
    ]);
  }

  /**
   * Same four regions, but clipped to one arbitrary dirty CellRange (e.g. a
   * single setCell patch) instead of "everything on screen" — the frozen
   * part of the range is always paintable (frozen cells are always shown),
   * the scrollable part still needs intersecting with what's currently
   * visible.
   */
  #quadrantsForRange(range: CellRange): Quadrant[] {
    const rowIndexRange: IndexRange = { start: range.startRow, count: range.endRow - range.startRow + 1 };
    const columnIndexRange: IndexRange = {
      start: range.startColumn,
      count: range.endColumn - range.startColumn + 1
    };

    const rowSplit = splitByFrozenBoundary(rowIndexRange, this.#layout.frozenRowCount);
    const columnSplit = splitByFrozenBoundary(columnIndexRange, this.#layout.frozenColumnCount);

    const scrollableRows = intersectIndexRange(
      rowSplit.scrollable,
      this.#layout.getVisibleRowRange(this.#viewport)
    );
    const scrollableColumns = intersectIndexRange(
      columnSplit.scrollable,
      this.#layout.getVisibleColumnRange(this.#viewport)
    );

    return nonEmptyQuadrants([
      { rows: rowSplit.frozen, columns: columnSplit.frozen, frozenRow: true, frozenColumn: true },
      { rows: rowSplit.frozen, columns: scrollableColumns, frozenRow: true, frozenColumn: false },
      { rows: scrollableRows, columns: columnSplit.frozen, frozenRow: false, frozenColumn: true },
      { rows: scrollableRows, columns: scrollableColumns, frozenRow: false, frozenColumn: false }
    ]);
  }

  #paintQuadrant(quadrant: Quadrant): void {
    const clipRect = this.#quadrantClipRect(quadrant);
    this.#ctx.save();
    this.#ctx.beginPath();
    this.#ctx.rect(clipRect.x, clipRect.y, clipRect.width, clipRect.height);
    this.#ctx.clip();

    const lastRow = quadrant.rows.start + quadrant.rows.count - 1;
    const lastColumn = quadrant.columns.start + quadrant.columns.count - 1;
    for (let row = quadrant.rows.start; row <= lastRow; row += 1) {
      for (let column = quadrant.columns.start; column <= lastColumn; column += 1) {
        this.#paintCell(row, column, quadrant.frozenRow, quadrant.frozenColumn);
      }
    }

    this.#ctx.restore();
  }

  #quadrantClipRect(quadrant: Quadrant): CellRect {
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

  #paintCell(row: number, column: number, frozenRow: boolean, frozenColumn: boolean): void {
    const rect = this.#translatedCellRect(row, column, frozenRow, frozenColumn);
    this.#ctx.clearRect(rect.x, rect.y, rect.width, rect.height);

    const value = this.#getCellValue(row, column);
    const context: CellPaintContext = { ctx: this.#ctx, rect, row, column, value };

    // Isolated in its own save/restore so a painter (built-in default or a
    // custom CanvasCellRenderer) that changes ctx state — font, fillStyle,
    // textAlign, ... — and never restores it itself can't leak that state
    // into the next cell painted on this same canvas. Canvas 2D state is
    // otherwise persistent across draw calls; without this, one column's
    // renderer choosing e.g. a different font silently changes how every
    // later cell that never sets its own font gets drawn.
    // Isolated in its own save/restore so a painter (built-in default or a
    // custom CanvasCellRenderer) that changes ctx state — font, fillStyle,
    // textAlign, ... — and never restores it itself can't leak that state
    // into the next cell painted on this same canvas. Canvas 2D state is
    // otherwise persistent across draw calls; without this, one column's
    // renderer choosing e.g. a different font silently changes how every
    // later cell that never sets its own font gets drawn.
    this.#ctx.save();
    let result: unknown;
    try {
      result = this.#painter(context);
    } catch (error) {
      this.#ctx.restore();
      this.#onError?.(error, { row, column });
      this.#ctx.save();
      defaultPainter(context);
      this.#ctx.restore();
      this.#paintGridLine(rect);
      return;
    }
    this.#ctx.restore();

    if (this.#checkSyncRenderers && isThenable(result)) {
      throw new CyExcelError(
        CyExcelErrorCode.AsyncRendererNotSupported,
        `Cell renderer for (${row}, ${column}) returned a Promise; cell renderers must be synchronous`
      );
    }

    this.#paintGridLine(rect);
  }

  #paintGridLine(rect: CellRect): void {
    if (!this.#showGridLines) {
      return;
    }
    this.#ctx.strokeStyle = GRID_LINE_COLOR;
    this.#ctx.lineWidth = 1;
    this.#ctx.strokeRect(rect.x, rect.y, rect.width, rect.height);
  }

  #translatedCellRect(row: number, column: number, frozenRow: boolean, frozenColumn: boolean): CellRect {
    return toViewportRect(this.#layout.getCellRect(row, column), this.#viewport, {
      row: frozenRow,
      column: frozenColumn
    });
  }
}
