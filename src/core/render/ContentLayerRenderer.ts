import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { CellRect, LayoutEngine, ViewportState } from "../layout";
import type { CellRange } from "../types";
import type { Canvas2DLike } from "./Canvas2DLike";
import type { DirtyState } from "./RenderScheduler";

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
}

const DEFAULT_PADDING = 4;

function defaultPainter(context: CellPaintContext): void {
  const { ctx, rect, value } = context;
  ctx.fillStyle = "#111827";
  ctx.textBaseline = "middle";
  ctx.textAlign = "left";
  ctx.fillText(
    value === undefined || value === null ? "" : String(value),
    rect.x + DEFAULT_PADDING,
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

function intersectRange(a: CellRange, b: CellRange): CellRange | null {
  const startRow = Math.max(a.startRow, b.startRow);
  const startColumn = Math.max(a.startColumn, b.startColumn);
  const endRow = Math.min(a.endRow, b.endRow);
  const endColumn = Math.min(a.endColumn, b.endColumn);

  if (startRow > endRow || startColumn > endColumn) {
    return null;
  }
  return { startRow, startColumn, endRow, endColumn };
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
  #viewport: ViewportState = { scrollTop: 0, scrollLeft: 0, width: 0, height: 0 };

  constructor(options: ContentLayerRendererOptions) {
    this.#ctx = options.ctx;
    this.#layout = options.layout;
    this.#getCellValue = options.getCellValue;
    this.#painter = options.painter ?? defaultPainter;
    this.#onError = options.onError;
    this.#checkSyncRenderers = options.checkSyncRenderers ?? true;
  }

  setViewport(viewport: ViewportState): void {
    this.#viewport = viewport;
  }

  paintContent(dirty: DirtyState): void {
    if (dirty.kind === "none") {
      return;
    }

    const visibleRows = this.#layout.getVisibleRowRange(this.#viewport);
    const visibleColumns = this.#layout.getVisibleColumnRange(this.#viewport);
    if (visibleRows.count === 0 || visibleColumns.count === 0) {
      return;
    }

    const visibleRange: CellRange = {
      startRow: visibleRows.start,
      startColumn: visibleColumns.start,
      endRow: visibleRows.start + visibleRows.count - 1,
      endColumn: visibleColumns.start + visibleColumns.count - 1
    };

    if (dirty.kind === "full") {
      this.#paintRange(visibleRange);
      return;
    }

    for (const range of dirty.ranges) {
      const clipped = intersectRange(range, visibleRange);
      if (clipped) {
        this.#paintRange(clipped);
      }
    }
  }

  #paintRange(range: CellRange): void {
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      for (let column = range.startColumn; column <= range.endColumn; column += 1) {
        this.#paintCell(row, column);
      }
    }
  }

  #paintCell(row: number, column: number): void {
    const rect = this.#layout.getCellRect(row, column);
    this.#ctx.clearRect(rect.x, rect.y, rect.width, rect.height);

    const value = this.#getCellValue(row, column);
    const context: CellPaintContext = { ctx: this.#ctx, rect, row, column, value };

    let result: unknown;
    try {
      result = this.#painter(context);
    } catch (error) {
      this.#onError?.(error, { row, column });
      defaultPainter(context);
      return;
    }

    if (this.#checkSyncRenderers && isThenable(result)) {
      throw new CyExcelError(
        CyExcelErrorCode.AsyncRendererNotSupported,
        `Cell renderer for (${row}, ${column}) returned a Promise; cell renderers must be synchronous`
      );
    }
  }
}
