import type { CellRect, IndexRange, LayoutEngine, ViewportState } from "../layout";
import { toViewportRect } from "../layout";
import type { Canvas2DLike } from "../render";
import {
  FROZEN_SEPARATOR_THICKNESS,
  HEADER_BACKGROUND_COLOR,
  HEADER_BORDER_COLOR,
  HEADER_FONT,
  TEXT_COLOR
} from "../render/theme";

export interface RowHeaderCanvasRendererOptions {
  readonly ctx: Canvas2DLike;
  readonly layout: LayoutEngine;
  readonly rowHeaderWidth: number;
  /** Pixel height of the column header band above the grid — 0 means no corner cell is drawn. Mirrors HeaderCanvasRenderer's own columnHeaderHeight/groupHeaderHeight. */
  readonly headerHeight: number;
}

/**
 * Paints the row-number band on the left — the row-axis mirror of
 * HeaderCanvasRenderer's column title band. There's no equivalent of
 * HeaderManager here: row numbers are always the sequential `row + 1`,
 * nothing to resolve from a per-row definition (no ColumnDefinition-style
 * config exists for rows), so this renderer has no companion "manager"
 * class and no custom-renderer hook — see doc/progress.md's row-header
 * patch for why that's a deliberate scope decision, not an oversight.
 *
 * This canvas spans the *full* viewport height (headerHeight + body),
 * unlike the body layers which start below the header band — so it can
 * also paint the top-left "corner" cell itself instead of needing a third
 * canvas/DOM element just for that one cell.
 */
export class RowHeaderCanvasRenderer {
  readonly #ctx: Canvas2DLike;
  readonly #layout: LayoutEngine;
  readonly #rowHeaderWidth: number;
  readonly #headerHeight: number;
  #viewport: ViewportState = { scrollTop: 0, scrollLeft: 0, width: 0, height: 0 };

  constructor(options: RowHeaderCanvasRendererOptions) {
    this.#ctx = options.ctx;
    this.#layout = options.layout;
    this.#rowHeaderWidth = options.rowHeaderWidth;
    this.#headerHeight = options.headerHeight;
  }

  setViewport(viewport: ViewportState): void {
    this.#viewport = viewport;
  }

  paint(): void {
    if (this.#rowHeaderWidth <= 0) {
      return;
    }

    const frozenRows: IndexRange = { start: 0, count: this.#layout.frozenRowCount };
    const visibleRows = this.#layout.getVisibleRowRange(this.#viewport);

    this.#paintRows(frozenRows, true);
    this.#paintRows(visibleRows, false);

    if (this.#headerHeight > 0) {
      this.#paintCorner();
    }

    this.#paintRightBorder();
  }

  #paintCorner(): void {
    this.#withClip({ x: 0, y: 0, width: this.#rowHeaderWidth, height: this.#headerHeight }, (rect) => {
      this.#ctx.clearRect(rect.x, rect.y, rect.width, rect.height);
      this.#ctx.fillStyle = HEADER_BACKGROUND_COLOR;
      this.#ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
    });
  }

  #paintRightBorder(): void {
    this.#ctx.fillStyle = HEADER_BORDER_COLOR;
    this.#ctx.fillRect(
      this.#rowHeaderWidth - FROZEN_SEPARATOR_THICKNESS / 2,
      0,
      FROZEN_SEPARATOR_THICKNESS,
      this.#headerHeight + this.#viewport.height
    );
  }

  #paintRows(range: IndexRange, frozen: boolean): void {
    for (let row = range.start; row < range.start + range.count; row += 1) {
      this.#paintRow(row, frozen);
    }
  }

  #paintRow(row: number, frozen: boolean): void {
    const rect = this.#rowRect(row, frozen);
    this.#withClip(rect, (clipped) => {
      this.#ctx.clearRect(clipped.x, clipped.y, clipped.width, clipped.height);
      this.#ctx.fillStyle = HEADER_BACKGROUND_COLOR;
      this.#ctx.fillRect(clipped.x, clipped.y, clipped.width, clipped.height);
      this.#ctx.fillStyle = TEXT_COLOR;
      this.#ctx.font = HEADER_FONT;
      this.#ctx.textBaseline = "middle";
      this.#ctx.textAlign = "center";
      this.#ctx.fillText(String(row + 1), clipped.x + clipped.width / 2, clipped.y + clipped.height / 2);
    });
  }

  #withClip(rect: CellRect, paint: (rect: CellRect) => void): void {
    this.#ctx.save();
    this.#ctx.beginPath();
    this.#ctx.rect(rect.x, rect.y, rect.width, rect.height);
    this.#ctx.clip();
    paint(rect);
    this.#ctx.restore();
  }

  #rowRect(row: number, frozen: boolean): CellRect {
    const absolute: CellRect = {
      x: 0,
      y: this.#headerHeight + row * this.#layout.rowHeight,
      width: this.#rowHeaderWidth,
      height: this.#layout.rowHeight
    };
    return toViewportRect(absolute, this.#viewport, { row: frozen, column: true });
  }
}
