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
import type { HeaderManager } from "./HeaderManager";
import type { HeaderGroupFragment, HeaderRenderContext, HeaderState } from "./types";

export interface HeaderCanvasRendererOptions {
  readonly ctx: Canvas2DLike;
  readonly layout: LayoutEngine;
  readonly headerManager: HeaderManager;
  readonly columnHeaderHeight: number;
  /** 0 (default) means no group row is drawn at all. */
  readonly groupHeaderHeight?: number;
  readonly getHeaderState?: (column: number) => HeaderState;
}

const DEFAULT_PADDING = 4;

function defaultHeaderText(title: string, state: HeaderState): string {
  const sortSuffix = state.sorted === "asc" ? " ▲" : state.sorted === "desc" ? " ▼" : "";
  const filterSuffix = state.filtered ? " ●" : "";
  return `${title}${sortSuffix}${filterSuffix}`;
}

function defaultHeaderPainter(context: HeaderRenderContext): void {
  const { ctx, rect, title, state } = context;
  ctx.fillStyle = TEXT_COLOR;
  ctx.font = HEADER_FONT;
  ctx.textBaseline = "middle";
  ctx.textAlign = "left";
  ctx.fillText(defaultHeaderText(title, state), rect.x + DEFAULT_PADDING, rect.y + rect.height / 2);
}

function rangesOverlap(start: number, end: number, range: IndexRange): boolean {
  if (range.count === 0) {
    return false;
  }
  const rangeEnd = range.start + range.count - 1;
  return start <= rangeEnd && end >= range.start;
}

/**
 * Paints the (up to two-row) column header band: an optional group row on
 * top, the leaf column title row below it. Headers never move vertically
 * (the band is always fully pinned), only horizontally as columns scroll —
 * frozen columns stay put on that axis too, via the same
 * toViewportRect(..., { column: frozen }) pattern ContentLayerRenderer and
 * SelectionLayerRenderer use (doc/architecture.md §6.1). A group whose
 * column span straddles the frozen boundary is drawn as two separate
 * clipped fragments (HeaderManager.getGroupFragments()) that both show the
 * same title — Step 9's 4th pre-written test.
 */
export class HeaderCanvasRenderer {
  readonly #ctx: Canvas2DLike;
  readonly #layout: LayoutEngine;
  readonly #headerManager: HeaderManager;
  readonly #columnHeaderHeight: number;
  readonly #groupHeaderHeight: number;
  readonly #getHeaderState: (column: number) => HeaderState;
  #viewport: ViewportState = { scrollTop: 0, scrollLeft: 0, width: 0, height: 0 };

  constructor(options: HeaderCanvasRendererOptions) {
    this.#ctx = options.ctx;
    this.#layout = options.layout;
    this.#headerManager = options.headerManager;
    this.#columnHeaderHeight = options.columnHeaderHeight;
    this.#groupHeaderHeight = options.groupHeaderHeight ?? 0;
    this.#getHeaderState = options.getHeaderState ?? (() => ({}));
  }

  /** Total pixel height the header band occupies — group row (if any) + leaf title row. */
  get totalHeight(): number {
    return this.#columnHeaderHeight + this.#groupHeaderHeight;
  }

  setViewport(viewport: ViewportState): void {
    this.#viewport = viewport;
  }

  paint(): void {
    const frozenColumns: IndexRange = { start: 0, count: this.#layout.frozenColumnCount };
    const visibleColumns = this.#layout.getVisibleColumnRange(this.#viewport);

    this.#paintLeafHeaders(frozenColumns, true);
    this.#paintLeafHeaders(visibleColumns, false);

    if (this.#groupHeaderHeight > 0) {
      for (const fragment of this.#headerManager.getGroupFragments(this.#layout.frozenColumnCount)) {
        if (fragment.frozenColumn || rangesOverlap(fragment.startColumn, fragment.endColumn, visibleColumns)) {
          this.#paintGroupFragment(fragment);
        }
      }
    }

    this.#paintBottomBorder();
  }

  /** A border separating the header band from the body, drawn once per paint() call. */
  #paintBottomBorder(): void {
    this.#ctx.fillStyle = HEADER_BORDER_COLOR;
    this.#ctx.fillRect(
      0,
      this.totalHeight - FROZEN_SEPARATOR_THICKNESS / 2,
      this.#viewport.width,
      FROZEN_SEPARATOR_THICKNESS
    );
  }

  #paintLeafHeaders(range: IndexRange, frozen: boolean): void {
    for (let column = range.start; column < range.start + range.count; column += 1) {
      this.#paintLeafHeader(column, frozen);
    }
  }

  #paintLeafHeader(column: number, frozen: boolean): void {
    const header = this.#headerManager.getHeader(column);
    const rect = this.#leafRect(column, frozen);

    this.#withClip(rect, () => {
      this.#ctx.clearRect(rect.x, rect.y, rect.width, rect.height);
      this.#ctx.fillStyle = HEADER_BACKGROUND_COLOR;
      this.#ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
      const context: HeaderRenderContext = {
        ctx: this.#ctx,
        rect,
        column,
        title: header.title,
        state: this.#getHeaderState(column)
      };

      if (header.renderer?.kind === "canvas") {
        header.renderer.paint(context);
      } else if (header.renderer?.kind !== "dom") {
        // No renderer at all -> default text painter. A "dom" renderer
        // paints nothing here — DomHeaderManager owns that cell entirely.
        defaultHeaderPainter(context);
      }
    });
  }

  #paintGroupFragment(fragment: HeaderGroupFragment): void {
    const rect = this.#groupRect(fragment);
    this.#withClip(rect, () => {
      this.#ctx.clearRect(rect.x, rect.y, rect.width, rect.height);
      this.#ctx.fillStyle = HEADER_BACKGROUND_COLOR;
      this.#ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
      this.#ctx.fillStyle = TEXT_COLOR;
      this.#ctx.font = HEADER_FONT;
      this.#ctx.textBaseline = "middle";
      this.#ctx.textAlign = "left";
      this.#ctx.fillText(fragment.group.title, rect.x + DEFAULT_PADDING, rect.y + rect.height / 2);
    });
  }

  #withClip(rect: CellRect, paint: () => void): void {
    this.#ctx.save();
    this.#ctx.beginPath();
    this.#ctx.rect(rect.x, rect.y, rect.width, rect.height);
    this.#ctx.clip();
    paint();
    this.#ctx.restore();
  }

  #leafRect(column: number, frozen: boolean): CellRect {
    const absolute: CellRect = {
      x: this.#layout.getColumnLeft(column),
      y: this.#groupHeaderHeight,
      width: this.#layout.getColumnWidth(column),
      height: this.#columnHeaderHeight
    };
    return toViewportRect(absolute, this.#viewport, { row: true, column: frozen });
  }

  #groupRect(fragment: HeaderGroupFragment): CellRect {
    const left = this.#layout.getColumnLeft(fragment.startColumn);
    const right =
      this.#layout.getColumnLeft(fragment.endColumn) + this.#layout.getColumnWidth(fragment.endColumn);
    const absolute: CellRect = { x: left, y: 0, width: right - left, height: this.#groupHeaderHeight };
    return toViewportRect(absolute, this.#viewport, { row: true, column: fragment.frozenColumn });
  }
}
