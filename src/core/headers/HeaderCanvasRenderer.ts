import type { CellRect } from "../layout";
import type { Canvas2DLike } from "../render";
import type { SceneColumnHeaderFragment, SceneSnapshot } from "../scene";
import {
  FROZEN_SEPARATOR_THICKNESS,
  GRID_LINE_COLOR,
  HEADER_BACKGROUND_COLOR,
  HEADER_BORDER_COLOR,
  HEADER_FONT,
  TEXT_COLOR
} from "../render/theme";
import type { HeaderManager } from "./HeaderManager";
import type { HeaderRenderContext, HeaderState, RuntimeHeaderGroup } from "./types";

export interface HeaderCanvasRendererOptions {
  readonly ctx: Canvas2DLike;
  readonly getScene: () => SceneSnapshot;
  readonly headerManager: HeaderManager;
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

/**
 * Paints the (up to two-row) column header band: an optional group row on
 * top, the leaf column title row below it. Headers never move vertically
 * (the band is always fully pinned), only horizontally as columns scroll —
 * frozen columns stay put on that axis too. Both membership and clipping
 * come from Scene's canonical column-header panes, so a partially scrolled cell can
 * never paint back across the frozen boundary. A group whose column span
 * straddles pane boundaries is drawn as one fragment per pane.
 */
export class HeaderCanvasRenderer {
  readonly #ctx: Canvas2DLike;
  readonly #getScene: () => SceneSnapshot;
  readonly #headerManager: HeaderManager;
  readonly #getHeaderState: (column: number) => HeaderState;

  constructor(options: HeaderCanvasRendererOptions) {
    this.#ctx = options.ctx;
    this.#getScene = options.getScene;
    this.#headerManager = options.headerManager;
    this.#getHeaderState = options.getHeaderState ?? (() => ({}));
  }

  paint(scene: SceneSnapshot = this.#getScene()): void {
    const surface = scene.getSurface("column-header");
    this.#ctx.clearRect(0, 0, surface.rootRect.width, surface.rootRect.height);
    const leafBand = scene.getColumnHeaderLeafBand();

    for (const paneFragment of scene.getVisibleColumnHeaderFragments(leafBand)) {
      for (let offset = 0; offset < paneFragment.columns.count; offset += 1) {
        const column = paneFragment.columns.start + offset;
        const fragment = scene
          .getColumnHeaderBandFragments(leafBand, column, column)
          .find((candidate) => candidate.pane === paneFragment.pane);
        if (fragment) {
          this.#paintLeafHeader(column, fragment);
        }
      }
    }

    for (const band of scene.getColumnHeaderBands()) {
      if (band.role !== "group") {
        continue;
      }
      for (const group of this.#headerManager.getGroups(band.level)) {
        for (const fragment of scene.getColumnHeaderBandFragments(
          band,
          group.startColumn,
          group.endColumn
        )) {
          this.#paintGroupFragment(group, fragment);
        }
      }
    }

    this.#paintBottomBorder(scene);
  }

  /** A border separating the header band from the body, drawn once per paint() call. */
  #paintBottomBorder(scene: SceneSnapshot): void {
    const surface = scene.getSurface("column-header");
    this.#ctx.fillStyle = HEADER_BORDER_COLOR;
    this.#ctx.fillRect(
      0,
      surface.rootRect.height - FROZEN_SEPARATOR_THICKNESS / 2,
      surface.rootRect.width,
      FROZEN_SEPARATOR_THICKNESS
    );
  }

  #paintLeafHeader(column: number, fragment: SceneColumnHeaderFragment): void {
    const header = this.#headerManager.getHeader(column);
    const rect = fragment.rect;

    this.#withPaneClip(rect, fragment.clipRect, () => {
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
      // Set fresh (not inherited from whatever a custom renderer above left
      // behind) — same reasoning as ContentLayerRenderer's own per-cell
      // gridline stroke, and safe regardless because #withClip's
      // save()/restore() already isolates this from every other cell.
      this.#ctx.strokeStyle = GRID_LINE_COLOR;
      this.#ctx.lineWidth = 1;
      this.#ctx.strokeRect(rect.x, rect.y, rect.width, rect.height);
    });
  }

  #paintGroupFragment(
    group: RuntimeHeaderGroup,
    fragment: SceneColumnHeaderFragment
  ): void {
    const rect = fragment.rect;
    this.#withPaneClip(rect, fragment.clipRect, () => {
      this.#ctx.clearRect(rect.x, rect.y, rect.width, rect.height);
      this.#ctx.fillStyle = HEADER_BACKGROUND_COLOR;
      this.#ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
      this.#ctx.fillStyle = TEXT_COLOR;
      this.#ctx.font = HEADER_FONT;
      this.#ctx.textBaseline = "middle";
      this.#ctx.textAlign = "left";
      this.#ctx.fillText(group.title, rect.x + DEFAULT_PADDING, rect.y + rect.height / 2);
      this.#ctx.strokeStyle = GRID_LINE_COLOR;
      this.#ctx.lineWidth = 1;
      this.#ctx.strokeRect(rect.x, rect.y, rect.width, rect.height);
    });
  }

  #withPaneClip(rect: CellRect, clipRect: CellRect, paint: () => void): void {
    const left = Math.max(rect.x, clipRect.x);
    const top = Math.max(rect.y, clipRect.y);
    const right = Math.min(rect.x + rect.width, clipRect.x + clipRect.width);
    const bottom = Math.min(rect.y + rect.height, clipRect.y + clipRect.height);
    if (right <= left || bottom <= top) {
      return;
    }
    this.#ctx.save();
    this.#ctx.beginPath();
    this.#ctx.rect(left, top, right - left, bottom - top);
    this.#ctx.clip();
    paint();
    this.#ctx.restore();
  }
}
