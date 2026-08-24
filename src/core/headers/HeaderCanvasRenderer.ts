import type { CellRect } from "../layout";
import type { Canvas2DLike } from "../render";
import type { SceneColumnHeaderFragment, SceneSnapshot } from "../scene";
import {
  FROZEN_SEPARATOR_THICKNESS,
  GRID_LINE_COLOR,
  HEADER_BACKGROUND_COLOR,
  HEADER_BORDER_COLOR,
  HEADER_FONT,
  HEADER_SORT_INDICATOR_COLOR,
  TEXT_COLOR
} from "../render/theme";
import type { HeaderManager } from "./HeaderManager";
import type { HeaderAction, HeaderRenderContext, HeaderState, RuntimeHeaderGroup } from "./types";

export interface HeaderCanvasRendererOptions {
  readonly ctx: Canvas2DLike;
  readonly getScene: () => SceneSnapshot;
  readonly headerManager: HeaderManager;
  readonly getHeaderState?: (column: number) => HeaderState;
}

const DEFAULT_PADDING = 4;

/**
 * Width, in pixels, of the clickable strip reserved at the trailing edge of
 * a *default*-rendered leaf header cell for `toggle-sort`. Deliberately not
 * the whole cell: the rest of the cell is where a plain click still selects
 * the column (`headers.spec.ts`/`nested-headers.spec.ts`/`row-headers.spec.ts`
 * all click at column midpoints and must keep doing that), so only a narrow
 * reserved zone lets both gestures coexist on the same default header.
 */
export const DEFAULT_HEADER_SORT_ZONE_WIDTH = 18;

/**
 * Narrowest column that gets a sort zone at all. Below this the zone would
 * cover the whole header and there would be nowhere left to click to select
 * the column — the two gestures stop coexisting and selection is simply
 * lost for that column. A column too narrow to hold both keeps selection,
 * which is the gesture that has always been there; sorting it stays
 * available through `sortByColumn`.
 */
export const DEFAULT_HEADER_SORT_ZONE_MIN_WIDTH = DEFAULT_HEADER_SORT_ZONE_WIDTH * 2;

/**
 * Hit-test for the default (no custom `HeaderRenderer`) header painter —
 * the canvas analog of `defaultHeaderPainter` above. Columns with a custom
 * canvas renderer supply their own `hitTest`; this is only consulted when
 * a leaf column has no renderer at all, so every default header gets
 * click-to-sort without every existing "click the header to select the
 * column" test changing meaning.
 *
 * `point` must be *rect-relative* (the same `localPoint` shape Scene's
 * `hitTestColumnHeader` produces, matching the coordinate convention used
 * by every other canvas hit-test in this codebase) — not absolute.
 */
export function defaultHeaderHitTest(zone: {
  readonly rect: CellRect;
  readonly point: { readonly x: number; readonly y: number };
  /**
   * Omitted means sortable, matching `HeaderState.sortable`. A column that
   * opted out claims no zone, so its whole header stays available for
   * click-to-select — and there is no mark drawn there to click either.
   */
  readonly sortable?: boolean;
}): HeaderAction | undefined {
  if (zone.sortable === false) {
    return undefined;
  }
  if (zone.rect.width < DEFAULT_HEADER_SORT_ZONE_MIN_WIDTH) {
    return undefined;
  }
  if (zone.point.x >= zone.rect.width - DEFAULT_HEADER_SORT_ZONE_WIDTH) {
    return { type: "toggle-sort" };
  }
  return undefined;
}

/**
 * Drawn on a sortable column that is not currently sorted. Without it the
 * column advertises nothing: before this existed, a header stayed bare until
 * *after* a sort, so the only way to discover that clicking sorted was to
 * click and find out — and the click had to land in the trailing zone while
 * the resulting ▲ appeared next to the title, at the other end of the cell.
 */
export const HEADER_SORT_NEUTRAL_GLYPH = "⇅";

/** `undefined` when the column advertises no sort affordance at all. */
function sortGlyph(state: HeaderState): string | undefined {
  if (state.sortable === false) {
    return undefined;
  }
  if (state.sorted === "asc") {
    return "▲";
  }
  if (state.sorted === "desc") {
    return "▼";
  }
  return HEADER_SORT_NEUTRAL_GLYPH;
}

function defaultHeaderText(title: string, state: HeaderState): string {
  return state.filtered ? `${title} ●` : title;
}

function defaultHeaderPainter(context: HeaderRenderContext): void {
  const { ctx, rect, title, state } = context;
  const glyph = sortGlyph(state);
  const middle = rect.y + rect.height / 2;
  ctx.font = HEADER_FONT;
  ctx.textBaseline = "middle";

  // The title is clipped to whatever the sort mark does not occupy, so a long
  // title cannot run underneath the mark. Nothing is reserved when the column
  // has no mark, so a non-sortable column keeps the full width it always had.
  const titleWidth = glyph === undefined
    ? rect.width
    : Math.max(0, rect.width - DEFAULT_HEADER_SORT_ZONE_WIDTH);
  ctx.save();
  ctx.beginPath();
  ctx.rect(rect.x, rect.y, titleWidth, rect.height);
  ctx.clip();
  ctx.fillStyle = TEXT_COLOR;
  ctx.textAlign = "left";
  ctx.fillText(defaultHeaderText(title, state), rect.x + DEFAULT_PADDING, middle);
  ctx.restore();

  if (glyph === undefined) {
    return;
  }
  // Right-aligned inside the same trailing strip `defaultHeaderHitTest`
  // claims, so what you click and what you see are the same place.
  ctx.fillStyle = state.sorted === "asc" || state.sorted === "desc"
    ? TEXT_COLOR
    : HEADER_SORT_INDICATOR_COLOR;
  ctx.textAlign = "right";
  ctx.fillText(glyph, rect.x + rect.width - DEFAULT_PADDING, middle);
  ctx.textAlign = "left";
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
