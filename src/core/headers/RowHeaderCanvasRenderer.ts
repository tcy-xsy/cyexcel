import { getRowBoundaryY } from "../interaction/RowReorderController";
import type { CellRect } from "../layout";
import type { Canvas2DLike } from "../render";
import type { SceneRowHeaderFragment, SceneSnapshot } from "../scene";
import {
  FROZEN_SEPARATOR_THICKNESS,
  GRID_LINE_COLOR,
  HEADER_BACKGROUND_COLOR,
  HEADER_BORDER_COLOR,
  HEADER_FONT,
  ROW_REORDER_INDICATOR_COLOR,
  ROW_REORDER_INDICATOR_THICKNESS,
  TEXT_COLOR
} from "../render/theme";

export interface RowHeaderCanvasRendererOptions {
  readonly ctx: Canvas2DLike;
  readonly getScene: () => SceneSnapshot;
  /** The row-reorder drag's current "insert before this row" boundary, or `null`/omitted when no drag is in progress — see RowReorderController. */
  readonly getReorderIndicatorRow?: () => number | null;
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
  readonly #getScene: () => SceneSnapshot;
  readonly #getReorderIndicatorRow: (() => number | null) | undefined;

  constructor(options: RowHeaderCanvasRendererOptions) {
    this.#ctx = options.ctx;
    this.#getScene = options.getScene;
    this.#getReorderIndicatorRow = options.getReorderIndicatorRow;
  }

  paint(scene: SceneSnapshot = this.#getScene()): void {
    const rowSurface = scene.getSurface("row-header");
    if (rowSurface.rootRect.width <= 0) {
      return;
    }

    const menuSurface = scene.getSurface("menu-overlay");
    this.#ctx.clearRect(0, 0, rowSurface.rootRect.width, menuSurface.rootRect.height);

    for (const paneFragment of scene.getRowHeaderFragments(0, scene.rowCount - 1)) {
      for (const row of scene.getRenderableRows(paneFragment.rows)) {
        const fragment = scene.getRowHeaderFragments(row, row)[0];
        if (fragment) {
          this.#paintRow(row, fragment, rowSurface.rootRect.y);
        }
      }
    }

    const corner = scene.getSurface("corner").rootRect;
    if (corner.width > 0 && corner.height > 0) {
      this.#paintCorner(corner);
    }

    this.#paintRightBorder(scene);
    this.#paintReorderIndicator(scene);
  }

  #paintReorderIndicator(scene: SceneSnapshot): void {
    const boundaryRow = this.#getReorderIndicatorRow?.() ?? null;
    const surface = scene.getSurface("row-header");
    if (boundaryRow === null || surface.rootRect.width <= 0) {
      return;
    }
    const y = surface.rootRect.y + getRowBoundaryY(boundaryRow, scene);
    this.#ctx.fillStyle = ROW_REORDER_INDICATOR_COLOR;
    this.#ctx.fillRect(
      0,
      y - ROW_REORDER_INDICATOR_THICKNESS / 2,
      surface.rootRect.width,
      ROW_REORDER_INDICATOR_THICKNESS
    );
  }

  #paintCorner(rect: CellRect): void {
    this.#withClip(rect, () => {
      this.#ctx.clearRect(rect.x, rect.y, rect.width, rect.height);
      this.#ctx.fillStyle = HEADER_BACKGROUND_COLOR;
      this.#ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
    });
  }

  #paintRightBorder(scene: SceneSnapshot): void {
    const rowSurface = scene.getSurface("row-header");
    const menuSurface = scene.getSurface("menu-overlay");
    this.#ctx.fillStyle = HEADER_BORDER_COLOR;
    this.#ctx.fillRect(
      rowSurface.rootRect.width - FROZEN_SEPARATOR_THICKNESS / 2,
      0,
      FROZEN_SEPARATOR_THICKNESS,
      menuSurface.rootRect.height
    );
  }

  #paintRow(
    row: number,
    fragment: SceneRowHeaderFragment,
    rootOffsetY: number
  ): void {
    const rect = {
      x: fragment.rect.x,
      y: rootOffsetY + fragment.rect.y,
      width: fragment.rect.width,
      height: fragment.rect.height
    };
    const top = rootOffsetY + Math.max(fragment.rect.y, fragment.clipRect.y);
    const bottom = rootOffsetY + Math.min(
      fragment.rect.y + fragment.rect.height,
      fragment.clipRect.y + fragment.clipRect.height
    );
    if (bottom <= top || rect.width <= 0) {
      return;
    }
    this.#withClip({ x: rect.x, y: top, width: rect.width, height: bottom - top }, () => {
      this.#ctx.clearRect(rect.x, rect.y, rect.width, rect.height);
      this.#ctx.fillStyle = HEADER_BACKGROUND_COLOR;
      this.#ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
      this.#ctx.fillStyle = TEXT_COLOR;
      this.#ctx.font = HEADER_FONT;
      this.#ctx.textBaseline = "middle";
      this.#ctx.textAlign = "center";
      this.#ctx.fillText(String(row + 1), rect.x + rect.width / 2, rect.y + rect.height / 2);
      // Same light-gridline treatment ContentLayerRenderer gives every body
      // cell — without it, adjacent row numbers have no visible boundary
      // between them (only the header/body divider on the far right did).
      this.#ctx.strokeStyle = GRID_LINE_COLOR;
      this.#ctx.lineWidth = 1;
      this.#ctx.strokeRect(rect.x, rect.y, rect.width, rect.height);
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

}
