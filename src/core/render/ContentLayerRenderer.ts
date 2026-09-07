import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { CellRect } from "../layout";
import type { SceneCellFragment, SceneSnapshot } from "../scene";
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
  readonly getScene: () => SceneSnapshot;
  readonly getCellValue: (row: number, column: number) => unknown;
  readonly painter?: CellPainter;
  readonly onError?: (error: unknown, cell: { row: number; column: number }) => void;
  /** Throw immediately if a painter returns a thenable. Default: true. */
  readonly checkSyncRenderers?: boolean;
  /** Draw a light gridline around every cell after it's painted. Default: true. */
  readonly showGridLines?: boolean;
  readonly findMergeAt?: (row: number, column: number) => CellRange | undefined;
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

/**
 * One of up to four rectangular row×column combinations a paint pass needs
 * to handle separately — frozen rows/columns are always shown regardless
 * of scroll, so they can't share one "visible range" with the scrollable
 * body the way a single un-frozen table could. See doc/architecture.md
 * §6.1 for why this is modeled as up to four independent regions rather
 * than one merged range.
 */
/**
 * Paints cell values onto a Canvas2DLike surface, clipped to whatever is
 * currently visible per SceneSnapshot. Never
 * depends on SheetModel/TableStore directly — only a getCellValue callback
 * — so it can be unit-tested without constructing a real data layer.
 */
export class ContentLayerRenderer {
  readonly #ctx: Canvas2DLike;
  readonly #getScene: () => SceneSnapshot;
  readonly #getCellValue: (row: number, column: number) => unknown;
  readonly #painter: CellPainter;
  readonly #onError:
    | ((error: unknown, cell: { row: number; column: number }) => void)
    | undefined;
  readonly #checkSyncRenderers: boolean;
  readonly #showGridLines: boolean;
  readonly #findMergeAt: ContentLayerRendererOptions["findMergeAt"];

  constructor(options: ContentLayerRendererOptions) {
    this.#ctx = options.ctx;
    this.#getScene = options.getScene;
    this.#getCellValue = options.getCellValue;
    this.#painter = options.painter ?? defaultPainter;
    this.#onError = options.onError;
    this.#checkSyncRenderers = options.checkSyncRenderers ?? true;
    this.#showGridLines = options.showGridLines ?? true;
    this.#findMergeAt = options.findMergeAt;
  }

  paintContent(dirty: DirtyState, scene: SceneSnapshot = this.#getScene()): void {
    if (dirty.kind === "none") {
      return;
    }
    if (dirty.kind === "full") {
      // Quadrant painting only clears each individual cell's own rect, at
      // its *current* position — fine as long as content only ever grows
      // or keeps the same footprint, but a structural change that shrinks
      // total content width/height (Step 17: hideColumn()/deleteColumns())
      // leaves whatever was painted at the old, now-uncovered trailing
      // edge on screen forever otherwise. A full repaint is exactly the
      // case where clearing the whole viewport first is both correct and
      // cheap (it's about to repaint every visible cell anyway).
      if (scene.width > 0 && scene.height > 0) {
        this.#ctx.clearRect(0, 0, scene.width, scene.height);
      }
      for (const fragment of scene.getVisibleFragments()) {
        this.#paintFragment(scene, fragment);
      }
    } else {
      for (const range of dirty.ranges) {
        for (const fragment of scene.getVisibleFragmentsForRange(range)) {
          this.#paintFragment(scene, fragment);
        }
      }
    }

    this.#paintFrozenSeparators(scene);
  }

  /**
   * A heavier line marking the fixed screen boundary between the frozen
   * band and the scrollable body — always at the same pixel position
   * regardless of scroll, since that's what "frozen" means. Drawn after
   * every paint pass (not just "full") because a regions repaint can clear
   * pixels right up to that boundary without necessarily repainting it.
   */
  #paintFrozenSeparators(scene: SceneSnapshot): void {
    const frozenCorner = scene.getPanes("body").find((pane) => pane.id === "frozen-corner");
    if (!frozenCorner) {
      return;
    }

    if (frozenCorner.rows.count > 0 && frozenCorner.clipRect.height > 0) {
      const y = frozenCorner.clipRect.y + frozenCorner.clipRect.height;
      this.#ctx.fillStyle = FROZEN_SEPARATOR_COLOR;
      this.#ctx.fillRect(
        0,
        y - FROZEN_SEPARATOR_THICKNESS / 2,
        scene.width,
        FROZEN_SEPARATOR_THICKNESS
      );
    }

    if (frozenCorner.columns.count > 0 && frozenCorner.clipRect.width > 0) {
      const x = frozenCorner.clipRect.x + frozenCorner.clipRect.width;
      this.#ctx.fillStyle = FROZEN_SEPARATOR_COLOR;
      this.#ctx.fillRect(
        x - FROZEN_SEPARATOR_THICKNESS / 2,
        0,
        FROZEN_SEPARATOR_THICKNESS,
        scene.height
      );
    }
  }

  #paintFragment(scene: SceneSnapshot, fragment: SceneCellFragment): void {
    const clipRect = fragment.pane.clipRect;
    this.#ctx.save();
    this.#ctx.beginPath();
    this.#ctx.rect(clipRect.x, clipRect.y, clipRect.width, clipRect.height);
    this.#ctx.clip();

    const lastRow = fragment.rows.start + fragment.rows.count - 1;
    const lastColumn = fragment.columns.start + fragment.columns.count - 1;

    // Every (row, column) this pass has already painted as part of a merged
    // cell's single spanning rect — a fresh Set on every #paintFragment call,
    // never a field: a persistent Set would go stale (new merges created
    // after the first frame would never get re-scanned into it) and could
    // leak entries from one pane's fragment into another's.
    const mergeCoveredCells = new Set<string>();
    if (this.#findMergeAt) {
      const paintedMerges = new Set<string>();
      for (let row = fragment.rows.start; row <= lastRow; row += 1) {
        for (let column = fragment.columns.start; column <= lastColumn; column += 1) {
          const merge = this.#findMergeAt(row, column);
          if (!merge) {
            continue;
          }
          mergeCoveredCells.add(`${row},${column}`);

          const mergeKey = `${merge.startRow},${merge.startColumn}`;
          if (paintedMerges.has(mergeKey)) {
            continue;
          }
          paintedMerges.add(mergeKey);

          const mergeFragment = scene
            .getFragmentsForRange(merge)
            .find((candidate) => candidate.pane.id === fragment.pane.id);
          if (!mergeFragment) {
            continue;
          }
          this.#paintCellAt(merge.startRow, merge.startColumn, mergeFragment.rect);
        }
      }
    }

    for (let row = fragment.rows.start; row <= lastRow; row += 1) {
      for (let column = fragment.columns.start; column <= lastColumn; column += 1) {
        if (mergeCoveredCells.has(`${row},${column}`)) {
          continue;
        }
        this.#paintCell(scene, row, column, fragment);
      }
    }

    this.#ctx.restore();
  }

  #paintCell(scene: SceneSnapshot, row: number, column: number, fragment: SceneCellFragment): void {
    const rect = scene.getCellRectInPane({ row, column }, fragment.pane);
    this.#paintCellAt(row, column, rect);
  }

  #paintCellAt(row: number, column: number, rect: CellRect): void {
    if (rect.width <= 0 || rect.height <= 0) {
      return;
    }
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

}
