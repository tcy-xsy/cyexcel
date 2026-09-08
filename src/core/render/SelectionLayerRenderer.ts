import type { SceneSelectionHandleFragment, SceneSnapshot } from "../scene";
import type { CellRange } from "../types";
import type { SelectionSnapshot } from "../interaction";
import type { Canvas2DLike } from "./Canvas2DLike";
import {
  SELECTION_HANDLE_SIZE,
  SELECTION_INACTIVE_ALPHA,
  SELECTION_STROKE_COLOR,
  SELECTION_STROKE_WIDTH
} from "./theme";

export interface SelectionLayerRendererOptions {
  readonly ctx: Canvas2DLike;
  readonly getScene: () => SceneSnapshot;
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
   * the solid selection box. The preview uses the same pane fragments as the
   * selection itself, so interaction direction does not leak into rendering
   * geometry assumptions.
   */
  readonly getFillPreviewRange?: () => CellRange | null;
  readonly findMergeAt?: (row: number, column: number) => CellRange | undefined;
  /**
   * Grows an arbitrary range to enclose every merge it intersects. Used for
   * the fill-preview rect, which — unlike a selection range — can be any
   * N×M shape, not just a single cell, so `#expandForMerge`'s single-cell
   * scope can't cover it.
   */
  readonly expandForMerges?: (range: CellRange) => CellRange;
}

/**
 * Draws every region in `SelectionSnapshot.ranges` (each already a plain
 * `CellRange` — `SelectionModel` collapses cell/row/column/all selection
 * kinds into one per region, so this never branches on `kind`) — see
 * `#paintRangeBox()` for the per-region frozen-aware quadrant splitting and
 * the non-active-region fade.
 */
export class SelectionLayerRenderer {
  readonly #ctx: Canvas2DLike;
  readonly #getScene: () => SceneSnapshot;
  readonly #getSelection: () => SelectionSnapshot;
  readonly #isVisible: () => boolean;
  readonly #getFillPreviewRange: () => CellRange | null;
  readonly #findMergeAt: SelectionLayerRendererOptions["findMergeAt"];
  readonly #expandForMerges: SelectionLayerRendererOptions["expandForMerges"];
  #width = 0;
  #height = 0;

  constructor(options: SelectionLayerRendererOptions) {
    this.#ctx = options.ctx;
    this.#getScene = options.getScene;
    this.#getSelection = options.getSelection;
    this.#isVisible = options.isVisible ?? (() => true);
    this.#getFillPreviewRange = options.getFillPreviewRange ?? (() => null);
    this.#findMergeAt = options.findMergeAt;
    this.#expandForMerges = options.expandForMerges;
  }

  setSize(width: number, height: number): void {
    this.#width = width;
    this.#height = height;
  }

  paintSelection(scene: SceneSnapshot = this.#getScene()): void {
    this.#ctx.clearRect(0, 0, this.#width, this.#height);
    if (!this.#isVisible()) {
      return;
    }

    const { ranges } = this.#getSelection();
    const activeIndex = ranges.length - 1;
    ranges.forEach((range, index) => {
      this.#paintRangeBox(scene, range, index === activeIndex);
    });

    const previewRange = this.#getFillPreviewRange();
    if (previewRange) {
      this.#paintFillPreview(scene, previewRange);
    }
  }

  /**
   * Strokes one region's bounding box, frozen-aware via the same up-to-four-
   * quadrant split ContentLayerRenderer uses (doc/architecture.md §6.1) — a
   * region straddling the frozen boundary is drawn as up to 4 separate
   * stroked rects, each translated per its own frozen-row/frozen-column
   * status. Only the active region gets the fill handle; non-active regions
   * (multi-region selection, Milestone 2 Step 16) are faded via
   * `globalAlpha` rather than a second color, keeping with this project's
   * restrained visual-chrome convention — and only bracketed around
   * non-active regions specifically, so a single-region (pre-multi-region)
   * selection never touches `globalAlpha` at all.
   */
  #paintRangeBox(scene: SceneSnapshot, range: CellRange, active: boolean): void {
    const expandedRange = this.#expandForMerge(range);
    const fragments = scene.getFragmentsForRange(expandedRange);

    this.#ctx.strokeStyle = SELECTION_STROKE_COLOR;
    this.#ctx.lineWidth = SELECTION_STROKE_WIDTH;
    if (!active) {
      this.#ctx.globalAlpha = SELECTION_INACTIVE_ALPHA;
    }
    for (const fragment of fragments) {
      this.#ctx.save();
      this.#ctx.beginPath();
      this.#ctx.rect(
        fragment.pane.clipRect.x,
        fragment.pane.clipRect.y,
        fragment.pane.clipRect.width,
        fragment.pane.clipRect.height
      );
      this.#ctx.clip();
      this.#ctx.strokeRect(fragment.rect.x, fragment.rect.y, fragment.rect.width, fragment.rect.height);
      this.#ctx.restore();
    }
    if (!active) {
      this.#ctx.globalAlpha = 1;
    }

    if (active && fragments.length > 0) {
      const handle = scene.getSelectionHandleFragment(expandedRange, SELECTION_HANDLE_SIZE);
      if (handle) {
        this.#paintFillHandle(handle);
      }
    }
  }

  /**
   * Widens a selection that is exactly one cell and happens to sit on a
   * merge's anchor into that merge's full range, so it gets the same
   * bounding-box treatment `getFragmentsForRange` already gives any
   * multi-cell range — without this, selecting a merged cell would draw a
   * box around only its top-left cell (design doc §5). Deliberately scoped
   * to just this 1x1-selection case: a multi-cell selection already gets a
   * correct spanning box, and overlap between a multi-cell selection and
   * several merges is out of scope (YAGNI).
   */
  #expandForMerge(range: CellRange): CellRange {
    if (range.startRow !== range.endRow || range.startColumn !== range.endColumn) {
      return range;
    }
    return this.#findMergeAt?.(range.startRow, range.startColumn) ?? range;
  }

  /** Dashed preview of the range a fill-handle drag would extend into. */
  #paintFillPreview(scene: SceneSnapshot, range: CellRange): void {
    const expandedRange = this.#expandForMerges?.(range) ?? range;
    const fragments = scene.getRangeFragments(expandedRange);
    this.#ctx.strokeStyle = SELECTION_STROKE_COLOR;
    this.#ctx.lineWidth = SELECTION_STROKE_WIDTH;
    this.#ctx.setLineDash([4, 3]);
    for (const fragment of fragments) {
      this.#ctx.save();
      this.#ctx.beginPath();
      this.#ctx.rect(
        fragment.pane.clipRect.x,
        fragment.pane.clipRect.y,
        fragment.pane.clipRect.width,
        fragment.pane.clipRect.height
      );
      this.#ctx.clip();
      this.#ctx.strokeRect(fragment.rect.x, fragment.rect.y, fragment.rect.width, fragment.rect.height);
      this.#ctx.restore();
    }
    this.#ctx.setLineDash([]);
  }

  /**
   * Purely cosmetic fill handle at the bottom-right corner of the
   * selection, matching Excel's affordance — no drag-to-fill interaction is
   * wired to it (see doc/progress.md's "不做的事" for this patch).
   */
  #paintFillHandle(fragment: SceneSelectionHandleFragment): void {
    this.#ctx.fillStyle = SELECTION_STROKE_COLOR;
    this.#ctx.save();
    this.#ctx.beginPath();
    this.#ctx.rect(
      fragment.clipRect.x,
      fragment.clipRect.y,
      fragment.clipRect.width,
      fragment.clipRect.height
    );
    this.#ctx.clip();
    this.#ctx.fillRect(fragment.rect.x, fragment.rect.y, fragment.rect.width, fragment.rect.height);
    this.#ctx.restore();
  }

}
