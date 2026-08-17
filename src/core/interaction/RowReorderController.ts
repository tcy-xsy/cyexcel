import type { SceneSnapshot } from "../scene";
import { DRAG_THRESHOLD_PX, resolveRow } from "./PointerController";
import type { SelectionSnapshot } from "./SelectionModel";

export interface RowReorderControllerOptions {
  /** Fired on every indicator change during a drag (the "insert before this row" boundary), and with `null` when the drag ends, is cancelled, or has no valid target. */
  readonly onIndicatorChange?: (row: number | null) => void;
}

export interface RowReorderIntent {
  readonly start: number;
  readonly count: number;
  readonly target: number;
  readonly structureVersion: number;
}

interface DragState {
  readonly sourceStart: number;
  readonly sourceCount: number;
  readonly startY: number;
  readonly thresholdExceeded: boolean;
  readonly structureVersion: number;
}

/** Small pixel tolerance around a row selection's top/bottom edge — same reasoning as PointerController's column-resize RESIZE_HANDLE_TOLERANCE. */
const EDGE_HANDLE_TOLERANCE = 4;

/**
 * Viewport-relative y of the boundary "before row `boundaryRow`" (so
 * `boundaryRow === rowCount` is the boundary after the last row). Frozen
 * rows — and the divider immediately below them, which is itself pinned —
 * never move with scroll; everything past that uses the same
 * scroll-subtraction `toViewportRect()` uses elsewhere. Exported so
 * `RowHeaderCanvasRenderer` can convert a committed `indicatorRow` back into
 * the same pixel position this controller used to compute it, the same way
 * Scene's selection-handle fragment is shared between
 * `FillHandleController` hit-testing and `SelectionLayerRenderer` drawing.
 */
export function getRowBoundaryY(boundaryRow: number, scene: SceneSnapshot): number {
  return scene.getRowBoundaryY(boundaryRow);
}

/**
 * Hit-tests a viewport-relative y (the same coordinate space
 * `resolveRow()`/`CyExcelCore`'s row-header handlers already use, i.e.
 * already past the headerHeight offset) against the current selection's
 * top/bottom edge. Only ever true for a row-kind selection — clicking
 * anywhere in the row-header band always produces one (see
 * `SelectionModel.selectRow()`), so a non-row-kind active selection means
 * the pointer landed somewhere the reorder handle was never meant to
 * appear.
 */
export function isPointOnRowReorderHandle(
  y: number,
  snapshot: SelectionSnapshot,
  scene: SceneSnapshot
): boolean {
  if (snapshot.kind !== "row") {
    return false;
  }
  const { startRow, endRow } = snapshot.range;
  const topY = getRowBoundaryY(startRow, scene);
  const bottomY = getRowBoundaryY(endRow + 1, scene);
  return Math.abs(y - topY) <= EDGE_HANDLE_TOLERANCE || Math.abs(y - bottomY) <= EDGE_HANDLE_TOLERANCE;
}

/**
 * Edge-handle drag-to-reorder for the row-number band: only starts when
 * pointerdown lands on the current row selection's top/bottom edge (see
 * `isPointOnRowReorderHandle`) — everywhere else in the row header stays
 * the existing drag-to-select-multiple-rows gesture, unchanged. Mirrors
 * `FillHandleController`'s pure, DOM-agnostic state machine: Core's shared
 * InputRouter supplies row-header surface points from a fresh interaction
 * frame and PointerInteractionSession owns capture/cancellation.
 * Only supports insert-style moves (matching `moveRows()`'s existing
 * splice semantics) — never an Excel-style destructive overwrite move; see
 * doc/progress.md's Step 18 entry for why that variant was explicitly
 * scoped out.
 */
export class RowReorderController {
  readonly #onIndicatorChange: (row: number | null) => void;
  #drag: DragState | null = null;
  #indicatorRow: number | null = null;

  constructor(options: RowReorderControllerOptions) {
    this.#onIndicatorChange = options.onIndicatorChange ?? ((): void => {});
  }

  get indicatorRow(): number | null {
    return this.#indicatorRow;
  }

  /** True only once the drag has moved past the jitter threshold — a bare hit-then-release never enters a "dragging" visual state. */
  get dragging(): boolean {
    return this.#drag !== null && this.#drag.thresholdExceeded;
  }

  isPointOnHandle(y: number, selection: SelectionSnapshot, scene: SceneSnapshot): boolean {
    return isPointOnRowReorderHandle(y, selection, scene);
  }

  /** Returns whether the point hit the handle and a drag was armed — callers use this to decide whether to suppress their own pointerdown handling (drag-to-select). */
  handlePointerDown(
    y: number,
    selection: SelectionSnapshot,
    scene: SceneSnapshot,
    structureVersion: number
  ): boolean {
    if (!isPointOnRowReorderHandle(y, selection, scene)) {
      return false;
    }
    const { startRow, endRow } = selection.range;
    this.#drag = {
      sourceStart: startRow,
      sourceCount: endRow - startRow + 1,
      startY: y,
      thresholdExceeded: false,
      structureVersion
    };
    return true;
  }

  handlePointerMove(y: number, scene: SceneSnapshot, structureVersion: number): void {
    if (!this.#hasCurrentStructure(structureVersion)) {
      return;
    }
    const drag = this.#drag as DragState;
    if (!drag.thresholdExceeded) {
      if (Math.abs(y - drag.startY) < DRAG_THRESHOLD_PX) {
        return;
      }
      this.#drag = { ...drag, thresholdExceeded: true };
    }

    const target = this.#resolveTarget(y, scene, drag.sourceStart, drag.sourceCount);
    this.#indicatorRow = target;
    this.#onIndicatorChange(target);
  }

  handlePointerUp(structureVersion: number): RowReorderIntent | null {
    if (!this.#hasCurrentStructure(structureVersion)) {
      return null;
    }
    const drag = this.#drag as DragState;
    const { sourceStart, sourceCount } = drag;
    const target = this.#indicatorRow;
    this.#drag = null;
    this.#indicatorRow = null;
    this.#onIndicatorChange(null);
    return target !== null && (target < sourceStart || target > sourceStart + sourceCount)
      ? Object.freeze({
          start: sourceStart,
          count: sourceCount,
          target,
          structureVersion: drag.structureVersion
        })
      : null;
  }

  /** Aborts an in-progress drag without committing — used for Esc and any other explicit cancellation. A no-op when no drag is in progress. */
  cancel(): void {
    if (!this.#drag) {
      return;
    }
    this.#drag = null;
    this.#indicatorRow = null;
    this.#onIndicatorChange(null);
  }

  /**
   * Resolves the hovered row into an "insert before this row" boundary —
   * the top half of the hovered row means "before it", the bottom half
   * means "after it" (i.e. before the next one). `null` when the pointer
   * is hovering strictly inside the block being dragged (the two endpoints
   * `sourceStart`/`sourceStart + sourceCount` are still valid boundaries —
   * they just resolve to a no-op commit, handled in handlePointerUp).
   */
  #resolveTarget(
    y: number,
    scene: SceneSnapshot,
    sourceStart: number,
    sourceCount: number
  ): number | null {
    const hoveredRow = resolveRow(y, scene);
    const rowTopY = getRowBoundaryY(hoveredRow, scene);
    // hoveredRow's own height, not the default — a row with a custom
    // height needs its actual midpoint, or the drop indicator would
    // resolve to the wrong half of a taller/shorter row.
    const midpointY = rowTopY + scene.getRowHeight(hoveredRow) / 2;
    const boundary = y < midpointY ? hoveredRow : hoveredRow + 1;

    if (boundary > sourceStart && boundary < sourceStart + sourceCount) {
      return null;
    }
    return boundary;
  }

  #hasCurrentStructure(structureVersion: number): boolean {
    if (!this.#drag) {
      return false;
    }
    if (this.#drag.structureVersion !== structureVersion) {
      this.cancel();
      return false;
    }
    return true;
  }
}
