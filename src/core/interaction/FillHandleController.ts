import type { SceneSnapshot, SurfacePoint } from "../scene";
import { SELECTION_HANDLE_SIZE } from "../render/theme";
import type { CellRange } from "../types";
import type { SelectionSnapshot } from "./SelectionModel";

export interface FillHandleControllerOptions {
  /** Fired on every preview change during a drag, and with `null` when the drag ends or produces no extension. */
  readonly onFillPreviewChange?: (range: CellRange | null) => void;
}

/** Immutable mutation intent returned to the input owner after a valid pointerup. */
export interface FillHandleCommitIntent {
  readonly sourceRange: CellRange;
  readonly targetRange: CellRange;
  readonly structureVersion: number;
}

interface DragState {
  readonly sourceRange: CellRange;
  readonly structureVersion: number;
}

/**
 * A few extra pixels of hit area around the visually tiny (6px)
 * fill-handle square — matching PointerController's own reasoning for
 * DRAG_THRESHOLD_PX: a target this small is hard to land a real pointer on
 * exactly.
 */
const HANDLE_HIT_PADDING = 3;

/**
 * Hit-tests a point (in the same contentCanvas-relative coordinate space
 * PointerController uses) against the current selection's fill handle.
 * Exported standalone so CyExcelCore's double-click handler can reuse the
 * exact same test FillHandleController's own pointerdown handling does,
 * without needing a full drag gesture.
 */
export function isPointOnFillHandle(
  point: SurfacePoint<"body">,
  range: CellRange,
  scene: SceneSnapshot
): boolean {
  const fragment = scene.getSelectionHandleFragment(range, SELECTION_HANDLE_SIZE);
  if (!fragment) {
    return false;
  }
  const left = Math.max(fragment.rect.x - HANDLE_HIT_PADDING, fragment.clipRect.x);
  const top = Math.max(fragment.rect.y - HANDLE_HIT_PADDING, fragment.clipRect.y);
  const right = Math.min(
    fragment.rect.x + fragment.rect.width + HANDLE_HIT_PADDING,
    fragment.clipRect.x + fragment.clipRect.width
  );
  const bottom = Math.min(
    fragment.rect.y + fragment.rect.height + HANDLE_HIT_PADDING,
    fragment.clipRect.y + fragment.clipRect.height
  );
  return (
    point.x >= left &&
    point.x < right &&
    point.y >= top &&
    point.y < bottom
  );
}

/**
 * Drag-to-fill state machine: hit-tests the fill handle, tracks a drag
 * extending the selection down or right (never up/left — see
 * doc/progress.md's fill-handle patch scope notes), and reports a preview
 * range as the pointer moves plus a final commit on pointerup. Every input
 * is explicitly body-surface-local and carries the immutable
 * Scene/selection/version captured by the caller's interaction frame. The
 * controller owns gesture state only; DOM listeners, coordinate conversion,
 * pointer capture, and mutation scheduling belong to the input layer. A
 * successful pointerup returns an intent instead of committing by callback.
 */
export class FillHandleController {
  readonly #onFillPreviewChange: (range: CellRange | null) => void;
  #drag: DragState | null = null;
  #previewRange: CellRange | null = null;

  constructor(options: FillHandleControllerOptions) {
    this.#onFillPreviewChange = options.onFillPreviewChange ?? ((): void => {});
  }

  get previewRange(): CellRange | null {
    return this.#previewRange;
  }

  get dragging(): boolean {
    return this.#drag !== null;
  }

  isPointOnHandle(
    point: SurfacePoint<"body">,
    selection: SelectionSnapshot,
    scene: SceneSnapshot
  ): boolean {
    return isPointOnFillHandle(point, selection.range, scene);
  }

  /** Returns whether the point hit the handle and a drag was started — callers use this to decide whether to suppress their own pointerdown handling. */
  handlePointerDown(
    point: SurfacePoint<"body">,
    selection: SelectionSnapshot,
    scene: SceneSnapshot,
    structureVersion: number
  ): boolean {
    if (!this.isPointOnHandle(point, selection, scene)) {
      return false;
    }
    this.#drag = {
      sourceRange: Object.freeze({ ...selection.range }),
      structureVersion
    };
    return true;
  }

  handlePointerMove(
    point: SurfacePoint<"body">,
    scene: SceneSnapshot,
    structureVersion: number
  ): void {
    const drag = this.#drag;
    if (!drag) {
      return;
    }
    if (structureVersion !== drag.structureVersion) {
      this.cancel();
      return;
    }
    const { sourceRange } = drag;
    const pointerRow = scene.resolveRow(point.y);
    const pointerColumn = scene.resolveColumn(point.x);
    const rowExtension = Math.max(0, pointerRow - sourceRange.endRow);
    const columnExtension = Math.max(0, pointerColumn - sourceRange.endColumn);

    const preview: CellRange | null =
      rowExtension === 0 && columnExtension === 0
        ? null
        : rowExtension >= columnExtension
          ? { ...sourceRange, endRow: sourceRange.endRow + rowExtension }
          : { ...sourceRange, endColumn: sourceRange.endColumn + columnExtension };

    this.#previewRange = preview;
    this.#onFillPreviewChange(preview);
  }

  handlePointerUp(structureVersion: number): FillHandleCommitIntent | null {
    const drag = this.#drag;
    if (!drag) {
      return null;
    }
    if (structureVersion !== drag.structureVersion) {
      this.cancel();
      return null;
    }
    const { sourceRange } = drag;
    const target = this.#previewRange;
    this.#drag = null;
    this.#previewRange = null;
    this.#onFillPreviewChange(null);
    return target
      ? Object.freeze({
          sourceRange,
          targetRange: Object.freeze({ ...target }),
          structureVersion: drag.structureVersion
        })
      : null;
  }

  /** Discards an active drag and preview without committing. Idempotent while idle. */
  cancel(): void {
    if (!this.#drag) {
      return;
    }
    this.#drag = null;
    this.#previewRange = null;
    this.#onFillPreviewChange(null);
  }
}
