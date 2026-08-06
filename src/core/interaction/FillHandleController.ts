import type { LayoutEngine, ViewportState } from "../layout";
import { getSelectionBottomRightRect } from "../layout";
import { SELECTION_HANDLE_SIZE } from "../render/theme";
import type { CellRange, Disposer } from "../types";
import type { SelectionSnapshot } from "./SelectionModel";

export interface FillHandleControllerOptions {
  readonly layout: LayoutEngine;
  readonly getSelection: () => SelectionSnapshot;
  /** Fired on every preview change during a drag, and with `null` when the drag ends or produces no extension. */
  readonly onFillPreviewChange?: (range: CellRange | null) => void;
  /** Fired once, on pointerup, only when the drag actually extended the selection. */
  readonly onFillCommit: (sourceRange: CellRange, targetRange: CellRange) => void;
}

interface DragState {
  readonly sourceRange: CellRange;
}

/**
 * A few extra pixels of hit area around the visually tiny (6px)
 * fill-handle square — matching PointerController's own reasoning for
 * DRAG_THRESHOLD_PX: a target this small is hard to land a real pointer on
 * exactly.
 */
const HANDLE_HIT_PADDING = 3;

function resolveRow(y: number, layout: LayoutEngine, viewport: ViewportState): number {
  return layout.getRowAtY(Math.max(0, y + viewport.scrollTop));
}

function resolveColumn(x: number, layout: LayoutEngine, viewport: ViewportState): number {
  return layout.getColumnAtX(Math.max(0, x + viewport.scrollLeft));
}

/**
 * Hit-tests a point (in the same contentCanvas-relative coordinate space
 * PointerController uses) against the current selection's fill handle.
 * Exported standalone so CyExcelCore's double-click handler can reuse the
 * exact same test FillHandleController's own pointerdown handling does,
 * without needing a full drag gesture.
 */
export function isPointOnFillHandle(
  point: { readonly x: number; readonly y: number },
  range: CellRange,
  layout: LayoutEngine,
  viewport: ViewportState
): boolean {
  const rect = getSelectionBottomRightRect(range, layout, viewport);
  const cornerX = rect.x + rect.width;
  const cornerY = rect.y + rect.height;
  const half = SELECTION_HANDLE_SIZE / 2 + HANDLE_HIT_PADDING;
  return (
    point.x >= cornerX - half &&
    point.x <= cornerX + half &&
    point.y >= cornerY - half &&
    point.y <= cornerY + half
  );
}

/**
 * Drag-to-fill state machine: hit-tests the fill handle, tracks a drag
 * extending the selection down or right (never up/left — see
 * doc/progress.md's fill-handle patch scope notes), and reports a preview
 * range as the pointer moves plus a final commit on pointerup. Mirrors
 * PointerController's own split of pure logic vs. a real-DOM `attach()` so
 * the drag math is unit-testable without a browser.
 */
export class FillHandleController {
  readonly #layout: LayoutEngine;
  readonly #getSelection: () => SelectionSnapshot;
  readonly #onFillPreviewChange: (range: CellRange | null) => void;
  readonly #onFillCommit: (sourceRange: CellRange, targetRange: CellRange) => void;
  #viewport: ViewportState = { scrollTop: 0, scrollLeft: 0, width: 0, height: 0 };
  #drag: DragState | null = null;
  #previewRange: CellRange | null = null;

  constructor(options: FillHandleControllerOptions) {
    this.#layout = options.layout;
    this.#getSelection = options.getSelection;
    this.#onFillPreviewChange = options.onFillPreviewChange ?? ((): void => {});
    this.#onFillCommit = options.onFillCommit;
  }

  setViewport(viewport: ViewportState): void {
    this.#viewport = viewport;
  }

  get previewRange(): CellRange | null {
    return this.#previewRange;
  }

  get dragging(): boolean {
    return this.#drag !== null;
  }

  isPointOnHandle(point: { readonly x: number; readonly y: number }): boolean {
    return isPointOnFillHandle(point, this.#getSelection().range, this.#layout, this.#viewport);
  }

  /** Returns whether the point hit the handle and a drag was started — callers use this to decide whether to suppress their own pointerdown handling. */
  handlePointerDown(point: { readonly x: number; readonly y: number }): boolean {
    if (!this.isPointOnHandle(point)) {
      return false;
    }
    this.#drag = { sourceRange: this.#getSelection().range };
    return true;
  }

  handlePointerMove(point: { readonly x: number; readonly y: number }): void {
    if (!this.#drag) {
      return;
    }
    const { sourceRange } = this.#drag;
    const pointerRow = resolveRow(point.y, this.#layout, this.#viewport);
    const pointerColumn = resolveColumn(point.x, this.#layout, this.#viewport);
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

  handlePointerUp(): void {
    if (!this.#drag) {
      return;
    }
    const { sourceRange } = this.#drag;
    const target = this.#previewRange;
    this.#drag = null;
    this.#previewRange = null;
    this.#onFillPreviewChange(null);
    if (target) {
      this.#onFillCommit(sourceRange, target);
    }
  }

  /**
   * Registers on `element` (the same contentCanvas PointerController
   * attaches to) for pointerdown, and on `window` for pointermove/pointerup
   * — matching PointerController's own attach() so a drag that leaves the
   * canvas mid-gesture still tracks correctly. Must be attached BEFORE
   * PointerController's own attach() so a handle hit can
   * stopImmediatePropagation() and pre-empt PointerController's
   * click-to-select behavior for that gesture — see CyExcelCore's
   * constructor ordering.
   */
  attach(element: HTMLElement): Disposer {
    const toPoint = (event: PointerEvent): { x: number; y: number } => {
      const rect = element.getBoundingClientRect();
      return { x: event.clientX - rect.left, y: event.clientY - rect.top };
    };
    const onPointerDown = (event: PointerEvent): void => {
      if (this.handlePointerDown(toPoint(event))) {
        event.stopImmediatePropagation();
      }
    };
    const onPointerMove = (event: PointerEvent): void => {
      this.handlePointerMove(toPoint(event));
    };
    const onPointerUp = (): void => {
      this.handlePointerUp();
    };

    element.addEventListener("pointerdown", onPointerDown);
    window.addEventListener("pointermove", onPointerMove);
    window.addEventListener("pointerup", onPointerUp);

    return () => {
      element.removeEventListener("pointerdown", onPointerDown);
      window.removeEventListener("pointermove", onPointerMove);
      window.removeEventListener("pointerup", onPointerUp);
    };
  }
}
