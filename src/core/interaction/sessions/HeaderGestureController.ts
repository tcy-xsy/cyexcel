import type { SceneSnapshot, SurfaceSetSnapshot, SurfacePoint } from "../../scene";
import { clientPoint } from "../../scene";
import type {
  HeaderResizeController,
  HeaderResizeIntent,
  HeaderResizePreview
} from "../HeaderResizeController";
import type { HeaderSelectionController } from "../HeaderSelectionController";
import type { RowReorderController, RowReorderIntent } from "../RowReorderController";
import type { SelectionModel } from "../SelectionModel";
import {
  PointerInteractionSession,
  type PointerCaptureTarget,
  type PointerInteractionSessionOptions
} from "./PointerInteractionSession";

/** Distance from a row-header edge at which auto-scroll starts, in CSS px. */
const AUTO_SCROLL_EDGE_PX = 32;
/**
 * Fastest auto-scroll speed (px per animation frame), reached right at the
 * edge; scales down to 0 at AUTO_SCROLL_EDGE_PX away from it.
 */
const AUTO_SCROLL_MAX_SPEED = 16;

export type HeaderGestureKind =
  | "column-resize"
  | "column-selection"
  | "row-resize"
  | "row-reorder"
  | "row-selection";

export type HeaderGestureSurface = "column-header" | "row-header";

export interface HeaderGestureFrame {
  readonly scene: SceneSnapshot;
  readonly surfaces: SurfaceSetSnapshot;
  readonly structureVersion: number | undefined;
}

/** Anything whose CSS cursor this controller drives during a gesture. */
export interface CursorTarget {
  readonly style: { cursor: string };
}

/**
 * Everything the gesture machine needs from its host, as one named seam
 * rather than a dozen loose callbacks. Every method here is either a
 * presentation notification or a commit — the controller decides *when*, the
 * host decides *how*.
 */
export interface HeaderGestureHost {
  withFrame<TResult>(
    eventName: string,
    run: (frame: HeaderGestureFrame) => TResult
  ): TResult | undefined;
  enqueueIntent(run: () => void): void;
  /**
   * Publishes a provisional resize root. Deliberately a host call: applying a
   * preview means preparing a new Layout root, and a gesture controller that
   * reached for LayoutEngine itself would be re-creating the second coordinate
   * truth ADR 0002 removed (and would not pass the §12 import gate).
   */
  publishResizePreview(preview: HeaderResizePreview): void;
  restoreAuthoritativeScene(): void;
  invalidateSelection(): void;
  invalidateRowHeader(): void;
  /** Applies one auto-scroll step to the body scroller and resyncs presentation. */
  scrollVerticallyBy(delta: number): void;
  commitResize(intent: HeaderResizeIntent): void;
  commitReorder(intent: RowReorderIntent): void;
  updateHoverAfterGesture(event: PointerEvent, surface: HeaderGestureSurface): void;
  reportError(error: unknown, eventName: string): void;
}

export interface HeaderGestureControllerOptions {
  readonly host: HeaderGestureHost;
  readonly resizeController: HeaderResizeController;
  readonly selectionController: HeaderSelectionController;
  readonly rowReorderController: RowReorderController | undefined;
  readonly selectionModel: SelectionModel;
  readonly columnCaptureTarget: PointerCaptureTarget & CursorTarget;
  readonly rowCaptureTarget: PointerCaptureTarget & CursorTarget;
  /** Usually `document.body` — the grab cursor during a row reorder. */
  readonly bodyCursorTarget: CursorTarget;
  /** Injectable so the auto-scroll loop is testable without a real rAF clock. */
  readonly requestFrame?: (callback: () => void) => number;
  readonly cancelFrame?: (handle: number) => void;
  readonly createSession?: (options: PointerInteractionSessionOptions) => PointerInteractionSession;
}

/**
 * Owns every header pointer gesture: column/row resize, column/row header
 * selection and row reorder, plus the reorder auto-scroll loop.
 *
 * These five kinds share one state machine on purpose — pointer-move,
 * pointer-up and cancel each branch over all of them against a single
 * `{kind, structureVersion}` fact, so splitting them would mean duplicating
 * the machine rather than dividing it.
 */
export class HeaderGestureController {
  readonly #options: HeaderGestureControllerOptions;
  readonly #host: HeaderGestureHost;
  readonly #requestFrame: (callback: () => void) => number;
  readonly #cancelFrame: (handle: number) => void;

  #gesture: { readonly kind: HeaderGestureKind; readonly structureVersion: number } | null = null;
  #session: PointerInteractionSession | null = null;
  /**
   * Last pointer position seen during an armed reorder drag — the auto-scroll
   * loop re-evaluates against this every animation frame, not just on the
   * pointermove events that update it, since scrolling has to keep happening
   * while the pointer itself stays still near an edge.
   */
  #lastClientY: number | null = null;
  #autoScrollFrame: number | null = null;

  constructor(options: HeaderGestureControllerOptions) {
    this.#options = options;
    this.#host = options.host;
    this.#requestFrame = options.requestFrame ?? ((callback) => requestAnimationFrame(callback));
    this.#cancelFrame = options.cancelFrame ?? ((handle) => cancelAnimationFrame(handle));
  }

  get isActive(): boolean {
    return this.#session !== null;
  }

  get activeKind(): HeaderGestureKind | null {
    return this.#gesture?.kind ?? null;
  }

  matchesPointer(pointerId: number): boolean {
    return this.#session?.matches(pointerId) ?? false;
  }

  begin(kind: HeaderGestureKind, event: PointerEvent, structureVersion: number): void {
    const captureTarget = kind.startsWith("column-")
      ? this.#options.columnCaptureTarget
      : this.#options.rowCaptureTarget;
    const sessionOptions: PointerInteractionSessionOptions = {
      pointerId: event.pointerId,
      captureTarget,
      onMove: this.#handleMove,
      onComplete: this.#handleComplete,
      onCancel: () => {
        if (this.#session === session) {
          this.cancel(true);
        }
      },
      onCaptureError: (error, phase) => {
        this.#host.reportError(error, `interaction:header-pointer-${phase}`);
      }
    };
    const session = this.#options.createSession
      ? this.#options.createSession(sessionOptions)
      : new PointerInteractionSession(sessionOptions);
    this.#gesture = { kind, structureVersion };
    this.#session = session;
  }

  /** Records the grab cursor a reorder drag runs under. */
  markReorderGrabbing(): void {
    this.#options.bodyCursorTarget.style.cursor = "grabbing";
  }

  /** Cancels only resize gestures; ordinary scroll must not kill row auto-scroll/reorder. */
  cancelResizeOnly(publishResizePreview: boolean): void {
    if (this.#gesture?.kind === "column-resize" || this.#gesture?.kind === "row-resize") {
      this.cancel(publishResizePreview);
    }
  }

  cancel(publishResizePreview: boolean): void {
    const hadResizePreview = this.#options.resizeController.previewPublished;
    const session = this.#session;
    this.#session = null;
    this.#gesture = null;
    session?.dispose();
    this.#options.resizeController.cancel();
    this.#options.selectionController.cancel();
    this.#options.rowReorderController?.cancel();
    this.#stopAutoScroll();
    this.#options.columnCaptureTarget.style.cursor = "";
    this.#options.rowCaptureTarget.style.cursor = "";
    this.#options.bodyCursorTarget.style.cursor = "";
    if (publishResizePreview && hadResizePreview) {
      this.#host.restoreAuthoritativeScene();
    }
  }

  #handleMove = (event: PointerEvent): void => {
    const gesture = this.#gesture;
    if (!gesture) {
      return;
    }

    const handled = this.#host.withFrame("header-pointer-move", (frame) => {
      if (frame.structureVersion !== gesture.structureVersion) {
        this.cancel(true);
        return false;
      }
      const surface: HeaderGestureSurface = gesture.kind.startsWith("column-")
        ? "column-header"
        : "row-header";
      const point = frame.surfaces.clientToSurface(
        clientPoint(event.clientX, event.clientY),
        surface
      );
      if (!point) {
        return false;
      }

      if (gesture.kind === "column-resize" || gesture.kind === "row-resize") {
        const preview = this.#options.resizeController.handlePointerMove(
          point,
          frame.scene,
          gesture.structureVersion
        );
        if (preview) {
          this.#host.publishResizePreview(preview);
        }
        return true;
      }

      if (gesture.kind === "column-selection" || gesture.kind === "row-selection") {
        const action = this.#options.selectionController.move(
          point as SurfacePoint<"column-header"> | SurfacePoint<"row-header">,
          frame.scene,
          this.#options.selectionModel,
          gesture.structureVersion
        );
        if (action) {
          this.#host.invalidateSelection();
        }
        return true;
      }

      this.#lastClientY = event.clientY;
      this.#applyReorderPointerMove(frame);
      this.#startAutoScroll();
      return true;
    });
    if (handled !== true) {
      this.cancel(true);
    }
  };

  #handleComplete = (event: PointerEvent): void => {
    const gesture = this.#gesture;
    const hadResizePreview = this.#options.resizeController.previewPublished;
    this.#session = null;
    this.#gesture = null;
    if (!gesture) {
      return;
    }
    if (gesture.kind === "row-reorder") {
      this.#stopAutoScroll();
      this.#options.bodyCursorTarget.style.cursor = "";
    }

    let completed = false;
    this.#host.withFrame("header-pointer-up", (frame) => {
      if (frame.structureVersion !== gesture.structureVersion) {
        return;
      }
      completed = true;
      if (gesture.kind === "column-resize" || gesture.kind === "row-resize") {
        const intent = this.#options.resizeController.handlePointerUp(gesture.structureVersion);
        if (intent) {
          this.#host.enqueueIntent(() => this.#host.commitResize(intent));
        } else if (hadResizePreview) {
          this.#host.enqueueIntent(() => this.#host.restoreAuthoritativeScene());
        }
      } else if (gesture.kind === "column-selection" || gesture.kind === "row-selection") {
        this.#options.selectionController.up(gesture.structureVersion);
      } else {
        const intent = this.#options.rowReorderController?.handlePointerUp(
          gesture.structureVersion
        );
        if (intent) {
          this.#host.enqueueIntent(() => this.#host.commitReorder(intent));
        }
      }
    });

    if (!completed) {
      this.#options.resizeController.cancel();
      this.#options.selectionController.cancel();
      this.#options.rowReorderController?.cancel();
      if (hadResizePreview) {
        this.#host.restoreAuthoritativeScene();
      }
    }
    this.#host.invalidateRowHeader();
    this.#host.updateHoverAfterGesture(
      event,
      gesture.kind.startsWith("column-") ? "column-header" : "row-header"
    );
  };

  #applyReorderPointerMove(frame: HeaderGestureFrame): void {
    const gesture = this.#gesture;
    if (
      gesture?.kind !== "row-reorder" ||
      frame.structureVersion !== gesture.structureVersion ||
      this.#lastClientY === null
    ) {
      return;
    }
    const rowHeader = frame.surfaces.get("row-header");
    if (!rowHeader) {
      return;
    }
    const point = frame.surfaces.clientToSurface(
      clientPoint(rowHeader.clientOrigin.x, this.#lastClientY),
      "row-header"
    );
    if (!point) {
      return;
    }
    this.#options.rowReorderController?.handlePointerMove(
      point.y,
      frame.scene,
      gesture.structureVersion
    );
    this.#host.invalidateRowHeader();
  }

  /** Pure geometry: how far this frame should scroll, or null if the gesture can no longer be trusted. */
  #computeAutoScrollDelta(frame: HeaderGestureFrame, structureVersion: number): number | null {
    if (frame.structureVersion !== structureVersion) {
      return null;
    }
    const rowHeader = frame.surfaces.get("row-header");
    if (!rowHeader || this.#lastClientY === null) {
      return null;
    }
    const bodyTop = rowHeader.clientOrigin.y;
    const bodyBottom = bodyTop + rowHeader.cssSize.height;
    const y = this.#lastClientY;
    if (y < bodyTop + AUTO_SCROLL_EDGE_PX) {
      const depth = Math.min(
        AUTO_SCROLL_EDGE_PX,
        Math.max(0, bodyTop + AUTO_SCROLL_EDGE_PX - y)
      );
      return -AUTO_SCROLL_MAX_SPEED * (depth / AUTO_SCROLL_EDGE_PX);
    }
    if (y > bodyBottom - AUTO_SCROLL_EDGE_PX) {
      const depth = Math.min(
        AUTO_SCROLL_EDGE_PX,
        Math.max(0, y - (bodyBottom - AUTO_SCROLL_EDGE_PX))
      );
      return AUTO_SCROLL_MAX_SPEED * (depth / AUTO_SCROLL_EDGE_PX);
    }
    return 0;
  }

  #startAutoScroll(): void {
    if (this.#autoScrollFrame !== null) {
      return;
    }
    const tick = (): void => {
      const gesture = this.#gesture;
      if (gesture?.kind !== "row-reorder" || this.#lastClientY === null) {
        this.#autoScrollFrame = null;
        return;
      }

      const delta = this.#host.withFrame("row-reorder-auto-scroll", (frame) =>
        this.#computeAutoScrollDelta(frame, gesture.structureVersion)
      );

      if (delta === undefined || delta === null) {
        this.cancel(true);
        return;
      }
      if (delta !== 0) {
        this.#host.scrollVerticallyBy(delta);
        const moved = this.#host.withFrame("row-reorder-auto-scroll-move", (frame) => {
          if (frame.structureVersion !== gesture.structureVersion) {
            return false;
          }
          this.#applyReorderPointerMove(frame);
          return true;
        });
        if (moved !== true) {
          this.cancel(true);
          return;
        }
      }

      if (this.#gesture?.kind === "row-reorder") {
        this.#autoScrollFrame = this.#requestFrame(tick);
      }
    };
    this.#autoScrollFrame = this.#requestFrame(tick);
  }

  #stopAutoScroll(): void {
    if (this.#autoScrollFrame !== null) {
      this.#cancelFrame(this.#autoScrollFrame);
      this.#autoScrollFrame = null;
    }
    this.#lastClientY = null;
  }
}
