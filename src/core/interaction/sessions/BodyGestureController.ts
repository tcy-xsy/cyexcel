import type { CellRange } from "../../types";
import type { SceneSnapshot, SurfaceSetSnapshot, SurfacePoint } from "../../scene";
import { clientPoint } from "../../scene";
import type { FillHandleController } from "../FillHandleController";
import type { PointerController } from "../PointerController";
import type { SelectionSnapshot } from "../SelectionModel";
import {
  PointerInteractionSession,
  type PointerCaptureTarget,
  type PointerInteractionSessionOptions
} from "./PointerInteractionSession";

export type BodyGestureKind = "fill" | "selection";

/** The frame shape this controller reads; CyExcelCore's InteractionFrame satisfies it structurally. */
export interface BodyGestureFrame {
  readonly scene: SceneSnapshot;
  readonly surfaces: SurfaceSetSnapshot;
  readonly selection: SelectionSnapshot | undefined;
  readonly structureVersion: number | undefined;
}

/** A completed fill drag. The controller never commits it — the caller owns commands. */
export interface BodyFillCommitIntent {
  readonly sourceRange: CellRange;
  readonly targetRange: CellRange;
  readonly structureVersion: number;
}

export interface BodyGestureControllerOptions {
  readonly pointerController: PointerController;
  readonly fillHandleController: FillHandleController;
  readonly captureTarget: PointerCaptureTarget;
  /** Focused on pointer-down so keyboard input follows the grid. */
  readonly focusTarget: { focus(options?: FocusOptions): void };
  /**
   * Runs `run` against a freshly captured interaction frame, or returns
   * `undefined` when no frame could be acquired. Frame acquisition, the read
   * lease and error reporting stay with the caller.
   */
  readonly withFrame: <TResult>(
    eventName: string,
    run: (frame: BodyGestureFrame) => TResult
  ) => TResult | undefined;
  /** Called once per completed fill drag, still inside the frame dispatch. */
  readonly onFillCommit: (intent: BodyFillCommitIntent) => void;
  readonly onError: (error: unknown, eventName: string) => void;
  /** Injectable so the lifecycle can be tested without a real pointer capture. */
  readonly createSession?: (options: PointerInteractionSessionOptions) => PointerInteractionSession;
}

/**
 * Owns the body's pointer gesture: which kind is in flight, the structure
 * version it started against, and the capture session's lifetime. Routing
 * (which candidate wins a pointer-down), frame acquisition and command
 * execution deliberately stay with the caller — this class decides only what a
 * gesture *means*, never how its result is committed.
 */
export class BodyGestureController {
  readonly #options: BodyGestureControllerOptions;
  #gesture: { readonly kind: BodyGestureKind; readonly structureVersion: number } | null = null;
  #session: PointerInteractionSession | null = null;

  constructor(options: BodyGestureControllerOptions) {
    this.#options = options;
  }

  get isActive(): boolean {
    return this.#session !== null;
  }

  /** The in-flight gesture's kind, or null when idle. Diagnostic/test surface. */
  get activeKind(): BodyGestureKind | null {
    return this.#gesture?.kind ?? null;
  }

  matchesPointer(pointerId: number): boolean {
    return this.#session?.matches(pointerId) ?? false;
  }

  /**
   * Starts a fill drag if the pointer actually landed on the fill handle.
   * Returns false when it did not, so the caller can decline the input and
   * let the next candidate try.
   */
  tryBeginFill(
    point: SurfacePoint<"body">,
    event: PointerEvent,
    frame: BodyGestureFrame
  ): boolean {
    if (event.button !== 0 || !frame.selection || frame.structureVersion === undefined) {
      return false;
    }
    if (!this.#options.fillHandleController.handlePointerDown(
      point,
      frame.selection,
      frame.scene,
      frame.structureVersion
    )) {
      return false;
    }
    event.preventDefault();
    this.#options.focusTarget.focus({ preventScroll: true });
    this.#begin("fill", event, frame.structureVersion);
    return true;
  }

  /**
   * Applies a selection pointer-down. A secondary button may change the
   * selection but never owns a drag, so no capture session is opened for it.
   */
  beginSelection(
    point: SurfacePoint<"body">,
    event: PointerEvent,
    frame: BodyGestureFrame
  ): void {
    const structureVersion = frame.structureVersion;
    if (structureVersion === undefined) {
      throw new Error("Body selection requires a structure version");
    }
    this.#options.pointerController.down(point, frame.scene, structureVersion, {
      shift: event.shiftKey,
      button: event.button,
      ctrlOrMeta: event.ctrlKey || event.metaKey
    });
    this.#options.focusTarget.focus({ preventScroll: true });

    if (event.button === 0) {
      this.#begin("selection", event, structureVersion);
    } else {
      this.#options.pointerController.cancel();
    }
  }

  cancel(): void {
    const session = this.#session;
    this.#session = null;
    this.#gesture = null;
    session?.dispose();
    this.#options.fillHandleController.cancel();
    this.#options.pointerController.cancel();
  }

  #begin(kind: BodyGestureKind, event: PointerEvent, structureVersion: number): void {
    const sessionOptions: PointerInteractionSessionOptions = {
      pointerId: event.pointerId,
      captureTarget: this.#options.captureTarget,
      onMove: this.#handleMove,
      onComplete: this.#handleComplete,
      onCancel: () => {
        // A session that has already been replaced must not cancel its
        // successor's gesture.
        if (this.#session === session) {
          this.cancel();
        }
      },
      onCaptureError: (error, phase) => {
        this.#options.onError(error, `interaction:body-pointer-${phase}`);
      }
    };
    const session = this.#options.createSession
      ? this.#options.createSession(sessionOptions)
      : new PointerInteractionSession(sessionOptions);
    this.#gesture = { kind, structureVersion };
    this.#session = session;
  }

  #handleMove = (event: PointerEvent): void => {
    const gesture = this.#gesture;
    if (!gesture) {
      return;
    }
    const handled = this.#options.withFrame("body-pointer-move", (frame) => {
      if (
        frame.structureVersion === undefined ||
        frame.structureVersion !== gesture.structureVersion
      ) {
        this.cancel();
        return;
      }
      const point = frame.surfaces.clientToSurface(
        clientPoint(event.clientX, event.clientY),
        "body"
      );
      if (!point) {
        this.cancel();
        return;
      }
      if (gesture.kind === "fill") {
        this.#options.fillHandleController.handlePointerMove(
          point,
          frame.scene,
          frame.structureVersion
        );
      } else {
        this.#options.pointerController.move(point, frame.scene, frame.structureVersion);
      }
      return true;
    });
    if (handled !== true) {
      this.cancel();
    }
  };

  #handleComplete = (_event: PointerEvent): void => {
    // Cleared before dispatching: user code reached from the frame may look at
    // this controller, and must not see a gesture that has already ended.
    const gesture = this.#gesture;
    this.#session = null;
    this.#gesture = null;
    if (!gesture) {
      return;
    }
    this.#options.withFrame("body-pointer-up", (frame) => {
      if (frame.structureVersion !== gesture.structureVersion) {
        return;
      }
      if (gesture.kind === "fill") {
        const intent = this.#options.fillHandleController.handlePointerUp(
          gesture.structureVersion
        );
        if (intent) {
          this.#options.onFillCommit(intent);
        }
      } else {
        this.#options.pointerController.up(gesture.structureVersion);
      }
    });
    // Unconditional: also covers frame acquisition failures and structure
    // mismatches, where neither branch above ran.
    this.#options.fillHandleController.cancel();
    this.#options.pointerController.cancel();
  };
}
