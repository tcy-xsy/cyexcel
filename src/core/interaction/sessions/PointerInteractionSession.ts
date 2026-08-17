export type PointerInteractionSessionState =
  | "active"
  | "completed"
  | "cancelled"
  | "disposed";

export type PointerInteractionCancelReason =
  | "pointercancel"
  | "lostpointercapture"
  | "window-blur"
  | "document-hidden"
  | "escape"
  | "programmatic";

export type PointerCaptureErrorPhase = "capture" | "release";

export interface PointerCaptureTarget extends EventTarget {
  setPointerCapture(pointerId: number): void;
  hasPointerCapture(pointerId: number): boolean;
  releasePointerCapture(pointerId: number): void;
}

export interface PointerInteractionSessionOptions {
  readonly pointerId: number;
  readonly captureTarget: PointerCaptureTarget;
  /** Injectable browser targets keep the lifecycle independently testable. */
  readonly windowTarget?: EventTarget;
  readonly documentTarget?: EventTarget;
  readonly isDocumentHidden?: () => boolean;
  readonly onMove: (event: PointerEvent) => void;
  readonly onComplete: (event: PointerEvent) => void;
  readonly onCancel: (reason: PointerInteractionCancelReason, event?: Event) => void;
  readonly onCaptureError?: (error: unknown, phase: PointerCaptureErrorPhase) => void;
}

/**
 * Owns one pointer gesture's capture and temporary browser listeners.
 * Gesture-specific state and mutations deliberately remain with the caller.
 */
export class PointerInteractionSession {
  readonly #options: PointerInteractionSessionOptions;
  readonly #windowTarget: EventTarget;
  readonly #documentTarget: EventTarget;
  readonly #isDocumentHidden: () => boolean;
  readonly #abortController = new AbortController();
  #state: PointerInteractionSessionState = "active";

  constructor(options: PointerInteractionSessionOptions) {
    this.#options = options;
    this.#windowTarget = options.windowTarget ?? window;
    this.#documentTarget = options.documentTarget ?? document;
    this.#isDocumentHidden =
      options.isDocumentHidden ?? (() => document.visibilityState === "hidden");

    const listenerOptions = { signal: this.#abortController.signal };
    this.#windowTarget.addEventListener("pointermove", this.#handlePointerMove, listenerOptions);
    this.#windowTarget.addEventListener("pointerup", this.#handlePointerUp, listenerOptions);
    this.#windowTarget.addEventListener("pointercancel", this.#handlePointerCancel, listenerOptions);
    this.#windowTarget.addEventListener("blur", this.#handleWindowBlur, listenerOptions);
    this.#windowTarget.addEventListener("keydown", this.#handleKeyDown, {
      capture: true,
      signal: this.#abortController.signal
    });
    options.captureTarget.addEventListener(
      "lostpointercapture",
      this.#handleLostPointerCapture,
      listenerOptions
    );
    this.#documentTarget.addEventListener(
      "visibilitychange",
      this.#handleVisibilityChange,
      listenerOptions
    );

    try {
      options.captureTarget.setPointerCapture(options.pointerId);
    } catch (error) {
      this.#reportCaptureError(error, "capture");
    }
  }

  get pointerId(): number {
    return this.#options.pointerId;
  }

  get state(): PointerInteractionSessionState {
    return this.#state;
  }

  get signal(): AbortSignal {
    return this.#abortController.signal;
  }

  matches(pointerId: number): boolean {
    return this.#state === "active" && pointerId === this.#options.pointerId;
  }

  cancel(reason: PointerInteractionCancelReason = "programmatic", event?: Event): void {
    if (!this.#close("cancelled")) {
      return;
    }
    this.#options.onCancel(reason, event);
  }

  /** Ends the session without reporting a gesture cancellation to its owner. */
  dispose(): void {
    this.#close("disposed");
  }

  #handlePointerMove = (event: Event): void => {
    const pointerEvent = event as PointerEvent;
    if (this.matches(pointerEvent.pointerId)) {
      this.#options.onMove(pointerEvent);
    }
  };

  #handlePointerUp = (event: Event): void => {
    const pointerEvent = event as PointerEvent;
    if (!this.matches(pointerEvent.pointerId) || !this.#close("completed")) {
      return;
    }
    this.#options.onComplete(pointerEvent);
  };

  #handlePointerCancel = (event: Event): void => {
    const pointerEvent = event as PointerEvent;
    if (this.matches(pointerEvent.pointerId)) {
      this.cancel("pointercancel", pointerEvent);
    }
  };

  #handleLostPointerCapture = (event: Event): void => {
    const pointerEvent = event as PointerEvent;
    if (this.matches(pointerEvent.pointerId)) {
      this.cancel("lostpointercapture", pointerEvent);
    }
  };

  #handleWindowBlur = (event: Event): void => {
    this.cancel("window-blur", event);
  };

  #handleVisibilityChange = (event: Event): void => {
    if (this.#isDocumentHidden()) {
      this.cancel("document-hidden", event);
    }
  };

  #handleKeyDown = (event: Event): void => {
    const keyboardEvent = event as KeyboardEvent;
    if (keyboardEvent.key !== "Escape") {
      return;
    }
    keyboardEvent.preventDefault();
    this.cancel("escape", keyboardEvent);
  };

  #close(nextState: Exclude<PointerInteractionSessionState, "active">): boolean {
    if (this.#state !== "active") {
      return false;
    }
    this.#state = nextState;

    // Abort first so releasePointerCapture cannot feed lostpointercapture
    // back into an already-ending session.
    this.#abortController.abort();
    try {
      if (this.#options.captureTarget.hasPointerCapture(this.#options.pointerId)) {
        this.#options.captureTarget.releasePointerCapture(this.#options.pointerId);
      }
    } catch (error) {
      this.#reportCaptureError(error, "release");
    }
    return true;
  }

  #reportCaptureError(error: unknown, phase: PointerCaptureErrorPhase): void {
    try {
      this.#options.onCaptureError?.(error, phase);
    } catch {
      // Error observers cannot interrupt listener or capture cleanup.
    }
  }
}
