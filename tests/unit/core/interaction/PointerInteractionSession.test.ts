import { describe, expect, it, vi } from "vitest";

import {
  PointerInteractionSession,
  type PointerCaptureErrorPhase,
  type PointerCaptureTarget,
  type PointerInteractionCancelReason
} from "../../../../src/core/interaction";

class FakeCaptureTarget extends EventTarget implements PointerCaptureTarget {
  capturedPointerId: number | null = null;
  captureError: unknown;
  releaseError: unknown;

  setPointerCapture(pointerId: number): void {
    if (this.captureError !== undefined) {
      throw this.captureError;
    }
    this.capturedPointerId = pointerId;
  }

  hasPointerCapture(pointerId: number): boolean {
    return this.capturedPointerId === pointerId;
  }

  releasePointerCapture(pointerId: number): void {
    if (this.releaseError !== undefined) {
      throw this.releaseError;
    }
    if (this.capturedPointerId === pointerId) {
      this.capturedPointerId = null;
    }
  }
}

function pointerEvent(type: string, pointerId: number): PointerEvent {
  const event = new Event(type);
  Object.defineProperty(event, "pointerId", { value: pointerId });
  return event as PointerEvent;
}

function keyboardEvent(key: string): KeyboardEvent {
  const event = new Event("keydown", { cancelable: true });
  Object.defineProperty(event, "key", { value: key });
  return event as KeyboardEvent;
}

interface SessionHarness {
  readonly session: PointerInteractionSession;
  readonly captureTarget: FakeCaptureTarget;
  readonly windowTarget: EventTarget;
  readonly documentTarget: EventTarget;
  readonly moves: number[];
  readonly completions: number[];
  readonly cancellations: PointerInteractionCancelReason[];
  readonly captureErrors: Array<{ readonly error: unknown; readonly phase: PointerCaptureErrorPhase }>;
  setHidden(hidden: boolean): void;
}

function createHarness(captureTarget = new FakeCaptureTarget()): SessionHarness {
  const windowTarget = new EventTarget();
  const documentTarget = new EventTarget();
  const moves: number[] = [];
  const completions: number[] = [];
  const cancellations: PointerInteractionCancelReason[] = [];
  const captureErrors: Array<{ readonly error: unknown; readonly phase: PointerCaptureErrorPhase }> = [];
  let hidden = false;
  const session = new PointerInteractionSession({
    pointerId: 7,
    captureTarget,
    windowTarget,
    documentTarget,
    isDocumentHidden: () => hidden,
    onMove: event => moves.push(event.pointerId),
    onComplete: event => completions.push(event.pointerId),
    onCancel: reason => cancellations.push(reason),
    onCaptureError: (error, phase) => captureErrors.push({ error, phase })
  });

  return {
    session,
    captureTarget,
    windowTarget,
    documentTarget,
    moves,
    completions,
    cancellations,
    captureErrors,
    setHidden(value) {
      hidden = value;
    }
  };
}

describe("PointerInteractionSession", () => {
  it("captures one pointer and ignores move/up events from other pointers", () => {
    const harness = createHarness();

    expect(harness.captureTarget.capturedPointerId).toBe(7);
    harness.windowTarget.dispatchEvent(pointerEvent("pointermove", 8));
    harness.windowTarget.dispatchEvent(pointerEvent("pointerup", 8));
    harness.windowTarget.dispatchEvent(pointerEvent("pointermove", 7));
    harness.windowTarget.dispatchEvent(pointerEvent("pointerup", 7));

    expect(harness.moves).toEqual([7]);
    expect(harness.completions).toEqual([7]);
    expect(harness.session.state).toBe("completed");
    expect(harness.session.signal.aborted).toBe(true);
    expect(harness.captureTarget.capturedPointerId).toBeNull();
  });

  it.each([
    ["pointercancel", (harness: SessionHarness) =>
      harness.windowTarget.dispatchEvent(pointerEvent("pointercancel", 7))],
    ["lostpointercapture", (harness: SessionHarness) =>
      harness.captureTarget.dispatchEvent(pointerEvent("lostpointercapture", 7))],
    ["window-blur", (harness: SessionHarness) =>
      harness.windowTarget.dispatchEvent(new Event("blur"))],
    ["document-hidden", (harness: SessionHarness) => {
      harness.setHidden(true);
      harness.documentTarget.dispatchEvent(new Event("visibilitychange"));
    }],
    ["escape", (harness: SessionHarness) =>
      harness.windowTarget.dispatchEvent(keyboardEvent("Escape"))]
  ] as const)("cancels on %s and releases all session resources", (reason, cancel) => {
    const harness = createHarness();

    cancel(harness);
    harness.windowTarget.dispatchEvent(pointerEvent("pointermove", 7));
    harness.windowTarget.dispatchEvent(pointerEvent("pointerup", 7));

    expect(harness.cancellations).toEqual([reason]);
    expect(harness.moves).toEqual([]);
    expect(harness.completions).toEqual([]);
    expect(harness.session.state).toBe("cancelled");
    expect(harness.session.signal.aborted).toBe(true);
    expect(harness.captureTarget.capturedPointerId).toBeNull();
  });

  it("keeps capture failure non-fatal so window events can complete the session", () => {
    const captureTarget = new FakeCaptureTarget();
    const captureError = new Error("capture failed");
    captureTarget.captureError = captureError;
    const harness = createHarness(captureTarget);

    harness.windowTarget.dispatchEvent(pointerEvent("pointerup", 7));

    expect(harness.completions).toEqual([7]);
    expect(harness.captureErrors).toEqual([{ error: captureError, phase: "capture" }]);
  });

  it("isolates release and error-observer failures while still aborting listeners", () => {
    const captureTarget = new FakeCaptureTarget();
    captureTarget.releaseError = new Error("release failed");
    const onComplete = vi.fn();
    const windowTarget = new EventTarget();
    const session = new PointerInteractionSession({
      pointerId: 7,
      captureTarget,
      windowTarget,
      documentTarget: new EventTarget(),
      onMove: () => {},
      onComplete,
      onCancel: () => {},
      onCaptureError: () => {
        throw new Error("observer failed");
      }
    });

    expect(() => windowTarget.dispatchEvent(pointerEvent("pointerup", 7))).not.toThrow();
    expect(onComplete).toHaveBeenCalledOnce();
    expect(session.signal.aborted).toBe(true);
  });

  it("dispose is idempotent and does not report a cancellation", () => {
    const harness = createHarness();

    harness.session.dispose();
    harness.session.dispose();

    expect(harness.session.state).toBe("disposed");
    expect(harness.cancellations).toEqual([]);
    expect(harness.captureTarget.capturedPointerId).toBeNull();
  });
});
