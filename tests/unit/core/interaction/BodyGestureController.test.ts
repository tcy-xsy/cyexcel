import { beforeEach, describe, expect, it, vi } from "vitest";

import {
  BodyGestureController,
  PointerInteractionSession,
  type BodyGestureFrame,
  type FillHandleController,
  type PointerCaptureTarget,
  type PointerController
} from "../../../../src/core/interaction";
import type { SurfacePoint } from "../../../../src/core/scene";

class FakeCaptureTarget extends EventTarget implements PointerCaptureTarget {
  capturedPointerId: number | null = null;
  focusCount = 0;

  setPointerCapture(pointerId: number): void {
    this.capturedPointerId = pointerId;
  }

  hasPointerCapture(pointerId: number): boolean {
    return this.capturedPointerId === pointerId;
  }

  releasePointerCapture(pointerId: number): void {
    if (this.capturedPointerId === pointerId) {
      this.capturedPointerId = null;
    }
  }

  focus(): void {
    this.focusCount += 1;
  }
}

function pointerEvent(type: string, overrides: Partial<PointerEvent> = {}): PointerEvent {
  const event = new Event(type, { cancelable: true });
  const defaults = { pointerId: 1, button: 0, clientX: 0, clientY: 0, shiftKey: false, ctrlKey: false, metaKey: false };
  for (const [key, value] of Object.entries({ ...defaults, ...overrides })) {
    Object.defineProperty(event, key, { value, configurable: true });
  }
  return event as PointerEvent;
}

const BODY_POINT = { x: 5, y: 5 } as SurfacePoint<"body">;

function createHarness(options: { readonly fillHit?: boolean } = {}) {
  const captureTarget = new FakeCaptureTarget();
  const documentTarget = new EventTarget();
  // One window target per session, so a test can fire an event at exactly one
  // session — a real stale session keeps listening to the same window as its
  // successor, and the point here is to isolate the stale one's callback.
  const sessionWindows: EventTarget[] = [];

  const fill = {
    handlePointerDown: vi.fn(() => options.fillHit ?? true),
    handlePointerMove: vi.fn(),
    handlePointerUp: vi.fn(() => null as unknown),
    cancel: vi.fn()
  };
  const pointer = {
    down: vi.fn(),
    move: vi.fn(),
    up: vi.fn(),
    cancel: vi.fn()
  };
  const onFillCommit = vi.fn();
  const onError = vi.fn();

  // A frame that is always acquirable; individual tests override it.
  let frame: BodyGestureFrame = {
    scene: { } as BodyGestureFrame["scene"],
    surfaces: {
      clientToSurface: () => BODY_POINT
    } as unknown as BodyGestureFrame["surfaces"],
    selection: { } as BodyGestureFrame["selection"],
    structureVersion: 7
  };
  let frameAcquirable = true;
  // Plain generic function rather than vi.fn(): a mock erases the call-site
  // generic, and this harness never asserts on the calls themselves.
  const withFrame = <TResult,>(
    _eventName: string,
    run: (f: BodyGestureFrame) => TResult
  ): TResult | undefined => (frameAcquirable ? run(frame) : undefined);

  const controller = new BodyGestureController({
    pointerController: pointer as unknown as PointerController,
    fillHandleController: fill as unknown as FillHandleController,
    captureTarget,
    focusTarget: captureTarget,
    withFrame,
    onFillCommit,
    onError,
    createSession: (sessionOptions) => {
      const windowTarget = new EventTarget();
      sessionWindows.push(windowTarget);
      return new PointerInteractionSession({
        ...sessionOptions,
        windowTarget,
        documentTarget,
        isDocumentHidden: () => false
      });
    }
  });

  return {
    controller,
    captureTarget,
    sessionWindows,
    /** The window target of the most recently opened session. */
    get windowTarget(): EventTarget {
      const target = sessionWindows[sessionWindows.length - 1];
      if (!target) {
        throw new Error("No session has been opened yet");
      }
      return target;
    },
    fill,
    pointer,
    onFillCommit,
    onError,
    setFrame: (next: Partial<BodyGestureFrame>) => {
      frame = { ...frame, ...next };
    },
    setFrameAcquirable: (value: boolean) => {
      frameAcquirable = value;
    }
  };
}

type Harness = ReturnType<typeof createHarness>;

function beginSelectionDrag(harness: Harness): void {
  harness.controller.beginSelection(BODY_POINT, pointerEvent("pointerdown"), {
    scene: {} as BodyGestureFrame["scene"],
    surfaces: {} as BodyGestureFrame["surfaces"],
    selection: undefined,
    structureVersion: 7
  });
}

describe("BodyGestureController.beginSelection", () => {
  let harness: Harness;
  beforeEach(() => {
    harness = createHarness();
  });

  it("applies the pointer-down and opens a capture session for the primary button", () => {
    beginSelectionDrag(harness);

    expect(harness.pointer.down).toHaveBeenCalledTimes(1);
    expect(harness.captureTarget.focusCount).toBe(1);
    expect(harness.controller.isActive).toBe(true);
    expect(harness.controller.activeKind).toBe("selection");
  });

  it("lets a secondary button change the selection but never own a drag", () => {
    harness.controller.beginSelection(BODY_POINT, pointerEvent("pointerdown", { button: 2 }), {
      scene: {} as BodyGestureFrame["scene"],
      surfaces: {} as BodyGestureFrame["surfaces"],
      selection: undefined,
      structureVersion: 7
    });

    expect(harness.pointer.down).toHaveBeenCalledTimes(1);
    expect(harness.pointer.cancel).toHaveBeenCalledTimes(1);
    expect(harness.controller.isActive).toBe(false);
  });

  it("refuses to start without a structure version", () => {
    expect(() =>
      harness.controller.beginSelection(BODY_POINT, pointerEvent("pointerdown"), {
        scene: {} as BodyGestureFrame["scene"],
        surfaces: {} as BodyGestureFrame["surfaces"],
        selection: undefined,
        structureVersion: undefined
      })
    ).toThrow(/structure version/);
  });
});

describe("BodyGestureController.tryBeginFill", () => {
  it("starts a fill gesture when the handle was hit", () => {
    const harness = createHarness({ fillHit: true });
    const event = pointerEvent("pointerdown");
    const began = harness.controller.tryBeginFill(BODY_POINT, event, {
      scene: {} as BodyGestureFrame["scene"],
      surfaces: {} as BodyGestureFrame["surfaces"],
      selection: {} as BodyGestureFrame["selection"],
      structureVersion: 7
    });

    expect(began).toBe(true);
    expect(harness.controller.activeKind).toBe("fill");
    expect(event.defaultPrevented).toBe(true);
  });

  it("declines when the handle was not hit, leaving no gesture behind", () => {
    const harness = createHarness({ fillHit: false });
    const began = harness.controller.tryBeginFill(BODY_POINT, pointerEvent("pointerdown"), {
      scene: {} as BodyGestureFrame["scene"],
      surfaces: {} as BodyGestureFrame["surfaces"],
      selection: {} as BodyGestureFrame["selection"],
      structureVersion: 7
    });

    expect(began).toBe(false);
    expect(harness.controller.isActive).toBe(false);
  });

  it("declines a secondary button without consulting the fill handle", () => {
    const harness = createHarness({ fillHit: true });
    const began = harness.controller.tryBeginFill(BODY_POINT, pointerEvent("pointerdown", { button: 2 }), {
      scene: {} as BodyGestureFrame["scene"],
      surfaces: {} as BodyGestureFrame["surfaces"],
      selection: {} as BodyGestureFrame["selection"],
      structureVersion: 7
    });

    expect(began).toBe(false);
    expect(harness.fill.handlePointerDown).not.toHaveBeenCalled();
  });
});

describe("BodyGestureController pointer move", () => {
  it("routes a selection drag to the pointer controller", () => {
    const harness = createHarness();
    beginSelectionDrag(harness);

    harness.windowTarget.dispatchEvent(pointerEvent("pointermove"));

    expect(harness.pointer.move).toHaveBeenCalledTimes(1);
    expect(harness.fill.handlePointerMove).not.toHaveBeenCalled();
    expect(harness.controller.isActive).toBe(true);
  });

  it("cancels the gesture when the structure changed underneath it", () => {
    const harness = createHarness();
    beginSelectionDrag(harness);
    harness.setFrame({ structureVersion: 8 });

    harness.windowTarget.dispatchEvent(pointerEvent("pointermove"));

    expect(harness.pointer.move).not.toHaveBeenCalled();
    expect(harness.controller.isActive).toBe(false);
    expect(harness.pointer.cancel).toHaveBeenCalled();
  });

  it("cancels the gesture when the point falls outside the body surface", () => {
    const harness = createHarness();
    beginSelectionDrag(harness);
    harness.setFrame({
      surfaces: { clientToSurface: () => undefined } as unknown as BodyGestureFrame["surfaces"]
    });

    harness.windowTarget.dispatchEvent(pointerEvent("pointermove"));

    expect(harness.controller.isActive).toBe(false);
  });

  it("cancels the gesture when no interaction frame could be acquired", () => {
    const harness = createHarness();
    beginSelectionDrag(harness);
    harness.setFrameAcquirable(false);

    harness.windowTarget.dispatchEvent(pointerEvent("pointermove"));

    expect(harness.controller.isActive).toBe(false);
  });
});

describe("BodyGestureController pointer up", () => {
  it("commits a selection drag and clears the gesture", () => {
    const harness = createHarness();
    beginSelectionDrag(harness);

    harness.windowTarget.dispatchEvent(pointerEvent("pointerup"));

    expect(harness.pointer.up).toHaveBeenCalledWith(7);
    expect(harness.controller.isActive).toBe(false);
    expect(harness.controller.activeKind).toBe(null);
  });

  it("hands a completed fill drag to the caller instead of committing it", () => {
    const harness = createHarness({ fillHit: true });
    const intent = { sourceRange: {}, targetRange: {}, structureVersion: 7 };
    harness.fill.handlePointerUp.mockReturnValue(intent);
    harness.controller.tryBeginFill(BODY_POINT, pointerEvent("pointerdown"), {
      scene: {} as BodyGestureFrame["scene"],
      surfaces: {} as BodyGestureFrame["surfaces"],
      selection: {} as BodyGestureFrame["selection"],
      structureVersion: 7
    });

    harness.windowTarget.dispatchEvent(pointerEvent("pointerup"));

    expect(harness.onFillCommit).toHaveBeenCalledWith(intent);
    expect(harness.pointer.up).not.toHaveBeenCalled();
  });

  it("still cancels both sub-controllers when the structure moved before pointer-up", () => {
    const harness = createHarness();
    beginSelectionDrag(harness);
    harness.setFrame({ structureVersion: 9 });

    harness.windowTarget.dispatchEvent(pointerEvent("pointerup"));

    expect(harness.pointer.up).not.toHaveBeenCalled();
    expect(harness.pointer.cancel).toHaveBeenCalled();
    expect(harness.fill.cancel).toHaveBeenCalled();
  });

  it("still cancels both sub-controllers when the frame could not be acquired", () => {
    const harness = createHarness();
    beginSelectionDrag(harness);
    harness.setFrameAcquirable(false);

    harness.windowTarget.dispatchEvent(pointerEvent("pointerup"));

    expect(harness.pointer.cancel).toHaveBeenCalled();
    expect(harness.fill.cancel).toHaveBeenCalled();
  });
});

describe("BodyGestureController session identity", () => {
  it("a replaced session's cancellation does not kill its successor", () => {
    const harness = createHarness();
    beginSelectionDrag(harness);
    const staleWindow = harness.windowTarget;

    // Start a second gesture without cancelling the first; the first session
    // is now stale but still holds live listeners.
    beginSelectionDrag(harness);
    expect(harness.sessionWindows).toHaveLength(2);
    expect(harness.controller.isActive).toBe(true);

    // The stale session's Escape cancellation must not touch the new gesture.
    const escape = new Event("keydown", { cancelable: true });
    Object.defineProperty(escape, "key", { value: "Escape" });
    staleWindow.dispatchEvent(escape);

    expect(harness.controller.isActive).toBe(true);
  });

  it("matchesPointer only recognises the active pointer", () => {
    const harness = createHarness();
    harness.controller.beginSelection(BODY_POINT, pointerEvent("pointerdown", { pointerId: 4 }), {
      scene: {} as BodyGestureFrame["scene"],
      surfaces: {} as BodyGestureFrame["surfaces"],
      selection: undefined,
      structureVersion: 7
    });

    expect(harness.controller.matchesPointer(4)).toBe(true);
    expect(harness.controller.matchesPointer(5)).toBe(false);
  });

  it("cancel() releases capture and stops both sub-controllers", () => {
    const harness = createHarness();
    beginSelectionDrag(harness);

    harness.controller.cancel();

    expect(harness.controller.isActive).toBe(false);
    expect(harness.captureTarget.capturedPointerId).toBe(null);
    expect(harness.pointer.cancel).toHaveBeenCalled();
    expect(harness.fill.cancel).toHaveBeenCalled();
  });
});
