import { describe, expect, it, vi } from "vitest";

import {
  HeaderGestureController,
  PointerInteractionSession,
  type HeaderGestureFrame,
  type HeaderGestureHost,
  type HeaderGestureKind,
  type HeaderResizeController,
  type HeaderSelectionController,
  type PointerCaptureTarget,
  type RowReorderController,
  type SelectionModel
} from "../../../../src/core/interaction";

class FakeCaptureTarget extends EventTarget implements PointerCaptureTarget {
  capturedPointerId: number | null = null;
  readonly style = { cursor: "" };

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
}

function pointerEvent(type: string, overrides: Partial<PointerEvent> = {}): PointerEvent {
  const event = new Event(type, { cancelable: true });
  const defaults = { pointerId: 1, button: 0, clientX: 0, clientY: 0 };
  for (const [key, value] of Object.entries({ ...defaults, ...overrides })) {
    Object.defineProperty(event, key, { value, configurable: true });
  }
  return event as PointerEvent;
}

const ROW_HEADER_SURFACE = {
  clientOrigin: { x: 0, y: 100 },
  cssSize: { width: 40, height: 200 }
};

function createHarness(options: { readonly previewPublished?: boolean } = {}) {
  const columnCaptureTarget = new FakeCaptureTarget();
  const rowCaptureTarget = new FakeCaptureTarget();
  const bodyCursorTarget = { style: { cursor: "" } };
  const sessionWindows: EventTarget[] = [];
  const pendingFrames: Array<() => void> = [];

  const resize = {
    previewPublished: options.previewPublished ?? false,
    handlePointerMove: vi.fn(() => null as unknown),
    handlePointerUp: vi.fn(() => null as unknown),
    cancel: vi.fn()
  };
  const selection = {
    move: vi.fn(() => null as unknown),
    up: vi.fn(),
    cancel: vi.fn()
  };
  const reorder = {
    handlePointerMove: vi.fn(),
    handlePointerUp: vi.fn(() => null as unknown),
    cancel: vi.fn()
  };

  let frame: HeaderGestureFrame = {
    scene: {} as HeaderGestureFrame["scene"],
    surfaces: {
      clientToSurface: () => ({ x: 5, y: 5 }),
      get: () => ROW_HEADER_SURFACE
    } as unknown as HeaderGestureFrame["surfaces"],
    structureVersion: 3
  };
  let frameAcquirable = true;

  const host = {
    withFrame: <TResult,>(
      _eventName: string,
      run: (f: HeaderGestureFrame) => TResult
    ): TResult | undefined => (frameAcquirable ? run(frame) : undefined),
    enqueueIntent: vi.fn((run: () => void) => run()),
    publishResizePreview: vi.fn(),
    restoreAuthoritativeScene: vi.fn(),
    invalidateSelection: vi.fn(),
    invalidateRowHeader: vi.fn(),
    scrollVerticallyBy: vi.fn(),
    commitResize: vi.fn(),
    commitReorder: vi.fn(),
    updateHoverAfterGesture: vi.fn(),
    reportError: vi.fn()
  } satisfies HeaderGestureHost;

  const controller = new HeaderGestureController({
    host,
    resizeController: resize as unknown as HeaderResizeController,
    selectionController: selection as unknown as HeaderSelectionController,
    rowReorderController: reorder as unknown as RowReorderController,
    selectionModel: {} as SelectionModel,
    columnCaptureTarget,
    rowCaptureTarget,
    bodyCursorTarget,
    requestFrame: (callback) => {
      pendingFrames.push(callback);
      return pendingFrames.length;
    },
    cancelFrame: () => {
      pendingFrames.length = 0;
    },
    createSession: (sessionOptions) => {
      const windowTarget = new EventTarget();
      sessionWindows.push(windowTarget);
      return new PointerInteractionSession({
        ...sessionOptions,
        windowTarget,
        documentTarget: new EventTarget(),
        isDocumentHidden: () => false
      });
    }
  });

  return {
    controller,
    host,
    resize,
    selection,
    reorder,
    columnCaptureTarget,
    rowCaptureTarget,
    bodyCursorTarget,
    sessionWindows,
    get windowTarget(): EventTarget {
      const target = sessionWindows[sessionWindows.length - 1];
      if (!target) {
        throw new Error("No session has been opened yet");
      }
      return target;
    },
    /** Runs the single pending animation frame, if any. */
    tick: () => {
      const next = pendingFrames.shift();
      next?.();
    },
    pendingFrames,
    setFrame: (next: Partial<HeaderGestureFrame>) => {
      frame = { ...frame, ...next };
    },
    setFrameAcquirable: (value: boolean) => {
      frameAcquirable = value;
    }
  };
}

type Harness = ReturnType<typeof createHarness>;

function begin(harness: Harness, kind: HeaderGestureKind, clientY = 150): void {
  harness.controller.begin(kind, pointerEvent("pointerdown", { clientY }), 3);
}

describe("HeaderGestureController capture target", () => {
  it("captures on the column canvas for a column gesture", () => {
    const harness = createHarness();
    begin(harness, "column-resize");
    expect(harness.columnCaptureTarget.capturedPointerId).toBe(1);
    expect(harness.rowCaptureTarget.capturedPointerId).toBe(null);
  });

  it("captures on the row canvas for a row gesture", () => {
    const harness = createHarness();
    begin(harness, "row-reorder");
    expect(harness.rowCaptureTarget.capturedPointerId).toBe(1);
    expect(harness.columnCaptureTarget.capturedPointerId).toBe(null);
  });
});

describe("HeaderGestureController resize", () => {
  it("hands a preview to the host instead of touching layout itself", () => {
    const harness = createHarness();
    const preview = { axis: "column", column: 2, width: 120 };
    harness.resize.handlePointerMove.mockReturnValue(preview);
    begin(harness, "column-resize");

    harness.windowTarget.dispatchEvent(pointerEvent("pointermove"));

    expect(harness.host.publishResizePreview).toHaveBeenCalledWith(preview);
  });

  it("commits a resize intent through the host on pointer-up", () => {
    const harness = createHarness();
    const intent = { axis: "column", column: 2, width: 120 };
    harness.resize.handlePointerUp.mockReturnValue(intent);
    begin(harness, "column-resize");

    harness.windowTarget.dispatchEvent(pointerEvent("pointerup"));

    expect(harness.host.commitResize).toHaveBeenCalledWith(intent);
  });

  it("restores the authoritative Scene when a published preview produced no intent", () => {
    const harness = createHarness({ previewPublished: true });
    harness.resize.handlePointerUp.mockReturnValue(null);
    begin(harness, "column-resize");

    harness.windowTarget.dispatchEvent(pointerEvent("pointerup"));

    expect(harness.host.commitResize).not.toHaveBeenCalled();
    expect(harness.host.restoreAuthoritativeScene).toHaveBeenCalled();
  });

  it("cancelResizeOnly leaves a reorder gesture running", () => {
    const harness = createHarness();
    begin(harness, "row-reorder");

    harness.controller.cancelResizeOnly(true);

    expect(harness.controller.isActive).toBe(true);
    expect(harness.controller.activeKind).toBe("row-reorder");
  });

  it("cancelResizeOnly does cancel a resize gesture", () => {
    const harness = createHarness();
    begin(harness, "row-resize");

    harness.controller.cancelResizeOnly(true);

    expect(harness.controller.isActive).toBe(false);
  });
});

describe("HeaderGestureController selection", () => {
  it("invalidates selection only when the sub-controller reports an action", () => {
    const harness = createHarness();
    begin(harness, "column-selection");

    harness.selection.move.mockReturnValue(null);
    harness.windowTarget.dispatchEvent(pointerEvent("pointermove"));
    expect(harness.host.invalidateSelection).not.toHaveBeenCalled();

    harness.selection.move.mockReturnValue({});
    harness.windowTarget.dispatchEvent(pointerEvent("pointermove"));
    expect(harness.host.invalidateSelection).toHaveBeenCalledTimes(1);
  });

  it("ends the selection on pointer-up", () => {
    const harness = createHarness();
    begin(harness, "row-selection");

    harness.windowTarget.dispatchEvent(pointerEvent("pointerup"));

    expect(harness.selection.up).toHaveBeenCalledWith(3);
    expect(harness.controller.isActive).toBe(false);
  });
});

describe("HeaderGestureController row reorder auto-scroll", () => {
  it("scrolls up near the top edge and keeps ticking", () => {
    const harness = createHarness();
    begin(harness, "row-reorder");

    // clientY 110 is 10px into the 32px top edge band starting at y=100.
    harness.windowTarget.dispatchEvent(pointerEvent("pointermove", { clientY: 110 }));
    harness.tick();

    const delta = harness.host.scrollVerticallyBy.mock.calls[0]?.[0] as number;
    expect(delta).toBeLessThan(0);
    expect(harness.pendingFrames.length).toBe(1);
  });

  it("does not scroll when the pointer sits away from both edges", () => {
    const harness = createHarness();
    begin(harness, "row-reorder");

    // clientY 200 is the middle of the 100..300 band.
    harness.windowTarget.dispatchEvent(pointerEvent("pointermove", { clientY: 200 }));
    harness.tick();

    expect(harness.host.scrollVerticallyBy).not.toHaveBeenCalled();
  });

  it("scrolls down near the bottom edge", () => {
    const harness = createHarness();
    begin(harness, "row-reorder");

    harness.windowTarget.dispatchEvent(pointerEvent("pointermove", { clientY: 295 }));
    harness.tick();

    expect(harness.host.scrollVerticallyBy.mock.calls[0]?.[0]).toBeGreaterThan(0);
  });

  it("cancels the gesture when the structure changes mid auto-scroll", () => {
    const harness = createHarness();
    begin(harness, "row-reorder");
    harness.windowTarget.dispatchEvent(pointerEvent("pointermove", { clientY: 110 }));

    harness.setFrame({ structureVersion: 4 });
    harness.tick();

    expect(harness.controller.isActive).toBe(false);
    expect(harness.reorder.cancel).toHaveBeenCalled();
  });

  it("stops the loop and clears the grab cursor once the drag completes", () => {
    const harness = createHarness();
    harness.controller.markReorderGrabbing();
    begin(harness, "row-reorder");
    harness.windowTarget.dispatchEvent(pointerEvent("pointermove", { clientY: 110 }));
    expect(harness.bodyCursorTarget.style.cursor).toBe("grabbing");

    harness.windowTarget.dispatchEvent(pointerEvent("pointerup"));

    expect(harness.bodyCursorTarget.style.cursor).toBe("");
    expect(harness.pendingFrames.length).toBe(0);
  });

  it("commits a reorder intent through the host", () => {
    const harness = createHarness();
    const intent = { start: 1, count: 2, target: 5, structureVersion: 3 };
    harness.reorder.handlePointerUp.mockReturnValue(intent);
    begin(harness, "row-reorder");

    harness.windowTarget.dispatchEvent(pointerEvent("pointerup"));

    expect(harness.host.commitReorder).toHaveBeenCalledWith(intent);
  });
});

describe("HeaderGestureController lifecycle", () => {
  it("cancels when the structure moved before a pointer move", () => {
    const harness = createHarness();
    begin(harness, "column-selection");
    harness.setFrame({ structureVersion: 9 });

    harness.windowTarget.dispatchEvent(pointerEvent("pointermove"));

    expect(harness.controller.isActive).toBe(false);
  });

  it("cancels every sub-controller when no frame could be acquired at pointer-up", () => {
    const harness = createHarness();
    begin(harness, "column-selection");
    harness.setFrameAcquirable(false);

    harness.windowTarget.dispatchEvent(pointerEvent("pointerup"));

    expect(harness.resize.cancel).toHaveBeenCalled();
    expect(harness.selection.cancel).toHaveBeenCalled();
    expect(harness.reorder.cancel).toHaveBeenCalled();
  });

  it("clears all three cursors on cancel", () => {
    const harness = createHarness();
    harness.controller.markReorderGrabbing();
    begin(harness, "row-reorder");
    harness.columnCaptureTarget.style.cursor = "col-resize";
    harness.rowCaptureTarget.style.cursor = "row-resize";

    harness.controller.cancel(false);

    expect(harness.columnCaptureTarget.style.cursor).toBe("");
    expect(harness.rowCaptureTarget.style.cursor).toBe("");
    expect(harness.bodyCursorTarget.style.cursor).toBe("");
  });

  it("a replaced session's cancellation does not kill its successor", () => {
    const harness = createHarness();
    begin(harness, "column-selection");
    const staleWindow = harness.windowTarget;

    begin(harness, "column-selection");
    expect(harness.sessionWindows).toHaveLength(2);

    const escape = new Event("keydown", { cancelable: true });
    Object.defineProperty(escape, "key", { value: "Escape" });
    staleWindow.dispatchEvent(escape);

    expect(harness.controller.isActive).toBe(true);
  });

  it("reports hover back to the host after the gesture ends", () => {
    const harness = createHarness();
    begin(harness, "column-selection");

    harness.windowTarget.dispatchEvent(pointerEvent("pointerup"));

    expect(harness.host.updateHoverAfterGesture).toHaveBeenCalledWith(
      expect.anything(),
      "column-header"
    );
  });
});
