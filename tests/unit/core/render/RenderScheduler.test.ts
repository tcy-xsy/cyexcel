import { describe, expect, it, vi } from "vitest";

import { RevisionGate } from "../../../../src/core/model";
import { RenderScheduler } from "../../../../src/core/render";
import type {
  DirtyState,
  FrameScheduler,
  RenderSchedulerOptions,
  RenderTarget
} from "../../../../src/core/render";
import type { RenderFrame } from "../../../../src/core/render/RenderScheduler";
import { compileDirtySet } from "../../../../src/core/data";

interface SceneContract {
  readonly modelVersion: number;
}

const PRESENTATION_DIRTY = {
  domCells: true,
  columnHeader: true,
  rowHeader: true,
  domHeaders: true
} as const;

const DEFAULT_SCENE: SceneContract = Object.freeze({ modelVersion: 0 });

type TestSchedulerOptions = Omit<RenderSchedulerOptions<SceneContract>, "revisionGate" | "getScene"> & {
  readonly revisionGate?: RevisionGate;
  readonly getScene?: () => SceneContract;
};

function createRenderScheduler(
  target: RenderTarget<SceneContract>,
  options: TestSchedulerOptions
): RenderScheduler<SceneContract> {
  return new RenderScheduler(target, {
    ...options,
    revisionGate: options.revisionGate ?? new RevisionGate(),
    getScene: options.getScene ?? (() => DEFAULT_SCENE)
  });
}

function createManualFrameScheduler() {
  let nextHandle = 1;
  const pending = new Map<number, () => void>();

  const frameScheduler: FrameScheduler = {
    requestFrame: (callback) => {
      const handle = nextHandle;
      nextHandle += 1;
      pending.set(handle, callback);
      return handle;
    },
    cancelFrame: (handle) => {
      pending.delete(handle);
    }
  };

  return {
    frameScheduler,
    requestFrameCallCount: () => nextHandle - 1,
    flush(): void {
      const callbacks = [...pending.values()];
      pending.clear();
      for (const callback of callbacks) {
        callback();
      }
    }
  };
}

function createCapturingTarget(): RenderTarget<SceneContract> & {
  lastDirty: DirtyState | undefined;
  paintContentCalls: number;
  paintSelectionCalls: number;
} {
  return {
    lastDirty: undefined,
    paintContentCalls: 0,
    paintSelectionCalls: 0,
    paintPresentation() {},
    paintContent(frame) {
      this.lastDirty = frame.dirty.content;
      this.paintContentCalls += 1;
    },
    paintSelection() {
      this.paintSelectionCalls += 1;
    }
  };
}

describe("RenderScheduler frame coalescing", () => {
  it("schedules exactly one frame for multiple invalidations in the same tick", () => {
    const { frameScheduler, requestFrameCallCount, flush } = createManualFrameScheduler();
    const target = createCapturingTarget();
    const scheduler = createRenderScheduler(target, { frameScheduler });

    scheduler.invalidateDirtySet(compileDirtySet([{ type: "cell-value", row: 0, column: 0 }]));
    scheduler.invalidateDirtySet(compileDirtySet([{ type: "cell-value", row: 1, column: 1 }]));
    scheduler.invalidateAll();

    expect(requestFrameCallCount()).toBe(1);

    flush();
    expect(target.paintContentCalls).toBe(1);
  });
});

describe("RenderScheduler dirty-region merging", () => {
  it("invalidates the given range for a range-value patch", () => {
    const { frameScheduler, flush } = createManualFrameScheduler();
    const target = createCapturingTarget();
    const scheduler = createRenderScheduler(target, { frameScheduler });

    scheduler.invalidateDirtySet(compileDirtySet([{
      type: "range-value",
      range: { startRow: 1, startColumn: 1, endRow: 3, endColumn: 3 }
    }]));
    flush();

    expect(target.lastDirty).toEqual({
      kind: "regions",
      ranges: [{ startRow: 1, startColumn: 1, endRow: 3, endColumn: 3 }]
    });
  });

  it("invalidates every shifted row carried by a row-layout patch", () => {
    const { frameScheduler, flush } = createManualFrameScheduler();
    const target = createCapturingTarget();
    const scheduler = createRenderScheduler(target, { frameScheduler });

    scheduler.invalidateDirtySet(compileDirtySet([{
      type: "row-layout",
      rows: [2, 5],
      range: { startRow: 2, startColumn: 0, endRow: 9, endColumn: 3 }
    }]));
    flush();

    expect(target.lastDirty).toEqual({
      kind: "regions",
      ranges: [{ startRow: 2, startColumn: 0, endRow: 9, endColumn: 3 }]
    });
  });

  it("stays full and keeps scheduling a frame when a region is added after a full invalidation", () => {
    const { frameScheduler, requestFrameCallCount, flush } = createManualFrameScheduler();
    const target = createCapturingTarget();
    const scheduler = createRenderScheduler(target, { frameScheduler });

    scheduler.invalidateAll();
    scheduler.invalidateDirtySet(compileDirtySet([{ type: "cell-value", row: 0, column: 0 }]));

    expect(requestFrameCallCount()).toBe(1);
    flush();

    expect(target.lastDirty).toEqual({ kind: "full" });
  });

  it("merges adjacent cell patches into a single region", () => {
    const { frameScheduler, flush } = createManualFrameScheduler();
    const target = createCapturingTarget();
    const scheduler = createRenderScheduler(target, { frameScheduler });

    scheduler.invalidateDirtySet(compileDirtySet([{ type: "cell-value", row: 0, column: 0 }]));
    scheduler.invalidateDirtySet(compileDirtySet([{ type: "cell-value", row: 0, column: 1 }]));
    flush();

    expect(target.lastDirty).toEqual({
      kind: "regions",
      ranges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }]
    });
  });

  it("keeps non-adjacent regions separate below the threshold", () => {
    const { frameScheduler, flush } = createManualFrameScheduler();
    const target = createCapturingTarget();
    const scheduler = createRenderScheduler(target, { frameScheduler, maxDirtyRegions: 3 });

    scheduler.invalidateDirtySet(compileDirtySet([{ type: "cell-value", row: 0, column: 0 }]));
    scheduler.invalidateDirtySet(compileDirtySet([{ type: "cell-value", row: 10, column: 10 }]));
    flush();

    expect(target.lastDirty).toEqual({
      kind: "regions",
      ranges: [
        { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 },
        { startRow: 10, startColumn: 10, endRow: 10, endColumn: 10 }
      ]
    });
  });

  it("degrades to a full repaint once the region count exceeds maxDirtyRegions", () => {
    const { frameScheduler, flush } = createManualFrameScheduler();
    const target = createCapturingTarget();
    const scheduler = createRenderScheduler(target, { frameScheduler, maxDirtyRegions: 3 });

    scheduler.invalidateDirtySet(compileDirtySet([{ type: "cell-value", row: 0, column: 0 }]));
    scheduler.invalidateDirtySet(compileDirtySet([{ type: "cell-value", row: 10, column: 10 }]));
    scheduler.invalidateDirtySet(compileDirtySet([{ type: "cell-value", row: 20, column: 20 }]));
    scheduler.invalidateDirtySet(compileDirtySet([{ type: "cell-value", row: 30, column: 30 }]));
    flush();

    expect(target.lastDirty).toEqual({ kind: "full" });
  });
});

describe("RenderScheduler structural patches", () => {
  it.each([
    ["row-structure", { type: "row-structure", operation: "insert", start: 0, count: 1 }],
    ["column-layout", { type: "column-layout", column: 0 }],
    ["full-reset", { type: "full-reset" }]
  ] as const)("%s always triggers a full repaint", (_label, patch) => {
    const { frameScheduler, flush } = createManualFrameScheduler();
    const target = createCapturingTarget();
    const scheduler = createRenderScheduler(target, { frameScheduler });

    scheduler.invalidateDirtySet(compileDirtySet([patch]));
    flush();

    expect(target.lastDirty).toEqual({ kind: "full" });
  });
});

describe("RenderScheduler selection/content independence", () => {
  it("invalidateSelection repaints the selection layer without touching content", () => {
    const { frameScheduler, flush } = createManualFrameScheduler();
    const target = createCapturingTarget();
    const scheduler = createRenderScheduler(target, { frameScheduler });

    scheduler.invalidateSelection();
    flush();

    expect(target.paintSelectionCalls).toBe(1);
    expect(target.paintContentCalls).toBe(0);
  });

  it("a pending content invalidation and a selection invalidation both flush in one frame", () => {
    const { frameScheduler, requestFrameCallCount, flush } = createManualFrameScheduler();
    const target = createCapturingTarget();
    const scheduler = createRenderScheduler(target, { frameScheduler });

    scheduler.invalidateDirtySet(compileDirtySet([{ type: "cell-value", row: 0, column: 0 }]));
    scheduler.invalidateSelection();

    expect(requestFrameCallCount()).toBe(1);
    flush();

    expect(target.paintContentCalls).toBe(1);
    expect(target.paintSelectionCalls).toBe(1);
  });
});

describe("RenderScheduler presentation layers", () => {
  it("shares one frame identity across presentation, content and selection", () => {
    const { frameScheduler, flush } = createManualFrameScheduler();
    const frames: RenderFrame<SceneContract>[] = [];
    const target: RenderTarget<SceneContract> = {
      paintPresentation: (frame) => frames.push(frame),
      paintContent: (frame) => frames.push(frame),
      paintSelection: (frame) => frames.push(frame)
    };
    const scheduler = createRenderScheduler(target, { frameScheduler });

    scheduler.invalidateAll();
    scheduler.invalidateSelection();
    flush();

    expect(frames).toHaveLength(3);
    expect(new Set(frames).size).toBe(1);
    expect(frames[0]?.dirty.presentation).toEqual(PRESENTATION_DIRTY);
  });

  it("keeps presentation clean for a selection-only frame", () => {
    const { frameScheduler, flush } = createManualFrameScheduler();
    const target = createCapturingTarget();
    const presentation = vi.spyOn(target, "paintPresentation");
    const scheduler = createRenderScheduler(target, { frameScheduler });

    scheduler.invalidateSelection();
    flush();

    expect(presentation).not.toHaveBeenCalled();
  });

  it("uses the same captured Scene when presentation publishes a newer source", () => {
    const { frameScheduler, flush } = createManualFrameScheduler();
    const first = Object.freeze({ modelVersion: 0, id: "first" });
    const second = Object.freeze({ modelVersion: 0, id: "second" });
    let current: SceneContract = first;
    const scenes: SceneContract[] = [];
    const target: RenderTarget<SceneContract> = {
      paintPresentation(frame) {
        scenes.push(frame.scene);
        current = second;
      },
      paintContent: (frame) => scenes.push(frame.scene),
      paintSelection: (frame) => scenes.push(frame.scene)
    };
    const scheduler = createRenderScheduler(target, {
      frameScheduler,
      getScene: () => current
    });

    scheduler.invalidateAll();
    scheduler.invalidateSelection();
    flush();

    expect(scenes).toEqual([first, first, first]);
  });
});

describe("RenderScheduler.flushNow", () => {
  it("cancels the scheduled callback and synchronously paints one complete frame", () => {
    const { frameScheduler, flush } = createManualFrameScheduler();
    const frames: RenderFrame<SceneContract>[] = [];
    const target: RenderTarget<SceneContract> = {
      paintPresentation: (frame) => frames.push(frame),
      paintContent: (frame) => frames.push(frame),
      paintSelection: (frame) => frames.push(frame)
    };
    const scheduler = createRenderScheduler(target, { frameScheduler });

    scheduler.invalidateAll();
    scheduler.invalidateSelection();
    scheduler.flushNow();

    expect(frames).toHaveLength(3);
    expect(new Set(frames).size).toBe(1);
    expect(frames[0]?.dirty).toEqual({
      content: { kind: "full" },
      selection: true,
      presentation: PRESENTATION_DIRTY
    });

    flush();
    expect(frames).toHaveLength(3);
  });

  it("is a no-op when no layer is dirty", () => {
    const { frameScheduler } = createManualFrameScheduler();
    const target: RenderTarget<SceneContract> = {
      paintPresentation: vi.fn(),
      paintContent: vi.fn(),
      paintSelection: vi.fn()
    };
    const scheduler = createRenderScheduler(target, { frameScheduler });

    scheduler.flushNow();

    expect(target.paintPresentation).not.toHaveBeenCalled();
    expect(target.paintContent).not.toHaveBeenCalled();
    expect(target.paintSelection).not.toHaveBeenCalled();
  });

  it("retains every dirty layer when a synchronous frame fails", () => {
    const { frameScheduler } = createManualFrameScheduler();
    const onFrameError = vi.fn();
    const frames: RenderFrame<SceneContract>[] = [];
    let shouldThrow = true;
    const target: RenderTarget<SceneContract> = {
      paintPresentation(frame) {
        frames.push(frame);
        if (shouldThrow) {
          throw new Error("presentation failed");
        }
      },
      paintContent: (frame) => frames.push(frame),
      paintSelection: (frame) => frames.push(frame)
    };
    const scheduler = createRenderScheduler(target, { frameScheduler, onFrameError });

    scheduler.invalidateAll();
    scheduler.invalidateSelection();
    scheduler.flushNow();
    expect(onFrameError).toHaveBeenCalledTimes(1);

    shouldThrow = false;
    scheduler.flushNow();

    expect(frames).toHaveLength(4);
    expect(frames.slice(1)).toHaveLength(3);
    expect(new Set(frames.slice(1)).size).toBe(1);
    expect(frames[1]?.dirty).toEqual({
      content: { kind: "full" },
      selection: true,
      presentation: PRESENTATION_DIRTY
    });
  });
});

describe("RenderScheduler frame snapshot contract", () => {
  it("captures one immutable frame under the lease and shares it across layers", () => {
    const { frameScheduler, flush } = createManualFrameScheduler();
    const revisionGate = new RevisionGate(4);
    const scene = Object.freeze({ modelVersion: 4 });
    const getScene = vi.fn(() => scene);
    const frames: RenderFrame<SceneContract>[] = [];
    const leaseCounts: number[] = [];
    const target: RenderTarget<SceneContract> = {
      paintPresentation(frame) {
        frames.push(frame);
        leaseCounts.push(revisionGate.activeReadLeases);
      },
      paintContent(frame) {
        frames.push(frame);
        leaseCounts.push(revisionGate.activeReadLeases);
      },
      paintSelection(frame) {
        frames.push(frame);
        leaseCounts.push(revisionGate.activeReadLeases);
      }
    };
    const scheduler = createRenderScheduler(target, { frameScheduler, revisionGate, getScene });

    scheduler.invalidateDirtySet(compileDirtySet([{ type: "cell-value", row: 2, column: 3 }]));
    scheduler.invalidateSelection();
    flush();

    expect(getScene).toHaveBeenCalledTimes(1);
    expect(frames).toHaveLength(3);
    expect(frames[0]).toBe(frames[1]);
    expect(frames[1]).toBe(frames[2]);
    expect(frames[0]).toMatchObject({
      scene,
      revision: 4,
      dirty: {
        content: {
          kind: "regions",
          ranges: [{ startRow: 2, startColumn: 3, endRow: 2, endColumn: 3 }]
        },
        selection: true
      }
    });
    expect(Object.isFrozen(frames[0])).toBe(true);
    expect(Object.isFrozen(frames[0]?.dirty)).toBe(true);
    expect(leaseCounts).toEqual([1, 1, 1]);
    expect(revisionGate.activeReadLeases).toBe(0);
  });

  it("does not let a scene source change during content paint tear the frame", () => {
    const { frameScheduler, flush } = createManualFrameScheduler();
    const revisionGate = new RevisionGate(7);
    const firstScene = Object.freeze({ modelVersion: 7, id: "first" });
    const secondScene = Object.freeze({ modelVersion: 7, id: "second" });
    let currentScene: SceneContract = firstScene;
    const getScene = vi.fn(() => currentScene);
    let contentFrame: RenderFrame<SceneContract> | undefined;
    let selectionFrame: RenderFrame<SceneContract> | undefined;
    const target: RenderTarget<SceneContract> = {
      paintPresentation() {},
      paintContent(frame) {
        contentFrame = frame;
        currentScene = secondScene;
      },
      paintSelection(frame) {
        selectionFrame = frame;
      }
    };
    const scheduler = createRenderScheduler(target, { frameScheduler, revisionGate, getScene });

    scheduler.invalidateAll();
    scheduler.invalidateSelection();
    flush();

    expect(getScene).toHaveBeenCalledTimes(1);
    expect(contentFrame).toBe(selectionFrame);
    expect(contentFrame?.scene).toBe(firstScene);
    expect(selectionFrame?.scene).not.toBe(secondScene);
  });

  it("reports a lease and scene revision mismatch without painting either layer", () => {
    const { frameScheduler, flush } = createManualFrameScheduler();
    const revisionGate = new RevisionGate(8);
    const scene = Object.freeze({ modelVersion: 7 });
    const onFrameError = vi.fn();
    const target: RenderTarget<SceneContract> = {
      paintPresentation: vi.fn(),
      paintContent: vi.fn(),
      paintSelection: vi.fn()
    };
    const scheduler = createRenderScheduler(target, {
      frameScheduler,
      revisionGate,
      getScene: () => scene,
      onFrameError
    });

    scheduler.invalidateAll();
    scheduler.invalidateSelection();
    expect(() => flush()).not.toThrow();

    expect(onFrameError).toHaveBeenCalledTimes(1);
    expect(String(onFrameError.mock.calls[0]?.[0])).toMatch(/revision|modelVersion/i);
    expect(target.paintContent).not.toHaveBeenCalled();
    expect(target.paintSelection).not.toHaveBeenCalled();
    expect(revisionGate.activeReadLeases).toBe(0);
  });

  it("keeps pending dirty work after rejecting a stale scene", () => {
    const { frameScheduler, flush } = createManualFrameScheduler();
    const revisionGate = new RevisionGate(3);
    let currentScene: SceneContract = Object.freeze({ modelVersion: 2 });
    const frames: RenderFrame<SceneContract>[] = [];
    const target: RenderTarget<SceneContract> = {
      paintPresentation: vi.fn(),
      paintContent: (frame) => frames.push(frame),
      paintSelection: (frame) => frames.push(frame)
    };
    const scheduler = createRenderScheduler(target, {
      frameScheduler,
      revisionGate,
      getScene: () => currentScene,
      onFrameError: vi.fn()
    });

    scheduler.invalidateDirtySet(compileDirtySet([{ type: "cell-value", row: 4, column: 5 }]));
    scheduler.invalidateSelection();
    flush();
    expect(frames).toEqual([]);

    currentScene = Object.freeze({ modelVersion: 3 });
    scheduler.invalidateSelection();
    flush();

    expect(frames).toHaveLength(2);
    expect(frames[0]).toBe(frames[1]);
    expect(frames[0]?.dirty).toEqual({
      content: {
        kind: "regions",
        ranges: [{ startRow: 4, startColumn: 5, endRow: 4, endColumn: 5 }]
      },
      presentation: { ...PRESENTATION_DIRTY, columnHeader: false, rowHeader: false, domHeaders: false },
      selection: true
    });
  });

  it("releases the lease and retains pending dirty work when a renderer throws", () => {
    const { frameScheduler, flush } = createManualFrameScheduler();
    const revisionGate = new RevisionGate(5);
    const onFrameError = vi.fn();
    const contentFrames: RenderFrame<SceneContract>[] = [];
    const selectionFrames: RenderFrame<SceneContract>[] = [];
    let shouldThrow = true;
    const target: RenderTarget<SceneContract> = {
      paintPresentation: vi.fn(),
      paintContent(frame) {
        contentFrames.push(frame);
        if (shouldThrow) {
          throw new Error("content paint failed");
        }
      },
      paintSelection(frame) {
        selectionFrames.push(frame);
      }
    };
    const scheduler = createRenderScheduler(target, {
      frameScheduler,
      revisionGate,
      getScene: () => Object.freeze({ modelVersion: 5 }),
      onFrameError
    });

    scheduler.invalidateDirtySet(compileDirtySet([{ type: "cell-value", row: 6, column: 7 }]));
    scheduler.invalidateSelection();
    flush();

    expect(onFrameError).toHaveBeenCalledTimes(1);
    expect(selectionFrames).toEqual([]);
    expect(revisionGate.activeReadLeases).toBe(0);

    shouldThrow = false;
    scheduler.invalidateSelection();
    flush();

    expect(contentFrames).toHaveLength(2);
    expect(selectionFrames).toHaveLength(1);
    expect(selectionFrames[0]).toBe(contentFrames[1]);
    expect(contentFrames[1]?.dirty).toEqual({
      content: {
        kind: "regions",
        ranges: [{ startRow: 6, startColumn: 7, endRow: 6, endColumn: 7 }]
      },
      presentation: { ...PRESENTATION_DIRTY, columnHeader: false, rowHeader: false, domHeaders: false },
      selection: true
    });
    expect(revisionGate.activeReadLeases).toBe(0);
  });
});

describe("RenderScheduler.dispose", () => {
  it("cancels an already-scheduled but not-yet-fired frame", () => {
    let cancelledHandle: number | undefined;
    const frameScheduler: FrameScheduler = {
      requestFrame: () => 42,
      cancelFrame: (handle) => {
        cancelledHandle = handle;
      }
    };
    const target = createCapturingTarget();
    const scheduler = createRenderScheduler(target, { frameScheduler });

    scheduler.invalidateAll();
    scheduler.dispose();

    expect(cancelledHandle).toBe(42);
  });

  it("is safe to call when no frame is pending", () => {
    const frameScheduler: FrameScheduler = {
      requestFrame: () => 1,
      cancelFrame: vi.fn()
    };
    const target = createCapturingTarget();
    const scheduler = createRenderScheduler(target, { frameScheduler });

    expect(() => scheduler.dispose()).not.toThrow();
    expect(frameScheduler.cancelFrame).not.toHaveBeenCalled();
  });
});
