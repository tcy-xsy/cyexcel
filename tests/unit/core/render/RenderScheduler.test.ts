import { describe, expect, it, vi } from "vitest";

import { RenderScheduler } from "../../../../src/core";
import type { DirtyState, FrameScheduler, RenderTarget } from "../../../../src/core";

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

function createCapturingTarget(): RenderTarget & {
  lastDirty: DirtyState | undefined;
  paintContentCalls: number;
  paintSelectionCalls: number;
} {
  return {
    lastDirty: undefined,
    paintContentCalls: 0,
    paintSelectionCalls: 0,
    paintContent(dirty) {
      this.lastDirty = dirty;
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
    const scheduler = new RenderScheduler(target, { frameScheduler });

    scheduler.invalidatePatch({ type: "cell-value", row: 0, column: 0 });
    scheduler.invalidatePatch({ type: "cell-value", row: 1, column: 1 });
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
    const scheduler = new RenderScheduler(target, { frameScheduler });

    scheduler.invalidatePatch({
      type: "range-value",
      range: { startRow: 1, startColumn: 1, endRow: 3, endColumn: 3 }
    });
    flush();

    expect(target.lastDirty).toEqual({
      kind: "regions",
      ranges: [{ startRow: 1, startColumn: 1, endRow: 3, endColumn: 3 }]
    });
  });

  it("stays full and keeps scheduling a frame when a region is added after a full invalidation", () => {
    const { frameScheduler, requestFrameCallCount, flush } = createManualFrameScheduler();
    const target = createCapturingTarget();
    const scheduler = new RenderScheduler(target, { frameScheduler });

    scheduler.invalidateAll();
    scheduler.invalidatePatch({ type: "cell-value", row: 0, column: 0 });

    expect(requestFrameCallCount()).toBe(1);
    flush();

    expect(target.lastDirty).toEqual({ kind: "full" });
  });

  it("merges adjacent cell patches into a single region", () => {
    const { frameScheduler, flush } = createManualFrameScheduler();
    const target = createCapturingTarget();
    const scheduler = new RenderScheduler(target, { frameScheduler });

    scheduler.invalidatePatch({ type: "cell-value", row: 0, column: 0 });
    scheduler.invalidatePatch({ type: "cell-value", row: 0, column: 1 });
    flush();

    expect(target.lastDirty).toEqual({
      kind: "regions",
      ranges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }]
    });
  });

  it("keeps non-adjacent regions separate below the threshold", () => {
    const { frameScheduler, flush } = createManualFrameScheduler();
    const target = createCapturingTarget();
    const scheduler = new RenderScheduler(target, { frameScheduler, maxDirtyRegions: 3 });

    scheduler.invalidatePatch({ type: "cell-value", row: 0, column: 0 });
    scheduler.invalidatePatch({ type: "cell-value", row: 10, column: 10 });
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
    const scheduler = new RenderScheduler(target, { frameScheduler, maxDirtyRegions: 3 });

    scheduler.invalidatePatch({ type: "cell-value", row: 0, column: 0 });
    scheduler.invalidatePatch({ type: "cell-value", row: 10, column: 10 });
    scheduler.invalidatePatch({ type: "cell-value", row: 20, column: 20 });
    scheduler.invalidatePatch({ type: "cell-value", row: 30, column: 30 });
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
    const scheduler = new RenderScheduler(target, { frameScheduler });

    scheduler.invalidatePatch(patch);
    flush();

    expect(target.lastDirty).toEqual({ kind: "full" });
  });
});

describe("RenderScheduler selection/content independence", () => {
  it("invalidateSelection repaints the selection layer without touching content", () => {
    const { frameScheduler, flush } = createManualFrameScheduler();
    const target = createCapturingTarget();
    const scheduler = new RenderScheduler(target, { frameScheduler });

    scheduler.invalidateSelection();
    flush();

    expect(target.paintSelectionCalls).toBe(1);
    expect(target.paintContentCalls).toBe(0);
  });

  it("a pending content invalidation and a selection invalidation both flush in one frame", () => {
    const { frameScheduler, requestFrameCallCount, flush } = createManualFrameScheduler();
    const target = createCapturingTarget();
    const scheduler = new RenderScheduler(target, { frameScheduler });

    scheduler.invalidatePatch({ type: "cell-value", row: 0, column: 0 });
    scheduler.invalidateSelection();

    expect(requestFrameCallCount()).toBe(1);
    flush();

    expect(target.paintContentCalls).toBe(1);
    expect(target.paintSelectionCalls).toBe(1);
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
    const scheduler = new RenderScheduler(target, { frameScheduler });

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
    const scheduler = new RenderScheduler(target, { frameScheduler });

    expect(() => scheduler.dispose()).not.toThrow();
    expect(frameScheduler.cancelFrame).not.toHaveBeenCalled();
  });
});
