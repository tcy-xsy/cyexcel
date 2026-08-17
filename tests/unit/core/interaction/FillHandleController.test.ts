import { describe, expect, it, vi } from "vitest";

import {
  FillHandleController,
  isPointOnFillHandle
} from "../../../../src/core/interaction/FillHandleController";
import type { SelectionSnapshot } from "../../../../src/core/interaction/SelectionModel";
import {
  LayoutEngine,
  SkeletonSnapshot,
  ViewportSnapshot,
  type ViewportState
} from "../../../../src/core/layout";
import { SceneSnapshot, surfacePoint } from "../../../../src/core/scene";
import type { CellRange } from "../../../../src/core/types";

function createLayout(freeze?: { readonly rows?: number; readonly columns?: number }) {
  return new LayoutEngine({
    rowCount: 20,
    columnCount: 10,
    rowHeight: 20,
    columnWidths: Array.from({ length: 10 }, () => 50),
    ...(freeze ? { freeze } : {})
  });
}

const VIEWPORT: ViewportState = { scrollTop: 0, scrollLeft: 0, width: 500, height: 400 };

function createScene(layout: LayoutEngine, viewport: ViewportState = VIEWPORT): SceneSnapshot {
  return new SceneSnapshot({
    version: 1,
    skeleton: SkeletonSnapshot.fromLayout(layout, 1),
    viewport: new ViewportSnapshot({ version: 1, ...viewport })
  });
}

const STRUCTURE_VERSION = 7;

function bodyPoint(x: number, y: number) {
  return surfacePoint("body", x, y);
}

function rangeSnapshot(range: CellRange): SelectionSnapshot {
  return {
    kind: "cell",
    active: { row: range.startRow, column: range.startColumn },
    anchor: { row: range.startRow, column: range.startColumn },
    range,
    ranges: [range]
  };
}

describe("isPointOnFillHandle", () => {
  it("hits near the bottom-right corner of a single-cell selection", () => {
    const layout = createLayout();
    const range: CellRange = { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 };

    // Cell (0,0) rect is {x:0,y:0,w:50,h:20} -> bottom-right corner (50,20).
    expect(isPointOnFillHandle(bodyPoint(50, 20), range, createScene(layout))).toBe(true);
  });

  it("misses a point far from the handle", () => {
    const layout = createLayout();
    const range: CellRange = { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 };

    expect(isPointOnFillHandle(bodyPoint(10, 10), range, createScene(layout))).toBe(false);
  });

  it("does not hit a scrollable handle hidden behind the frozen-column pane", () => {
    const layout = createLayout({ columns: 1 });
    const range: CellRange = { startRow: 0, startColumn: 1, endRow: 0, endColumn: 1 };
    const scene = createScene(layout, { ...VIEWPORT, scrollLeft: 54 });

    expect(isPointOnFillHandle(bodyPoint(46, 20), range, scene)).toBe(false);
  });

  it("clips the hit area to the same pane edge as a partially visible handle", () => {
    const layout = createLayout({ columns: 1 });
    const range: CellRange = { startRow: 0, startColumn: 1, endRow: 0, endColumn: 1 };
    const scene = createScene(layout, { ...VIEWPORT, scrollLeft: 50 });

    expect(scene.getSelectionHandleFragment(range, 6)?.rect).toEqual({
      x: 47,
      y: 17,
      width: 6,
      height: 6
    });
    expect(isPointOnFillHandle(bodyPoint(49, 20), range, scene)).toBe(false);
    expect(isPointOnFillHandle(bodyPoint(50, 20), range, scene)).toBe(true);
  });

  it("does not hit a handle outside the viewport", () => {
    const layout = createLayout();
    const range: CellRange = { startRow: 0, startColumn: 5, endRow: 0, endColumn: 5 };
    const scene = createScene(layout, { ...VIEWPORT, width: 100 });

    expect(isPointOnFillHandle(bodyPoint(300, 20), range, scene)).toBe(false);
  });
});

describe("FillHandleController drag", () => {
  it("does not start a drag when pointerdown misses the handle", () => {
    const layout = createLayout();
    const scene = createScene(layout);
    const range: CellRange = { startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 };
    const controller = new FillHandleController({});

    const started = controller.handlePointerDown(
      bodyPoint(10, 10),
      rangeSnapshot(range),
      scene,
      STRUCTURE_VERSION
    );

    expect(started).toBe(false);
    expect(controller.dragging).toBe(false);
  });

  it("starts a drag when pointerdown hits the handle", () => {
    const layout = createLayout();
    const scene = createScene(layout);
    const range: CellRange = { startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 };
    const controller = new FillHandleController({});

    // Cell (2,2) rect: x=100,y=40,w=50,h=20 -> corner (150,60).
    const started = controller.handlePointerDown(
      bodyPoint(150, 60),
      rangeSnapshot(range),
      scene,
      STRUCTURE_VERSION
    );

    expect(started).toBe(true);
    expect(controller.dragging).toBe(true);
  });

  it("extends the row axis when the vertical drag distance is larger", () => {
    const layout = createLayout();
    const scene = createScene(layout);
    const range: CellRange = { startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 };
    const onPreviewChange = vi.fn();
    const controller = new FillHandleController({
      onFillPreviewChange: onPreviewChange
    });
    controller.handlePointerDown(bodyPoint(150, 60), rangeSnapshot(range), scene, STRUCTURE_VERSION);

    // Row 5 center: y = 5*20+10 = 110. Column stays at 2 (x=110).
    controller.handlePointerMove(bodyPoint(110, 110), scene, STRUCTURE_VERSION);

    expect(controller.previewRange).toEqual({ startRow: 2, startColumn: 2, endRow: 5, endColumn: 2 });
    expect(onPreviewChange).toHaveBeenLastCalledWith({
      startRow: 2,
      startColumn: 2,
      endRow: 5,
      endColumn: 2
    });
  });

  it("extends the column axis when the horizontal drag distance is larger", () => {
    const layout = createLayout();
    const scene = createScene(layout);
    const range: CellRange = { startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 };
    const controller = new FillHandleController({});
    controller.handlePointerDown(bodyPoint(150, 60), rangeSnapshot(range), scene, STRUCTURE_VERSION);

    // Column 7 center: x = 7*50+25 = 375. Row stays at 2 (y=50).
    controller.handlePointerMove(bodyPoint(375, 50), scene, STRUCTURE_VERSION);

    expect(controller.previewRange).toEqual({ startRow: 2, startColumn: 2, endRow: 2, endColumn: 7 });
  });

  it("produces no preview when the pointer moves back over the source selection", () => {
    const layout = createLayout();
    const scene = createScene(layout);
    const range: CellRange = { startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 };
    const controller = new FillHandleController({});
    controller.handlePointerDown(bodyPoint(150, 60), rangeSnapshot(range), scene, STRUCTURE_VERSION);

    // Dragging up/left of the source cell must not shrink or reverse it.
    controller.handlePointerMove(bodyPoint(10, 10), scene, STRUCTURE_VERSION);

    expect(controller.previewRange).toBeNull();
  });

  it("returns an immutable commit intent for the final preview range on pointerup", () => {
    const layout = createLayout();
    const scene = createScene(layout);
    const range: CellRange = { startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 };
    const controller = new FillHandleController({});
    controller.handlePointerDown(bodyPoint(150, 60), rangeSnapshot(range), scene, STRUCTURE_VERSION);
    controller.handlePointerMove(bodyPoint(110, 110), scene, STRUCTURE_VERSION);

    const intent = controller.handlePointerUp(STRUCTURE_VERSION);

    expect(intent).toEqual({
      sourceRange: range,
      targetRange: {
        startRow: 2,
        startColumn: 2,
        endRow: 5,
        endColumn: 2
      },
      structureVersion: STRUCTURE_VERSION
    });
    expect(Object.isFrozen(intent)).toBe(true);
    expect(Object.isFrozen(intent?.targetRange)).toBe(true);
    expect(controller.dragging).toBe(false);
    expect(controller.previewRange).toBeNull();
  });

  it("does not commit when pointerup happens without any extension", () => {
    const layout = createLayout();
    const scene = createScene(layout);
    const range: CellRange = { startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 };
    const controller = new FillHandleController({});
    controller.handlePointerDown(bodyPoint(150, 60), rangeSnapshot(range), scene, STRUCTURE_VERSION);

    const intent = controller.handlePointerUp(STRUCTURE_VERSION);

    expect(intent).toBeNull();
  });

  it("clears the preview via onFillPreviewChange(null) on pointerup", () => {
    const layout = createLayout();
    const scene = createScene(layout);
    const range: CellRange = { startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 };
    const onPreviewChange = vi.fn();
    const controller = new FillHandleController({
      onFillPreviewChange: onPreviewChange
    });
    controller.handlePointerDown(bodyPoint(150, 60), rangeSnapshot(range), scene, STRUCTURE_VERSION);
    controller.handlePointerMove(bodyPoint(110, 110), scene, STRUCTURE_VERSION);

    controller.handlePointerUp(STRUCTURE_VERSION);

    expect(onPreviewChange).toHaveBeenLastCalledWith(null);
  });

  it("ignores pointermove/pointerup when no drag is in progress", () => {
    const layout = createLayout();
    const scene = createScene(layout);
    const controller = new FillHandleController({});

    controller.handlePointerMove(bodyPoint(110, 110), scene, STRUCTURE_VERSION);
    const intent = controller.handlePointerUp(STRUCTURE_VERSION);

    expect(intent).toBeNull();
    expect(controller.previewRange).toBeNull();
  });

  it("cancels an active preview when structure changes during pointermove", () => {
    const layout = createLayout();
    const scene = createScene(layout);
    const range: CellRange = { startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 };
    const onPreviewChange = vi.fn();
    const controller = new FillHandleController({
      onFillPreviewChange: onPreviewChange
    });
    controller.handlePointerDown(bodyPoint(150, 60), rangeSnapshot(range), scene, STRUCTURE_VERSION);
    controller.handlePointerMove(bodyPoint(110, 110), scene, STRUCTURE_VERSION);

    controller.handlePointerMove(bodyPoint(110, 130), scene, STRUCTURE_VERSION + 1);

    expect(controller.dragging).toBe(false);
    expect(controller.previewRange).toBeNull();
    expect(onPreviewChange).toHaveBeenLastCalledWith(null);
    expect(controller.handlePointerUp(STRUCTURE_VERSION + 1)).toBeNull();
  });

  it("cancels instead of committing when structure changes before pointerup", () => {
    const layout = createLayout();
    const scene = createScene(layout);
    const range: CellRange = { startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 };
    const onPreviewChange = vi.fn();
    const controller = new FillHandleController({
      onFillPreviewChange: onPreviewChange
    });
    controller.handlePointerDown(bodyPoint(150, 60), rangeSnapshot(range), scene, STRUCTURE_VERSION);
    controller.handlePointerMove(bodyPoint(110, 110), scene, STRUCTURE_VERSION);

    const intent = controller.handlePointerUp(STRUCTURE_VERSION + 1);

    expect(controller.dragging).toBe(false);
    expect(controller.previewRange).toBeNull();
    expect(onPreviewChange).toHaveBeenLastCalledWith(null);
    expect(intent).toBeNull();
  });

  it("cancel clears an active preview without committing", () => {
    const layout = createLayout();
    const scene = createScene(layout);
    const range: CellRange = { startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 };
    const onPreviewChange = vi.fn();
    const controller = new FillHandleController({
      onFillPreviewChange: onPreviewChange
    });
    controller.handlePointerDown(bodyPoint(150, 60), rangeSnapshot(range), scene, STRUCTURE_VERSION);
    controller.handlePointerMove(bodyPoint(110, 110), scene, STRUCTURE_VERSION);

    controller.cancel();

    expect(controller.dragging).toBe(false);
    expect(controller.previewRange).toBeNull();
    expect(onPreviewChange).toHaveBeenLastCalledWith(null);
    expect(controller.handlePointerUp(STRUCTURE_VERSION)).toBeNull();
  });

  it("cancel is silent while idle", () => {
    const onPreviewChange = vi.fn();
    const controller = new FillHandleController({
      onFillPreviewChange: onPreviewChange
    });

    controller.cancel();
    controller.cancel();

    expect(onPreviewChange).not.toHaveBeenCalled();
  });
});
