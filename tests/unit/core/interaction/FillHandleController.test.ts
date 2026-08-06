import { describe, expect, it, vi } from "vitest";

import { FillHandleController, LayoutEngine, isPointOnFillHandle } from "../../../../src/core";
import type { CellRange, SelectionSnapshot, ViewportState } from "../../../../src/core";

function createLayout() {
  return new LayoutEngine({
    rowCount: 20,
    columnCount: 10,
    rowHeight: 20,
    columnWidths: Array.from({ length: 10 }, () => 50)
  });
}

const VIEWPORT: ViewportState = { scrollTop: 0, scrollLeft: 0, width: 500, height: 400 };

function rangeSnapshot(range: CellRange): SelectionSnapshot {
  return {
    kind: "cell",
    active: { row: range.startRow, column: range.startColumn },
    anchor: { row: range.startRow, column: range.startColumn },
    range
  };
}

describe("isPointOnFillHandle", () => {
  it("hits near the bottom-right corner of a single-cell selection", () => {
    const layout = createLayout();
    const range: CellRange = { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 };

    // Cell (0,0) rect is {x:0,y:0,w:50,h:20} -> bottom-right corner (50,20).
    expect(isPointOnFillHandle({ x: 50, y: 20 }, range, layout, VIEWPORT)).toBe(true);
  });

  it("misses a point far from the handle", () => {
    const layout = createLayout();
    const range: CellRange = { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 };

    expect(isPointOnFillHandle({ x: 10, y: 10 }, range, layout, VIEWPORT)).toBe(false);
  });
});

describe("FillHandleController drag", () => {
  it("does not start a drag when pointerdown misses the handle", () => {
    const layout = createLayout();
    const range: CellRange = { startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 };
    const controller = new FillHandleController({
      layout,
      getSelection: () => rangeSnapshot(range),
      onFillCommit: vi.fn()
    });
    controller.setViewport(VIEWPORT);

    const started = controller.handlePointerDown({ x: 10, y: 10 });

    expect(started).toBe(false);
    expect(controller.dragging).toBe(false);
  });

  it("starts a drag when pointerdown hits the handle", () => {
    const layout = createLayout();
    const range: CellRange = { startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 };
    const controller = new FillHandleController({
      layout,
      getSelection: () => rangeSnapshot(range),
      onFillCommit: vi.fn()
    });
    controller.setViewport(VIEWPORT);

    // Cell (2,2) rect: x=100,y=40,w=50,h=20 -> corner (150,60).
    const started = controller.handlePointerDown({ x: 150, y: 60 });

    expect(started).toBe(true);
    expect(controller.dragging).toBe(true);
  });

  it("extends the row axis when the vertical drag distance is larger", () => {
    const layout = createLayout();
    const range: CellRange = { startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 };
    const onPreviewChange = vi.fn();
    const controller = new FillHandleController({
      layout,
      getSelection: () => rangeSnapshot(range),
      onFillPreviewChange: onPreviewChange,
      onFillCommit: vi.fn()
    });
    controller.setViewport(VIEWPORT);
    controller.handlePointerDown({ x: 150, y: 60 });

    // Row 5 center: y = 5*20+10 = 110. Column stays at 2 (x=110).
    controller.handlePointerMove({ x: 110, y: 110 });

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
    const range: CellRange = { startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 };
    const controller = new FillHandleController({
      layout,
      getSelection: () => rangeSnapshot(range),
      onFillCommit: vi.fn()
    });
    controller.setViewport(VIEWPORT);
    controller.handlePointerDown({ x: 150, y: 60 });

    // Column 7 center: x = 7*50+25 = 375. Row stays at 2 (y=50).
    controller.handlePointerMove({ x: 375, y: 50 });

    expect(controller.previewRange).toEqual({ startRow: 2, startColumn: 2, endRow: 2, endColumn: 7 });
  });

  it("produces no preview when the pointer moves back over the source selection", () => {
    const layout = createLayout();
    const range: CellRange = { startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 };
    const controller = new FillHandleController({
      layout,
      getSelection: () => rangeSnapshot(range),
      onFillCommit: vi.fn()
    });
    controller.setViewport(VIEWPORT);
    controller.handlePointerDown({ x: 150, y: 60 });

    // Dragging up/left of the source cell must not shrink or reverse it.
    controller.handlePointerMove({ x: 10, y: 10 });

    expect(controller.previewRange).toBeNull();
  });

  it("commits the final preview range on pointerup", () => {
    const layout = createLayout();
    const range: CellRange = { startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 };
    const onFillCommit = vi.fn();
    const controller = new FillHandleController({
      layout,
      getSelection: () => rangeSnapshot(range),
      onFillCommit
    });
    controller.setViewport(VIEWPORT);
    controller.handlePointerDown({ x: 150, y: 60 });
    controller.handlePointerMove({ x: 110, y: 110 });

    controller.handlePointerUp();

    expect(onFillCommit).toHaveBeenCalledWith(range, {
      startRow: 2,
      startColumn: 2,
      endRow: 5,
      endColumn: 2
    });
    expect(controller.dragging).toBe(false);
    expect(controller.previewRange).toBeNull();
  });

  it("does not commit when pointerup happens without any extension", () => {
    const layout = createLayout();
    const range: CellRange = { startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 };
    const onFillCommit = vi.fn();
    const controller = new FillHandleController({
      layout,
      getSelection: () => rangeSnapshot(range),
      onFillCommit
    });
    controller.setViewport(VIEWPORT);
    controller.handlePointerDown({ x: 150, y: 60 });

    controller.handlePointerUp();

    expect(onFillCommit).not.toHaveBeenCalled();
  });

  it("clears the preview via onFillPreviewChange(null) on pointerup", () => {
    const layout = createLayout();
    const range: CellRange = { startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 };
    const onPreviewChange = vi.fn();
    const controller = new FillHandleController({
      layout,
      getSelection: () => rangeSnapshot(range),
      onFillPreviewChange: onPreviewChange,
      onFillCommit: vi.fn()
    });
    controller.setViewport(VIEWPORT);
    controller.handlePointerDown({ x: 150, y: 60 });
    controller.handlePointerMove({ x: 110, y: 110 });

    controller.handlePointerUp();

    expect(onPreviewChange).toHaveBeenLastCalledWith(null);
  });

  it("ignores pointermove/pointerup when no drag is in progress", () => {
    const layout = createLayout();
    const range: CellRange = { startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 };
    const onFillCommit = vi.fn();
    const controller = new FillHandleController({
      layout,
      getSelection: () => rangeSnapshot(range),
      onFillCommit
    });
    controller.setViewport(VIEWPORT);

    controller.handlePointerMove({ x: 110, y: 110 });
    controller.handlePointerUp();

    expect(onFillCommit).not.toHaveBeenCalled();
    expect(controller.previewRange).toBeNull();
  });
});
