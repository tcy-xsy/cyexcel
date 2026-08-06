import { describe, expect, it } from "vitest";

import { LayoutEngine, SelectionLayerRenderer } from "../../../../src/core";
import type { SelectionSnapshot } from "../../../../src/core";
import { RecordingCanvasContext } from "../../../helpers/RecordingCanvasContext";

function cellSnapshot(
  startRow: number,
  startColumn: number,
  endRow = startRow,
  endColumn = startColumn
): SelectionSnapshot {
  return {
    kind: "cell",
    active: { row: startRow, column: startColumn },
    anchor: { row: startRow, column: startColumn },
    range: { startRow, startColumn, endRow, endColumn }
  };
}

function createLayout(freeze?: { rows?: number; columns?: number }): LayoutEngine {
  return new LayoutEngine({
    rowCount: 20,
    columnCount: 20,
    rowHeight: 20,
    columnWidths: Array.from({ length: 20 }, () => 50),
    ...(freeze ? { freeze } : {})
  });
}

describe("SelectionLayerRenderer.paintSelection — no freeze", () => {
  it("clears its own canvas area at the last size set via setSize", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new SelectionLayerRenderer({
      ctx,
      layout: createLayout(),
      getSelection: () => cellSnapshot(0, 0)
    });

    renderer.setSize(300, 150);
    renderer.paintSelection();

    expect(ctx.calls[0]).toEqual({ method: "clearRect", args: [0, 0, 300, 150] });
  });

  it("strokes a rect at the selected cell's viewport-relative position", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new SelectionLayerRenderer({
      ctx,
      layout: createLayout(),
      getSelection: () => cellSnapshot(3, 2)
    });
    renderer.setViewport({ scrollTop: 10, scrollLeft: 5, width: 200, height: 200 });

    renderer.paintSelection();

    const strokeRectCalls = ctx.calls.filter((call) => call.method === "strokeRect");
    // Absolute rect for (row 3, col 2): x = 2*50=100, y = 3*20=60. Minus scroll: x=95, y=50.
    expect(strokeRectCalls).toEqual([{ method: "strokeRect", args: [95, 50, 50, 20] }]);
  });

  it("strokes the bounding box of a multi-cell range", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new SelectionLayerRenderer({
      ctx,
      layout: createLayout(),
      getSelection: () => cellSnapshot(1, 1, 2, 3)
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 200, height: 200 });

    renderer.paintSelection();

    const strokeRectCalls = ctx.calls.filter((call) => call.method === "strokeRect");
    // rows 1-2 (y: 20 to 60), columns 1-3 (x: 50 to 200).
    expect(strokeRectCalls).toEqual([{ method: "strokeRect", args: [50, 20, 150, 40] }]);
  });

  it("draws exactly one box for a plain cell selection", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new SelectionLayerRenderer({
      ctx,
      layout: createLayout(),
      getSelection: () => cellSnapshot(0, 0)
    });

    renderer.paintSelection();

    expect(ctx.calls.filter((call) => call.method === "strokeRect")).toHaveLength(1);
  });

  it("sets a thicker lineWidth before stroking the box", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new SelectionLayerRenderer({
      ctx,
      layout: createLayout(),
      getSelection: () => cellSnapshot(0, 0)
    });

    renderer.paintSelection();

    const lineWidthCall = ctx.calls.find((call) => call.method === "set:lineWidth");
    expect(lineWidthCall?.args).toEqual([1.5]);
    const strokeIndex = ctx.calls.findIndex((call) => call.method === "strokeRect");
    const lineWidthIndex = ctx.calls.indexOf(lineWidthCall!);
    expect(lineWidthIndex).toBeLessThan(strokeIndex);
  });

  it("draws a small fill-handle square at the selection's bottom-right corner", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new SelectionLayerRenderer({
      ctx,
      layout: createLayout(),
      getSelection: () => cellSnapshot(3, 2)
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 200, height: 200 });

    renderer.paintSelection();

    // Selected rect is {x:100, y:60, width:50, height:20}; handle is a 6px
    // square centered on the bottom-right corner (103+50-3, 60+20-3).
    const fillRectCalls = ctx.calls.filter((call) => call.method === "fillRect");
    expect(fillRectCalls).toEqual([{ method: "fillRect", args: [147, 77, 6, 6] }]);
  });

  it("draws no fill handle when isVisible() returns false", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new SelectionLayerRenderer({
      ctx,
      layout: createLayout(),
      getSelection: () => cellSnapshot(0, 0),
      isVisible: () => false
    });

    renderer.paintSelection();

    expect(ctx.calls.filter((call) => call.method === "fillRect")).toHaveLength(0);
  });
});

describe("SelectionLayerRenderer.paintSelection — isVisible", () => {
  it("still clears the canvas but draws no box when isVisible() returns false", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new SelectionLayerRenderer({
      ctx,
      layout: createLayout(),
      getSelection: () => cellSnapshot(0, 0),
      isVisible: () => false
    });
    renderer.setSize(300, 150);

    renderer.paintSelection();

    expect(ctx.calls.filter((call) => call.method === "clearRect")).toHaveLength(1);
    expect(ctx.calls.filter((call) => call.method === "strokeRect")).toHaveLength(0);
  });

  it("defaults to always visible when isVisible is not provided", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new SelectionLayerRenderer({
      ctx,
      layout: createLayout(),
      getSelection: () => cellSnapshot(0, 0)
    });

    renderer.paintSelection();

    expect(ctx.calls.filter((call) => call.method === "strokeRect")).toHaveLength(1);
  });
});

describe("SelectionLayerRenderer.paintSelection — fill preview", () => {
  it("draws no extra dashed stroke when getFillPreviewRange returns null", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new SelectionLayerRenderer({
      ctx,
      layout: createLayout(),
      getSelection: () => cellSnapshot(0, 0),
      getFillPreviewRange: () => null
    });

    renderer.paintSelection();

    expect(ctx.calls.filter((call) => call.method === "setLineDash")).toHaveLength(0);
    expect(ctx.calls.filter((call) => call.method === "strokeRect")).toHaveLength(1);
  });

  it("behaves exactly as before when getFillPreviewRange is not provided", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new SelectionLayerRenderer({
      ctx,
      layout: createLayout(),
      getSelection: () => cellSnapshot(0, 0)
    });

    renderer.paintSelection();

    expect(ctx.calls.filter((call) => call.method === "setLineDash")).toHaveLength(0);
  });

  it("strokes a dashed rect for the preview range, resetting the dash afterwards", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new SelectionLayerRenderer({
      ctx,
      layout: createLayout(),
      getSelection: () => cellSnapshot(0, 0),
      getFillPreviewRange: () => ({ startRow: 0, startColumn: 0, endRow: 2, endColumn: 0 })
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 200, height: 200 });

    renderer.paintSelection();

    const strokeRectCalls = ctx.calls.filter((call) => call.method === "strokeRect");
    // Selection box (row 0) + preview box (rows 0-2): x=0, y=0, width=50, height=60.
    expect(strokeRectCalls).toHaveLength(2);
    expect(strokeRectCalls[1]).toEqual({ method: "strokeRect", args: [0, 0, 50, 60] });

    const dashCalls = ctx.calls.filter((call) => call.method === "setLineDash");
    expect(dashCalls).toEqual([
      { method: "setLineDash", args: [[4, 3]] },
      { method: "setLineDash", args: [[]] }
    ]);
    // The dash reset must come after the preview's own strokeRect, and the
    // dash-set must come before it, so only the preview box is dashed.
    const previewStrokeIndex = ctx.calls.indexOf(strokeRectCalls[1]!);
    const dashSetIndex = ctx.calls.indexOf(dashCalls[0]!);
    const dashResetIndex = ctx.calls.indexOf(dashCalls[1]!);
    expect(dashSetIndex).toBeLessThan(previewStrokeIndex);
    expect(dashResetIndex).toBeGreaterThan(previewStrokeIndex);
  });
});

describe("SelectionLayerRenderer.paintSelection — frozen rows/columns", () => {
  it("draws one unshifted box when the whole selection is inside the frozen corner", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new SelectionLayerRenderer({
      ctx,
      layout: createLayout({ rows: 2, columns: 2 }),
      getSelection: () => cellSnapshot(0, 0)
    });
    renderer.setViewport({ scrollTop: 500, scrollLeft: 500, width: 200, height: 200 });

    renderer.paintSelection();

    const strokeRectCalls = ctx.calls.filter((call) => call.method === "strokeRect");
    // Frozen corner cell (0,0): absolute {x:0,y:0}, unaffected by scroll.
    expect(strokeRectCalls).toEqual([{ method: "strokeRect", args: [0, 0, 50, 20] }]);
  });

  it("splits a selection straddling the frozen row boundary into a frozen box and a scrollable box", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new SelectionLayerRenderer({
      ctx,
      layout: createLayout({ rows: 1 }),
      getSelection: () => cellSnapshot(0, 0, 1, 0) // rows 0-1: row 0 frozen, row 1 scrollable
    });
    renderer.setViewport({ scrollTop: 100, scrollLeft: 0, width: 200, height: 200 });

    renderer.paintSelection();

    const strokeRectCalls = ctx.calls.filter((call) => call.method === "strokeRect");
    expect(strokeRectCalls).toHaveLength(2);
    // Frozen row 0: y unshifted.
    expect(strokeRectCalls).toContainEqual({ method: "strokeRect", args: [0, 0, 50, 20] });
    // Scrollable row 1: absolute y=20, minus scrollTop 100 = -80.
    expect(strokeRectCalls).toContainEqual({ method: "strokeRect", args: [0, -80, 50, 20] });
  });

  it("keeps only the actually-frozen axis unshifted across 0/1/multiple frozen rows and columns", () => {
    const cases: Array<{
      readonly freeze: { rows?: number; columns?: number } | undefined;
      readonly expected: readonly [number, number, number, number];
    }> = [
      { freeze: undefined, expected: [-300, -300, 50, 20] }, // neither axis frozen
      { freeze: { rows: 1 }, expected: [-300, 0, 50, 20] }, // row frozen, column isn't
      { freeze: { rows: 2, columns: 2 }, expected: [0, 0, 50, 20] } // both axes frozen
    ];

    for (const { freeze, expected } of cases) {
      const ctx = new RecordingCanvasContext();
      const renderer = new SelectionLayerRenderer({
        ctx,
        layout: createLayout(freeze),
        getSelection: () => cellSnapshot(0, 0)
      });
      renderer.setViewport({ scrollTop: 300, scrollLeft: 300, width: 200, height: 200 });

      renderer.paintSelection();

      const strokeRectCalls = ctx.calls.filter((call) => call.method === "strokeRect");
      expect(strokeRectCalls).toHaveLength(1);
      expect(strokeRectCalls[0]?.args).toEqual(expected);
    }
  });
});
