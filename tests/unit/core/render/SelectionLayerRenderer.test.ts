import { describe, expect, it } from "vitest";

import {
  LayoutEngine,
  SkeletonSnapshot,
  ViewportSnapshot
} from "../../../../src/core/layout";
import type { ViewportState } from "../../../../src/core/layout";
import { SelectionLayerRenderer } from "../../../../src/core/render";
import { SceneSnapshot } from "../../../../src/core/scene";
import type { SelectionSnapshot } from "../../../../src/core/interaction";
import type { CellRange } from "../../../../src/core/types";
import { RecordingCanvasContext } from "../../../helpers/RecordingCanvasContext";

function multiSnapshot(ranges: readonly CellRange[]): SelectionSnapshot {
  const active = ranges[ranges.length - 1]!;
  return {
    kind: "cell",
    active: { row: active.startRow, column: active.startColumn },
    anchor: { row: active.startRow, column: active.startColumn },
    range: active,
    ranges
  };
}

function cellSnapshot(
  startRow: number,
  startColumn: number,
  endRow = startRow,
  endColumn = startColumn
): SelectionSnapshot {
  const range = { startRow, startColumn, endRow, endColumn };
  return {
    kind: "cell",
    active: { row: startRow, column: startColumn },
    anchor: { row: startRow, column: startColumn },
    range,
    ranges: [range]
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

function createScene(layout: LayoutEngine, viewport: ViewportState): SceneSnapshot {
  return new SceneSnapshot({
    version: 1,
    skeleton: SkeletonSnapshot.fromLayout(layout, 1),
    viewport: new ViewportSnapshot({ version: 1, ...viewport })
  });
}

function createSelectionRenderer(
  options: Omit<ConstructorParameters<typeof SelectionLayerRenderer>[0], "getScene"> & {
    readonly layout: LayoutEngine;
    readonly viewport?: ViewportState;
  }
): SelectionLayerRenderer & { readonly setViewport: (viewport: ViewportState) => void } {
  const { layout, viewport = { scrollTop: 0, scrollLeft: 0, width: 0, height: 0 }, ...rendererOptions } = options;
  let scene = createScene(layout, viewport);
  const renderer = new SelectionLayerRenderer({
    ...rendererOptions,
    getScene: () => scene
  });
  return Object.assign(renderer, {
    setViewport(nextViewport: ViewportState) {
      scene = createScene(layout, nextViewport);
    }
  });
}

describe("SelectionLayerRenderer.paintSelection — no freeze", () => {
  it("uses an explicitly captured frame Scene without rereading the live Scene source", () => {
    const ctx = new RecordingCanvasContext();
    const frameScene = createScene(createLayout(), {
      scrollTop: 0,
      scrollLeft: 0,
      width: 200,
      height: 100
    });
    const renderer = new SelectionLayerRenderer({
      ctx,
      getScene: () => {
        throw new Error("live Scene must not be read during a captured render frame");
      },
      getSelection: () => cellSnapshot(0, 0)
    });
    renderer.setSize(200, 100);

    expect(() => renderer.paintSelection(frameScene)).not.toThrow();
    expect(ctx.calls.filter((call) => call.method === "strokeRect")).toHaveLength(1);
  });

  it("clears its own canvas area at the last size set via setSize", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createSelectionRenderer({
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
    const renderer = createSelectionRenderer({
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
    const renderer = createSelectionRenderer({
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
    const renderer = createSelectionRenderer({
      ctx,
      layout: createLayout(),
      getSelection: () => cellSnapshot(0, 0)
    });

    renderer.paintSelection();

    expect(ctx.calls.filter((call) => call.method === "strokeRect")).toHaveLength(1);
  });

  it("sets a thicker lineWidth before stroking the box", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createSelectionRenderer({
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
    const renderer = createSelectionRenderer({
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

  it("clips a partially visible fill handle to its canonical pane", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createSelectionRenderer({
      ctx,
      layout: createLayout({ columns: 1 }),
      getSelection: () => cellSnapshot(0, 1)
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 50, width: 180, height: 90 });

    renderer.paintSelection();

    const fillIndex = ctx.calls.findIndex((call) => call.method === "fillRect");
    expect(ctx.calls.slice(fillIndex - 4, fillIndex + 2)).toEqual([
      { method: "save", args: [] },
      { method: "beginPath", args: [] },
      { method: "rect", args: [50, 0, 130, 90] },
      { method: "clip", args: [] },
      { method: "fillRect", args: [47, 17, 6, 6] },
      { method: "restore", args: [] }
    ]);
  });

  it("does not paint a handle hidden behind a frozen pane or outside the viewport", () => {
    const cases = [
      {
        layout: createLayout({ columns: 1 }),
        selection: cellSnapshot(0, 1),
        viewport: { scrollTop: 0, scrollLeft: 54, width: 180, height: 90 }
      },
      {
        layout: createLayout(),
        selection: cellSnapshot(0, 5),
        viewport: { scrollTop: 0, scrollLeft: 0, width: 100, height: 90 }
      }
    ];

    for (const testCase of cases) {
      const ctx = new RecordingCanvasContext();
      const renderer = createSelectionRenderer({
        ctx,
        layout: testCase.layout,
        getSelection: () => testCase.selection
      });
      renderer.setViewport(testCase.viewport);

      renderer.paintSelection();

      expect(ctx.calls.filter((call) => call.method === "fillRect")).toHaveLength(0);
    }
  });

  it("draws no fill handle when isVisible() returns false", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createSelectionRenderer({
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
    const renderer = createSelectionRenderer({
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
    const renderer = createSelectionRenderer({
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
    const renderer = createSelectionRenderer({
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
    const renderer = createSelectionRenderer({
      ctx,
      layout: createLayout(),
      getSelection: () => cellSnapshot(0, 0)
    });

    renderer.paintSelection();

    expect(ctx.calls.filter((call) => call.method === "setLineDash")).toHaveLength(0);
  });

  it("strokes a dashed rect for the preview range, resetting the dash afterwards", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createSelectionRenderer({
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

  it("splits a preview crossing frozen rows and columns into canonical pane fragments", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createSelectionRenderer({
      ctx,
      layout: createLayout({ rows: 1, columns: 1 }),
      getSelection: () => cellSnapshot(0, 0),
      getFillPreviewRange: () => ({ startRow: 0, startColumn: 0, endRow: 2, endColumn: 2 })
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 180, height: 90 });

    renderer.paintSelection();

    const dashStart = ctx.calls.findIndex(
      (call) => call.method === "setLineDash" && call.args[0]?.toString() === "4,3"
    );
    const dashEnd = ctx.calls.findIndex(
      (call, index) => index > dashStart && call.method === "setLineDash"
    );
    const previewStrokes = ctx.calls
      .slice(dashStart + 1, dashEnd)
      .filter((call) => call.method === "strokeRect");

    expect(previewStrokes).toEqual([
      { method: "strokeRect", args: [0, 0, 50, 20] },
      { method: "strokeRect", args: [50, 0, 100, 20] },
      { method: "strokeRect", args: [0, 20, 50, 40] },
      { method: "strokeRect", args: [50, 20, 100, 40] }
    ]);
  });

  it("clips a partially scrolled preview so scrollable fragments cannot paint into frozen panes", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createSelectionRenderer({
      ctx,
      layout: createLayout({ rows: 1, columns: 1 }),
      getSelection: () => cellSnapshot(0, 0),
      getFillPreviewRange: () => ({ startRow: 0, startColumn: 0, endRow: 2, endColumn: 2 })
    });
    renderer.setViewport({ scrollTop: 7, scrollLeft: 13, width: 180, height: 90 });

    renderer.paintSelection();

    const dashStart = ctx.calls.findIndex(
      (call) => call.method === "setLineDash" && call.args[0]?.toString() === "4,3"
    );
    const dashEnd = ctx.calls.findIndex(
      (call, index) => index > dashStart && call.method === "setLineDash"
    );
    const previewFragmentCalls = ctx.calls
      .slice(dashStart + 1, dashEnd)
      .filter((call) => ["save", "beginPath", "rect", "clip", "strokeRect", "restore"].includes(call.method));

    expect(previewFragmentCalls).toEqual([
      { method: "save", args: [] },
      { method: "beginPath", args: [] },
      { method: "rect", args: [0, 0, 50, 20] },
      { method: "clip", args: [] },
      { method: "strokeRect", args: [0, 0, 50, 20] },
      { method: "restore", args: [] },
      { method: "save", args: [] },
      { method: "beginPath", args: [] },
      { method: "rect", args: [50, 0, 130, 20] },
      { method: "clip", args: [] },
      { method: "strokeRect", args: [37, 0, 100, 20] },
      { method: "restore", args: [] },
      { method: "save", args: [] },
      { method: "beginPath", args: [] },
      { method: "rect", args: [0, 20, 50, 70] },
      { method: "clip", args: [] },
      { method: "strokeRect", args: [0, 13, 50, 40] },
      { method: "restore", args: [] },
      { method: "save", args: [] },
      { method: "beginPath", args: [] },
      { method: "rect", args: [50, 20, 130, 70] },
      { method: "clip", args: [] },
      { method: "strokeRect", args: [37, 13, 100, 40] },
      { method: "restore", args: [] }
    ]);
  });
});

describe("SelectionLayerRenderer.paintSelection — frozen rows/columns", () => {
  it("draws one unshifted box when the whole selection is inside the frozen corner", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createSelectionRenderer({
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
    const renderer = createSelectionRenderer({
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
      const renderer = createSelectionRenderer({
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

  it("clips every selection fragment to its fixed pane during partial-cell scroll", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createSelectionRenderer({
      ctx,
      layout: createLayout({ rows: 1, columns: 1 }),
      getSelection: () => cellSnapshot(0, 0, 2, 2)
    });
    renderer.setViewport({ scrollTop: 7, scrollLeft: 13, width: 180, height: 90 });

    renderer.paintSelection();

    const handleFillIndex = ctx.calls.findIndex((call) => call.method === "fillRect");
    const fragmentCalls = ctx.calls.slice(0, handleFillIndex - 4).filter((call) =>
      ["save", "beginPath", "rect", "clip", "strokeRect", "restore"].includes(call.method)
    );
    expect(fragmentCalls).toEqual([
      { method: "save", args: [] },
      { method: "beginPath", args: [] },
      { method: "rect", args: [0, 0, 50, 20] },
      { method: "clip", args: [] },
      { method: "strokeRect", args: [0, 0, 50, 20] },
      { method: "restore", args: [] },
      { method: "save", args: [] },
      { method: "beginPath", args: [] },
      { method: "rect", args: [50, 0, 130, 20] },
      { method: "clip", args: [] },
      { method: "strokeRect", args: [37, 0, 100, 20] },
      { method: "restore", args: [] },
      { method: "save", args: [] },
      { method: "beginPath", args: [] },
      { method: "rect", args: [0, 20, 50, 70] },
      { method: "clip", args: [] },
      { method: "strokeRect", args: [0, 13, 50, 40] },
      { method: "restore", args: [] },
      { method: "save", args: [] },
      { method: "beginPath", args: [] },
      { method: "rect", args: [50, 20, 130, 70] },
      { method: "clip", args: [] },
      { method: "strokeRect", args: [37, 13, 100, 40] },
      { method: "restore", args: [] }
    ]);

    const clipRects = fragmentCalls.filter((call) => call.method === "rect");
    expect(clipRects.map((call) => call.args)).toEqual([
      [0, 0, 50, 20],
      [50, 0, 130, 20],
      [0, 20, 50, 70],
      [50, 20, 130, 70]
    ]);
  });
});

describe("SelectionLayerRenderer.paintSelection — multiple regions", () => {
  it("draws a stroked box for every selected region, not just the active one", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createSelectionRenderer({
      ctx,
      layout: createLayout(),
      getSelection: () =>
        multiSnapshot([
          { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 },
          { startRow: 3, startColumn: 2, endRow: 3, endColumn: 2 }
        ])
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 200, height: 200 });

    renderer.paintSelection();

    const strokeRectCalls = ctx.calls.filter((call) => call.method === "strokeRect");
    expect(strokeRectCalls).toEqual([
      { method: "strokeRect", args: [0, 0, 50, 20] },
      { method: "strokeRect", args: [100, 60, 50, 20] }
    ]);
  });

  it("only paints the fill handle for the active (last) region", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createSelectionRenderer({
      ctx,
      layout: createLayout(),
      getSelection: () =>
        multiSnapshot([
          { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 },
          { startRow: 3, startColumn: 2, endRow: 3, endColumn: 2 }
        ])
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 200, height: 200 });

    renderer.paintSelection();

    expect(ctx.calls.filter((call) => call.method === "fillRect")).toHaveLength(1);
  });

  it("paints non-active regions at reduced opacity, bracketed back to full opacity around each one", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createSelectionRenderer({
      ctx,
      layout: createLayout(),
      getSelection: () =>
        multiSnapshot([
          { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 },
          { startRow: 3, startColumn: 2, endRow: 3, endColumn: 2 }
        ])
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 200, height: 200 });

    renderer.paintSelection();

    const alphaCalls = ctx.calls.filter((call) => call.method === "set:globalAlpha");
    // One non-active region -> lowered then restored to 1 around its own box;
    // the active region never touches globalAlpha at all (see the
    // single-region test below — it must stay byte-for-byte unchanged).
    expect(alphaCalls).toHaveLength(2);
    expect(alphaCalls[0]?.args[0]).toBeLessThan(1);
    expect(alphaCalls[1]?.args).toEqual([1]);

    const nonActiveStroke = ctx.calls.findIndex(
      (call) => call.method === "strokeRect" && call.args[0] === 0 && call.args[1] === 0
    );
    expect(ctx.calls.indexOf(alphaCalls[0]!)).toBeLessThan(nonActiveStroke);
    expect(ctx.calls.indexOf(alphaCalls[1]!)).toBeGreaterThan(nonActiveStroke);
  });

  it("a single-region selection never touches globalAlpha (unchanged from before multi-region support)", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createSelectionRenderer({
      ctx,
      layout: createLayout(),
      getSelection: () => cellSnapshot(0, 0)
    });

    renderer.paintSelection();

    expect(ctx.calls.filter((call) => call.method === "set:globalAlpha")).toHaveLength(0);
    expect(ctx.calls.filter((call) => call.method === "strokeRect")).toHaveLength(1);
  });
});
