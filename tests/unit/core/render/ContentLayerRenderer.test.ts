import { describe, expect, it } from "vitest";

import {
  LayoutEngine,
  SkeletonSnapshot,
  ViewportSnapshot
} from "../../../../src/core/layout";
import type { ViewportState } from "../../../../src/core/layout";
import { createColumnAwarePainter } from "../../../../src/core/cells";
import { ContentLayerRenderer } from "../../../../src/core/render";
import { SceneSnapshot } from "../../../../src/core/scene";
import { CyExcelError, CyExcelErrorCode } from "../../../../src/core/errors";
import { mergesIntersectingRange } from "../../../../src/core/sheet/merges";
import { RecordingCanvasContext } from "../../../helpers/RecordingCanvasContext";
import type { CellValue } from "../../../../src/core/sheet/types";

function createLayout(rowCount: number, columnCount: number): LayoutEngine {
  return new LayoutEngine({
    rowCount,
    columnCount,
    rowHeight: 20,
    columnWidths: Array.from({ length: columnCount }, () => 50)
  });
}

function createScene(layout: LayoutEngine, viewport: ViewportState): SceneSnapshot {
  return new SceneSnapshot({
    version: 1,
    skeleton: SkeletonSnapshot.fromLayout(layout, 1),
    viewport: new ViewportSnapshot({ version: 1, ...viewport })
  });
}

function createContentRenderer(
  options: Omit<ConstructorParameters<typeof ContentLayerRenderer>[0], "getScene"> & {
    readonly layout: LayoutEngine;
    readonly viewport?: ViewportState;
  }
): ContentLayerRenderer & { readonly setViewport: (viewport: ViewportState) => void } {
  const { layout, viewport = { scrollTop: 0, scrollLeft: 0, width: 0, height: 0 }, ...rendererOptions } = options;
  let scene = createScene(layout, viewport);
  const renderer = new ContentLayerRenderer({
    ...rendererOptions,
    getScene: () => scene
  });
  return Object.assign(renderer, {
    setViewport(nextViewport: ViewportState) {
      scene = createScene(layout, nextViewport);
    }
  });
}

describe("ContentLayerRenderer.paintContent — dirty regions", () => {
  it("uses an explicitly captured frame Scene without rereading the live Scene source", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(1, 1);
    const frameScene = createScene(layout, {
      scrollTop: 0,
      scrollLeft: 0,
      width: 50,
      height: 20
    });
    const renderer = new ContentLayerRenderer({
      ctx,
      getScene: () => {
        throw new Error("live Scene must not be read during a captured render frame");
      },
      getCellValue: () => "frame"
    });

    expect(() => renderer.paintContent({ kind: "full" }, frameScene)).not.toThrow();
    expect(ctx.calls.filter((call) => call.method === "fillText")[0]?.args[0]).toBe("frame");
  });

  it("only paints the cell in a single-cell dirty region", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(10, 10);
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: (row, column) => `${row}:${column}`
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 500, height: 200 });

    renderer.paintContent({
      kind: "regions",
      ranges: [{ startRow: 2, startColumn: 3, endRow: 2, endColumn: 3 }]
    });

    const fillTextCalls = ctx.calls.filter((call) => call.method === "fillText");
    expect(fillTextCalls).toHaveLength(1);
    expect(fillTextCalls[0]?.args[0]).toBe("2:3");
  });

  it("draws an empty string for the default painter when the cell is empty", () => {
    // 空格子只有 `undefined` 一种形态：`CellValue` 里没有 `null`，未占用的格子
    // 在 `cellData` 里干脆不存在（sheet/types.ts 不变量 1）。
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(1, 2);
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: () => undefined
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 20 });

    renderer.paintContent({ kind: "full" });

    const fillTextCalls = ctx.calls.filter((call) => call.method === "fillText");
    expect(fillTextCalls).toHaveLength(2);
    expect(fillTextCalls[0]?.args[0]).toBe("");
    expect(fillTextCalls[1]?.args[0]).toBe("");
  });

  it("skips dirty regions that fall entirely outside the visible viewport", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(100, 100);
    const renderer = createContentRenderer({ ctx, layout, getCellValue: () => "x" });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 40 });

    renderer.paintContent({
      kind: "regions",
      ranges: [{ startRow: 50, startColumn: 50, endRow: 50, endColumn: 50 }]
    });

    expect(ctx.calls.filter((call) => call.method === "fillText")).toHaveLength(0);
  });
});

describe("ContentLayerRenderer.paintContent — full repaint stays clipped to the visible range", () => {
  it("paints only the visible cells of a 1,000,000-row table, not the whole table", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(1_000_000, 50);
    const renderer = createContentRenderer({ ctx, layout, getCellValue: () => "x" });
    const viewport = { scrollTop: 500_000, scrollLeft: 0, width: 400, height: 200 };
    renderer.setViewport(viewport);

    renderer.paintContent({ kind: "full" });

    const visibleRows = layout.getVisibleRowRange(viewport);
    const visibleColumns = layout.getVisibleColumnRange(viewport);
    const expectedCells = visibleRows.count * visibleColumns.count;

    const fillTextCalls = ctx.calls.filter((call) => call.method === "fillText");
    expect(fillTextCalls).toHaveLength(expectedCells);
    expect(expectedCells).toBeLessThan(1000);
  });
});

describe("ContentLayerRenderer.paintContent — scroll offset translation", () => {
  it("draws visible cells at their viewport-relative position, not their absolute content position", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(100, 10); // rowHeight 20, columnWidths 50
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: (row, column) => `${row}:${column}`
    });
    renderer.setViewport({ scrollTop: 45, scrollLeft: 30, width: 200, height: 100 });

    renderer.paintContent({
      kind: "regions",
      ranges: [{ startRow: 3, startColumn: 1, endRow: 3, endColumn: 1 }]
    });

    // Absolute rect for (row 3, column 1): x = 1 * 50 = 50, y = 3 * 20 = 60.
    // Viewport-relative (after subtracting scrollLeft/scrollTop): x = 20, y = 15.
    const clearRectCall = ctx.calls.find((call) => call.method === "clearRect");
    expect(clearRectCall?.args.slice(0, 2)).toEqual([20, 15]);

    const fillTextCall = ctx.calls.find((call) => call.method === "fillText");
    expect(fillTextCall?.args[1]).toBe(20 + 4); // rect.x + DEFAULT_PADDING
    expect(fillTextCall?.args[2]).toBe(15 + 10); // rect.y + rowHeight / 2
  });

  it("passes the translated rect to custom painters too, not the absolute one", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(100, 10);
    const seenRects: Array<{ x: number; y: number }> = [];
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: () => "x",
      painter: (context) => {
        seenRects.push({ x: context.rect.x, y: context.rect.y });
      }
    });
    renderer.setViewport({ scrollTop: 45, scrollLeft: 30, width: 200, height: 100 });

    renderer.paintContent({
      kind: "regions",
      ranges: [{ startRow: 3, startColumn: 1, endRow: 3, endColumn: 1 }]
    });

    expect(seenRects).toEqual([{ x: 20, y: 15 }]);
  });
});

describe("ContentLayerRenderer — frozen rows/columns (Step 9)", () => {
  function createFrozenLayout(): LayoutEngine {
    return new LayoutEngine({
      rowCount: 20,
      columnCount: 20,
      rowHeight: 20,
      columnWidths: Array.from({ length: 20 }, () => 50),
      freeze: { rows: 1, columns: 1 }
    });
  }

  it("paints all four quadrants on a full repaint: frozen corner, frozen row band, frozen column band, scrollable body", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createFrozenLayout();
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: (row, column) => `${row}:${column}`
    });
    // Scrolled deep enough that row 0/column 0 (the frozen band) are nowhere
    // near the scrollable body's current position.
    renderer.setViewport({ scrollTop: 200, scrollLeft: 100, width: 150, height: 100 });

    renderer.paintContent({ kind: "full" });

    const fillTextCalls = ctx.calls.filter((call) => call.method === "fillText");
    const painted = new Set(fillTextCalls.map((call) => call.args[0]));

    // Visible columns for this viewport are [3, 5] (see getVisibleColumnRange
    // math: frozenWidth 50 + scrollLeft 100 = contentLeft 150 -> column 3).
    expect(painted.has("0:0")).toBe(true); // frozen corner
    expect(painted.has("0:4")).toBe(true); // frozen row band (row 0, a scrollable column)
    expect(painted.has("11:0")).toBe(true); // frozen column band (column 0, a scrollable row)
    expect(painted.has("11:4")).toBe(true); // scrollable body
  });

  it("keeps a frozen row's y fixed and a frozen column's x fixed regardless of scroll", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createFrozenLayout();
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: () => "x"
    });
    renderer.setViewport({ scrollTop: 200, scrollLeft: 100, width: 150, height: 100 });

    renderer.paintContent({ kind: "full" });

    const clearRectCalls = ctx.calls.filter((call) => call.method === "clearRect");
    // Frozen corner cell (0, 0): absolute rect is {x:0, y:0}; neither axis
    // should be shifted by scroll since both row 0 and column 0 are frozen.
    expect(clearRectCalls.some((call) => call.args[0] === 0 && call.args[1] === 0)).toBe(true);
  });

  it("region dirty invalidation repaints a frozen cell (previously unreachable)", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createFrozenLayout();
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: (row, column) => `${row}:${column}`
    });
    renderer.setViewport({ scrollTop: 200, scrollLeft: 100, width: 150, height: 100 });

    renderer.paintContent({
      kind: "regions",
      ranges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }]
    });

    const fillTextCalls = ctx.calls.filter((call) => call.method === "fillText");
    expect(fillTextCalls.map((call) => call.args[0])).toEqual(["0:0"]);
  });

  it("clips each quadrant's drawing with save/clip/restore", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createFrozenLayout();
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: () => "x"
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 150, height: 100 });

    renderer.paintContent({ kind: "full" });

    const methodCounts = (method: string) => ctx.calls.filter((call) => call.method === method).length;
    const cellCount = ctx.calls.filter((call) => call.method === "fillText").length;
    expect(methodCounts("save")).toBeGreaterThan(0);
    expect(methodCounts("save")).toBe(methodCounts("restore"));
    // One save/clip/restore per quadrant (region clipping) plus one
    // additional save/restore per cell (state isolation between cells —
    // see the "canvas state isolation" describe block below) — clip only
    // happens at the quadrant level, so it's a strict subset of save.
    expect(methodCounts("clip")).toBeGreaterThan(0);
    expect(methodCounts("clip")).toBeLessThan(methodCounts("save"));
    expect(methodCounts("save")).toBe(methodCounts("clip") + cellCount);
  });

  it("degenerates to a single unclipped-by-freeze region when nothing is frozen (no regression)", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(10, 10); // no freeze option — matches every pre-Step-9 test's setup
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: (row, column) => `${row}:${column}`
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 40 });

    renderer.paintContent({ kind: "full" });

    // Exactly one quadrant is non-empty when frozenRowCount/frozenColumnCount
    // are both 0, so exactly one clip region is entered — "save" itself is
    // no longer 1-to-1 with quadrants once per-cell isolation (below) is
    // counted too, so this asserts on "clip" instead.
    expect(ctx.calls.filter((call) => call.method === "clip")).toHaveLength(1);
  });
});

describe("ContentLayerRenderer — canvas state isolation between cells", () => {
  it("wraps each cell's painter call in its own save/restore, independent of the quadrant-level clip", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(1, 3);
    const renderer = createContentRenderer({ ctx, layout, getCellValue: () => "x" });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 150, height: 20 });

    renderer.paintContent({ kind: "full" });

    // 3 cells, 1 quadrant: 1 save+clip+restore for the quadrant, plus one
    // save+restore per cell.
    expect(ctx.calls.filter((call) => call.method === "fillText")).toHaveLength(3);
    expect(ctx.calls.filter((call) => call.method === "save")).toHaveLength(4);
    expect(ctx.calls.filter((call) => call.method === "restore")).toHaveLength(4);
    expect(ctx.calls.filter((call) => call.method === "clip")).toHaveLength(1);
  });

  it("does not let one cell's painter-modified ctx state leak into the next cell's default painting", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(1, 2);
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: (_row, column) => (column === 0 ? "custom" : "plain"),
      painter: (context) => {
        if (context.column === 0) {
          // Simulates a custom CanvasCellRenderer that changes ctx state and
          // never restores it itself — ContentLayerRenderer must isolate
          // this from column 1, which relies on the framework leaving it a
          // clean, known state (exactly what happened with the real
          // playground bug this test is guarding against).
          context.ctx.font = "16px serif";
          context.ctx.fillText(String(context.value), context.rect.x, context.rect.y);
          return;
        }
        // Column 1 deliberately never sets font itself — same shape as
        // createColumnAwarePainter's default text path before its own fix.
        context.ctx.fillText(String(context.value), context.rect.x, context.rect.y);
      }
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 20 });

    renderer.paintContent({ kind: "full" });

    const fontCalls = ctx.calls.filter((call) => call.method === "set:font");
    // Column 0 sets "16px serif"; by the time column 1 is painted, a
    // restore() must have already reset it back to whatever it was before
    // column 0 ran (RecordingCanvasContext doesn't simulate save/restore
    // state snapshotting itself, so this asserts on the call shape instead:
    // a restore() must appear between the two cells' calls).
    const restoreIndex = ctx.calls.findIndex(
      (call, index) => call.method === "restore" && index > ctx.calls.indexOf(fontCalls[0]!)
    );
    const column1FillTextIndex = ctx.calls.findIndex(
      (call, index) => call.method === "fillText" && call.args[0] === "plain" && index > restoreIndex
    );
    expect(restoreIndex).toBeGreaterThan(-1);
    expect(column1FillTextIndex).toBeGreaterThan(restoreIndex);
  });
});

describe("ContentLayerRenderer error isolation", () => {
  it("falls back to the default painter and reports onError when a custom painter throws", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(3, 1);
    const errors: Array<{ error: unknown; cell: { row: number; column: number } }> = [];
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: (row) => `row-${row}`,
      painter: (context) => {
        if (context.row === 1) {
          throw new Error("boom");
        }
      },
      onError: (error, cell) => errors.push({ error, cell })
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 50, height: 60 });

    renderer.paintContent({ kind: "full" });

    expect(errors).toHaveLength(1);
    expect(errors[0]?.cell).toEqual({ row: 1, column: 0 });

    // Rows 0 and 2 are painted by the custom painter, which never calls
    // fillText; only row 1's fallback to the default painter does.
    const fillTextCalls = ctx.calls.filter((call) => call.method === "fillText");
    expect(fillTextCalls).toHaveLength(1);
    expect(fillTextCalls[0]?.args[0]).toBe("row-1");
  });
});

describe("ContentLayerRenderer synchronous renderer contract", () => {
  it("throws immediately when a painter returns a thenable and checkSyncRenderers is true", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(1, 1);
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: () => "x",
      painter: () => Promise.resolve()
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 50, height: 20 });

    try {
      renderer.paintContent({ kind: "full" });
      expect.unreachable("expected paintContent to throw");
    } catch (error) {
      expect(error).toBeInstanceOf(CyExcelError);
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.AsyncRendererNotSupported);
    }
  });

  it("does not check for thenables when checkSyncRenderers is false", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(1, 1);
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: () => "x",
      painter: () => Promise.resolve(),
      checkSyncRenderers: false
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 50, height: 20 });

    expect(() => renderer.paintContent({ kind: "full" })).not.toThrow();
  });
});

describe("ContentLayerRenderer — gridlines", () => {
  it("does not paint zero-width hidden columns", () => {
    const ctx = new RecordingCanvasContext();
    const layout = new LayoutEngine({
      rowCount: 1,
      columnCount: 2,
      rowHeight: 20,
      columnWidths: [0, 50]
    });
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: (_row, column) => `column-${column}`
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 20 });

    renderer.paintContent({ kind: "full" });

    expect(ctx.calls.filter((call) => call.method === "fillText").map((call) => call.args[0])).toEqual([
      "column-1"
    ]);
    expect(ctx.calls.filter((call) => call.method === "strokeRect")).toEqual([
      { method: "strokeRect", args: [0, 0, 50, 20] }
    ]);
  });

  it("strokes a gridline around each painted cell by default", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(1, 2);
    const renderer = createContentRenderer({ ctx, layout, getCellValue: () => "x" });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 20 });

    renderer.paintContent({ kind: "full" });

    const strokeRectCalls = ctx.calls.filter((call) => call.method === "strokeRect");
    expect(strokeRectCalls).toEqual([
      { method: "strokeRect", args: [0, 0, 50, 20] },
      { method: "strokeRect", args: [50, 0, 50, 20] }
    ]);
  });

  it("does not stroke gridlines when showGridLines is false", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(1, 2);
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: () => "x",
      showGridLines: false
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 20 });

    renderer.paintContent({ kind: "full" });

    expect(ctx.calls.filter((call) => call.method === "strokeRect")).toHaveLength(0);
  });

  it("still strokes a gridline around a cell whose custom painter throws (default-painter fallback path)", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(1, 1);
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: () => "x",
      painter: () => {
        throw new Error("boom");
      },
      onError: () => {}
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 50, height: 20 });

    renderer.paintContent({ kind: "full" });

    expect(ctx.calls.filter((call) => call.method === "strokeRect")).toHaveLength(1);
  });
});

describe("ContentLayerRenderer — frozen row/column separators", () => {
  it("draws no separators when nothing is frozen", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(5, 5);
    const renderer = createContentRenderer({ ctx, layout, getCellValue: () => "x" });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 100 });

    renderer.paintContent({ kind: "full" });

    // fillRect is only ever used for separators in this renderer — none of
    // them should appear when frozenRowCount/frozenColumnCount are both 0.
    expect(ctx.calls.filter((call) => call.method === "fillRect")).toHaveLength(0);
  });

  it("draws a horizontal separator at the frozen row boundary and a vertical one at the frozen column boundary", () => {
    const ctx = new RecordingCanvasContext();
    const layout = new LayoutEngine({
      rowCount: 10,
      columnCount: 10,
      rowHeight: 20,
      columnWidths: Array.from({ length: 10 }, () => 50),
      freeze: { rows: 2, columns: 3 }
    });
    const renderer = createContentRenderer({ ctx, layout, getCellValue: () => "x" });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 200, height: 150 });

    renderer.paintContent({ kind: "full" });

    const fillRectCalls = ctx.calls.filter((call) => call.method === "fillRect");
    // Row separator: y = frozenRowCount(2) * rowHeight(20) = 40, centered
    // 2px band -> starts at 39, spans the viewport width.
    expect(fillRectCalls).toContainEqual({ method: "fillRect", args: [0, 39, 200, 2] });
    // Column separator: x = sum of first 3 column widths = 150, same 2px band.
    expect(fillRectCalls).toContainEqual({ method: "fillRect", args: [149, 0, 2, 150] });
  });

  it("still draws separators for a regions (non-full) repaint", () => {
    const ctx = new RecordingCanvasContext();
    const layout = new LayoutEngine({
      rowCount: 10,
      columnCount: 10,
      rowHeight: 20,
      columnWidths: Array.from({ length: 10 }, () => 50),
      freeze: { rows: 1, columns: 0 }
    });
    const renderer = createContentRenderer({ ctx, layout, getCellValue: () => "x" });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 200, height: 150 });

    renderer.paintContent({
      kind: "regions",
      ranges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }]
    });

    const fillRectCalls = ctx.calls.filter((call) => call.method === "fillRect");
    expect(fillRectCalls).toEqual([{ method: "fillRect", args: [0, 19, 200, 2] }]);
  });
});

describe("ContentLayerRenderer — merged cells", () => {
  function mergeLookup(merge: { startRow: number; startColumn: number; endRow: number; endColumn: number }) {
    return (range: { startRow: number; startColumn: number; endRow: number; endColumn: number }) =>
      mergesIntersectingRange([merge], range);
  }

  it("paints a merged cell once as a single spanning rect using the anchor's value, skipping covered cells", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(3, 2); // rowHeight 20, columnWidths [50, 50]
    const values: Record<string, CellValue> = {
      "0,0": "Anchor",
      "1,0": "ShouldNeverPaint",
      "0,1": "0:1",
      "1,1": "1:1",
      "2,0": "2:0",
      "2,1": "2:1"
    };
    const reads: string[] = [];
    const merge = { startRow: 0, startColumn: 0, endRow: 1, endColumn: 0 };
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: (row, column) => {
        const key = `${row},${column}`;
        reads.push(key);
        return values[key];
      },
      getMergesInRange: mergeLookup(merge)
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 60 });

    renderer.paintContent({ kind: "full" });

    const fillTextCalls = ctx.calls.filter((call) => call.method === "fillText");
    expect(fillTextCalls.map((call) => call.args[0]).sort()).toEqual(
      ["0:1", "1:1", "2:0", "2:1", "Anchor"].sort()
    );
    // The covered cell's own value is never read — it's painted as part of
    // the anchor's spanning rect, not individually.
    expect(reads).not.toContain("1,0");

    // The merged rect spans both rows (height 40) at the anchor's (x, y).
    const clearRectCalls = ctx.calls.filter((call) => call.method === "clearRect");
    const mergeClears = clearRectCalls.filter(
      (call) => call.args[0] === 0 && call.args[1] === 0 && call.args[2] === 50 && call.args[3] === 40
    );
    expect(mergeClears).toHaveLength(1); // painted exactly once, not once per covered row
  });

  it("re-scans and repaints the merge on every paintContent call — the dedupe Set never persists across frames", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(2, 1);
    const merge = { startRow: 0, startColumn: 0, endRow: 1, endColumn: 0 };
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: (row, column) => (row === 0 && column === 0 ? "Anchor" : "covered"),
      getMergesInRange: mergeLookup(merge)
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 50, height: 40 });

    renderer.paintContent({ kind: "full" });
    renderer.paintContent({ kind: "full" });

    const anchorFillTextCalls = ctx.calls.filter(
      (call) => call.method === "fillText" && call.args[0] === "Anchor"
    );
    // If the "already painted this frame" Set were a persistent field
    // instead of local to each #paintFragment call, the second frame would
    // find the merge key already marked painted and skip it.
    expect(anchorFillTextCalls).toHaveLength(2);
  });

  it("behaves exactly as before (no merge scan) when getMergesInRange is not provided", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(2, 1);
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: (row) => `row-${row}`
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 50, height: 40 });

    renderer.paintContent({ kind: "full" });

    const fillTextCalls = ctx.calls.filter((call) => call.method === "fillText");
    expect(fillTextCalls.map((call) => call.args[0])).toEqual(["row-0", "row-1"]);
  });

  it("queries getMergesInRange once per fragment, not once per visible cell — regression sentinel for the scroll-stutter fix", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(50, 40); // 2,000 visible cells if this ever regresses to per-cell queries
    let callCount = 0;
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: () => "x",
      getMergesInRange: () => {
        callCount += 1;
        return [];
      }
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 2000, height: 1000 });

    renderer.paintContent({ kind: "full" });

    // One unfrozen pane ("body") for this layout, so exactly one fragment,
    // hence exactly one range query — not the 2,000 point-queries a per-cell
    // scan would have made even with zero merges anywhere on the sheet.
    expect(callCount).toBe(1);
  });
});

describe("ContentLayerRenderer no-op cases", () => {
  it("does nothing when nothing is visible", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(10, 10);
    const renderer = createContentRenderer({ ctx, layout, getCellValue: () => "x" });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 0, height: 0 });

    renderer.paintContent({ kind: "full" });

    expect(ctx.calls).toHaveLength(0);
  });

  it("does nothing when dirty.kind is none", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(10, 10);
    const renderer = createContentRenderer({ ctx, layout, getCellValue: () => "x" });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 100 });

    renderer.paintContent({ kind: "none" });

    expect(ctx.calls).toHaveLength(0);
  });
});

describe("ContentLayerRenderer — invalid-data mark", () => {
  function marks(ctx: RecordingCanvasContext): Array<readonly unknown[]> {
    // 三角形 = moveTo → lineTo → lineTo → closePath → fill；只取 moveTo 起点
    return ctx.calls.filter((call) => call.method === "moveTo").map((call) => call.args);
  }

  it("draws a small triangle in the top-right corner of each invalid cell only", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(2, 2); // rowHeight 20, columnWidths [50, 50]
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: () => "x",
      isCellInvalid: (row, column) => row === 1 && column === 0
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 40 });

    renderer.paintContent({ kind: "full" });

    // (1, 0)：x 0..50、y 20..40，起点在右上角向左 6px
    expect(marks(ctx)).toEqual([[44, 20]]);
    expect(ctx.calls.filter((call) => call.method === "fill")).toHaveLength(1);
  });

  it("paints nothing extra when no cell is invalid, or when the option is absent", () => {
    const layout = createLayout(2, 2);

    const none = new RecordingCanvasContext();
    const withOption = createContentRenderer({ ctx: none, layout, getCellValue: () => "x", isCellInvalid: () => false });
    withOption.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 40 });
    withOption.paintContent({ kind: "full" });

    const absent = new RecordingCanvasContext();
    const withoutOption = createContentRenderer({ ctx: absent, layout, getCellValue: () => "x" });
    withoutOption.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 40 });
    withoutOption.paintContent({ kind: "full" });

    expect(marks(none)).toEqual([]);
    expect(none.calls).toEqual(absent.calls);
  });

  it("uses the merged rect's top-right corner for a merged cell, painted once", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(3, 2);
    const merge = { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 };
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: () => "x",
      getMergesInRange: (range) => mergesIntersectingRange([merge], range),
      isCellInvalid: (row, column) => row === 0 && column === 0
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 60 });

    renderer.paintContent({ kind: "full" });

    // 合并区 x 0..100：右上角向左 6px
    expect(marks(ctx)).toEqual([[94, 0]]);
  });

  it("does not leak the mark's fill colour into the next cell", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(1, 2);
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: () => "x",
      isCellInvalid: (_row, column) => column === 0
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 20 });

    renderer.paintContent({ kind: "full" });

    const saves = ctx.calls.filter((call) => call.method === "save").length;
    const restores = ctx.calls.filter((call) => call.method === "restore").length;
    expect(saves).toBe(restores);
  });
});

const ARROW = { mode: "arrow" } as const;

describe("ContentLayerRenderer — dropdown arrow", () => {
  function triangles(ctx: RecordingCanvasContext): Array<readonly unknown[]> {
    return ctx.calls.filter((call) => call.method === "moveTo").map((call) => call.args);
  }

  it("draws a down-pointing triangle centred in the arrow zone of each dropdown cell only", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(2, 2); // rowHeight 20, columnWidths [50, 50]
    const renderer = createContentRenderer({
      ctx,
      layout,
      getCellValue: () => "x",
      getCellDropdown: (row, column) => (row === 0 && column === 1 ? ARROW : undefined)
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 40 });

    renderer.paintContent({ kind: "full" });

    // (0, 1)：x 50..100、y 0..20；区宽 18 → 中心 x = 100 - 9 = 91，y = 10
    expect(triangles(ctx)).toEqual([[87.5, 8]]);
    const lineTos = ctx.calls.filter((call) => call.method === "lineTo").map((call) => call.args);
    expect(lineTos).toEqual([
      [94.5, 8],
      [91, 12.5] // 尖朝下
    ]);
  });

  it("is independent of the invalid mark: both can appear on the same cell", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createContentRenderer({
      ctx,
      layout: createLayout(1, 1),
      getCellValue: () => "x",
      isCellInvalid: () => true,
      getCellDropdown: () => ARROW
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 50, height: 20 });

    renderer.paintContent({ kind: "full" });

    expect(ctx.calls.filter((call) => call.method === "fill")).toHaveLength(2);
  });

  it("uses the merged rect's right edge for a merged cell, painted once", () => {
    const ctx = new RecordingCanvasContext();
    const merge = { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 };
    const renderer = createContentRenderer({
      ctx,
      layout: createLayout(3, 2),
      getCellValue: () => "x",
      getMergesInRange: (range) => mergesIntersectingRange([merge], range),
      getCellDropdown: (row, column) => (row === 0 && column === 0 ? ARROW : undefined)
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 60 });

    renderer.paintContent({ kind: "full" });

    // 合并区 x 0..100、y 0..40 → 中心 (91, 20)
    expect(triangles(ctx)).toEqual([[87.5, 18]]);
  });

  it("paints nothing extra when no cell has a dropdown, or when the option is absent", () => {
    const layout = createLayout(2, 2);
    const none = new RecordingCanvasContext();
    const a = createContentRenderer({ ctx: none, layout, getCellValue: () => "x", getCellDropdown: () => undefined });
    a.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 40 });
    a.paintContent({ kind: "full" });

    const absent = new RecordingCanvasContext();
    const b = createContentRenderer({ ctx: absent, layout, getCellValue: () => "x" });
    b.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 40 });
    b.paintContent({ kind: "full" });

    expect(triangles(none)).toEqual([]);
    expect(none.calls).toEqual(absent.calls);
  });

  it("keeps save/restore balanced", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createContentRenderer({
      ctx,
      layout: createLayout(1, 2),
      getCellValue: () => "x",
      getCellDropdown: () => ARROW
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 20 });

    renderer.paintContent({ kind: "full" });

    expect(ctx.calls.filter((c) => c.method === "save")).toHaveLength(ctx.calls.filter((c) => c.method === "restore").length);
  });
});

describe("ContentLayerRenderer — room for the dropdown arrow", () => {
  function capture() {
    const rects: Array<{ row: number; column: number; rect: { x: number; y: number; width: number; height: number } }> = [];
    const painter = (context: { row: number; column: number; rect: { x: number; y: number; width: number; height: number } }) => {
      rects.push({ row: context.row, column: context.column, rect: { ...context.rect } });
    };
    return { rects, painter };
  }

  it("hands the painter a rect whose right edge is pulled in by the arrow zone, only for dropdown cells", () => {
    const ctx = new RecordingCanvasContext();
    const { rects, painter } = capture();
    const renderer = createContentRenderer({
      ctx,
      layout: createLayout(1, 2), // 每列 50 宽、20 高
      getCellValue: () => "x",
      painter,
      getCellDropdown: (_row, column) => (column === 1 ? ARROW : undefined)
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 20 });

    renderer.paintContent({ kind: "full" });

    expect(rects.find((r) => r.column === 0)?.rect).toEqual({ x: 0, y: 0, width: 50, height: 20 });
    expect(rects.find((r) => r.column === 1)?.rect).toEqual({ x: 50, y: 0, width: 32, height: 20 }); // 50 - 18
  });

  it("clips the painter to the narrowed rect so long text stops before the arrow", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createContentRenderer({
      ctx,
      layout: createLayout(1, 1),
      getCellValue: () => "a very long value",
      getCellDropdown: () => ARROW
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 50, height: 20 });

    renderer.paintContent({ kind: "full" });

    const clipAt = ctx.calls.findIndex((call, i) => call.method === "clip" && ctx.calls[i - 1]?.method === "rect" && ctx.calls[i - 1]?.args[2] === 32);
    const textAt = ctx.calls.findIndex((call) => call.method === "fillText");
    expect(clipAt).toBeGreaterThan(-1);
    expect(textAt).toBeGreaterThan(clipAt); // 先裁剪、后画字
    // 裁剪区是窄矩形（x 0..32），不是整格
    expect(ctx.calls[clipAt - 1]?.args).toEqual([0, 0, 32, 20]);
  });

  it("does not clip or narrow a cell without a dropdown", () => {
    const withoutOption = new RecordingCanvasContext();
    const a = createContentRenderer({ ctx: withoutOption, layout: createLayout(1, 1), getCellValue: () => "x" });
    a.setViewport({ scrollTop: 0, scrollLeft: 0, width: 50, height: 20 });
    a.paintContent({ kind: "full" });

    const noDropdown = new RecordingCanvasContext();
    const b = createContentRenderer({
      ctx: noDropdown,
      layout: createLayout(1, 1),
      getCellValue: () => "x",
      getCellDropdown: () => undefined
    });
    b.setViewport({ scrollTop: 0, scrollLeft: 0, width: 50, height: 20 });
    b.paintContent({ kind: "full" });

    expect(noDropdown.calls).toEqual(withoutOption.calls);
  });

  it("moves right-aligned text left of the arrow", () => {
    const at = (dropdown: boolean): number => {
      const ctx = new RecordingCanvasContext();
      const renderer = createContentRenderer({
        ctx,
        layout: createLayout(1, 1),
        getCellValue: () => "x",
        painter: createColumnAwarePainter(() => [
          {
            width: 50,
            typeName: "text",
            format: (value: unknown) => String(value ?? ""),
            clipboard: { copy: true, paste: true },
            wrap: false,
            sortable: false,
            filterable: false,
            version: 1
          }
        ]),
        getCellStyle: () => ({ align: "right" }),
        getCellDropdown: () => (dropdown ? ARROW : undefined)
      });
      renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 50, height: 20 });
      renderer.paintContent({ kind: "full" });
      return ctx.calls.find((call) => call.method === "fillText")!.args[1] as number;
    };

    expect(at(false) - at(true)).toBe(18);
  });

  it("keeps the full rect when the cell is no wider than the arrow zone", () => {
    const ctx = new RecordingCanvasContext();
    const { rects, painter } = capture();
    const renderer = createContentRenderer({
      ctx,
      layout: new LayoutEngine({ rowCount: 1, columnCount: 1, rowHeight: 20, columnWidths: [16] }),
      getCellValue: () => "x",
      painter,
      getCellDropdown: () => ARROW
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 16, height: 20 });

    renderer.paintContent({ kind: "full" });

    expect(rects[0]?.rect.width).toBe(16);
  });

  it("still paints the arrow and the border against the full rect", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createContentRenderer({
      ctx,
      layout: createLayout(1, 1),
      getCellValue: () => "x",
      getCellDropdown: () => ARROW
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 50, height: 20 });

    renderer.paintContent({ kind: "full" });

    const moveTo = ctx.calls.find((call) => call.method === "moveTo");
    expect(moveTo?.args).toEqual([87.5 - 50, 8]); // 中心 x = 50 - 9 = 41，三角左顶点 37.5
  });
});

describe("ContentLayerRenderer — dropdown display modes", () => {
  function render(options: {
    readonly dropdown: { mode: "text" | "arrow" | "custom"; color?: string } | undefined;
    readonly value?: CellValue;
    readonly width?: number;
    readonly style?: { background?: string };
  }) {
    const ctx = new RecordingCanvasContext();
    let painted = 0;
    const width = options.width ?? 50;
    const renderer = createContentRenderer({
      ctx,
      layout: new LayoutEngine({ rowCount: 1, columnCount: 1, rowHeight: 20, columnWidths: [width] }),
      getCellValue: () => options.value ?? "ab",
      painter: () => {
        painted += 1;
      },
      ...(options.style !== undefined ? { getCellStyle: () => options.style } : {}),
      getCellDropdown: () => options.dropdown
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width, height: 20 });
    renderer.paintContent({ kind: "full" });
    return { ctx, painted: () => painted };
  }

  const fills = (ctx: RecordingCanvasContext) =>
    ctx.calls.filter((call) => call.method === "set:fillStyle").map((call) => call.args[0]);

  describe("text", () => {
    it("draws no arrow and does not narrow the painter's rect", () => {
      const { ctx, painted } = render({ dropdown: { mode: "text" } });
      expect(ctx.calls.filter((call) => call.method === "moveTo")).toEqual([]);
      expect(painted()).toBe(1);
    });

    it("uses the item's colour as the cell background", () => {
      const { ctx } = render({ dropdown: { mode: "text", color: "#d93025" } });
      const fillRect = ctx.calls.findIndex((call) => call.method === "fillRect");
      expect(fillRect).toBeGreaterThan(-1);
      expect(ctx.calls[fillRect - 1]).toMatchObject({ method: "set:fillStyle", args: ["#d93025"] });
    });
  });

  describe("arrow", () => {
    it("uses the item's colour as the cell background too, but a style background wins", () => {
      const coloured = render({ dropdown: { mode: "arrow", color: "#188038" } });
      expect(fills(coloured.ctx)).toContain("#188038");

      const styled = render({ dropdown: { mode: "arrow", color: "#188038" }, style: { background: "#ffffff" } });
      expect(fills(styled.ctx)).toContain("#ffffff");
      expect(fills(styled.ctx)).not.toContain("#188038");
    });
  });

  describe("custom (capsule)", () => {
    // 格子 50x20，文字 "ab" = 14 宽（记录器每字 7）→ 胶囊 x 6..42、y 3..17
    it("draws a rounded capsule with the text and the arrow inside, and skips the painter", () => {
      const { ctx, painted } = render({ dropdown: { mode: "custom" } });

      expect(painted()).toBe(0);
      expect(ctx.calls.filter((call) => call.method === "arc")).toHaveLength(4); // 四个圆角
      const text = ctx.calls.find((call) => call.method === "fillText");
      expect(text?.args).toEqual(["ab", 10, 10]);
      // 箭头三角在胶囊右端：中心 x = 6 + 36 - 4 - 7 = 31
      const arrow = ctx.calls.filter((call) => call.method === "moveTo").at(-1);
      expect(arrow?.args).toEqual([31 - 3.5, 8]);
    });

    it("fills the capsule with the item colour, or the default grey when it has none", () => {
      expect(fills(render({ dropdown: { mode: "custom", color: "#4285f4" } }).ctx)).toContain("#4285f4");
      expect(fills(render({ dropdown: { mode: "custom" } }).ctx)).toContain("#e8eaed");
    });

    it("does not tint the whole cell — the colour belongs to the capsule", () => {
      const { ctx } = render({ dropdown: { mode: "custom", color: "#4285f4" } });
      expect(ctx.calls.filter((call) => call.method === "fillRect")).toEqual([]);
    });

    it("clips long text to the capsule's text area", () => {
      const { ctx } = render({ dropdown: { mode: "custom" }, value: "a very long value indeed" });
      const textIndex = ctx.calls.findIndex((call) => call.method === "fillText");
      const clipIndex = ctx.calls.slice(0, textIndex).map((call) => call.method).lastIndexOf("clip");
      // 胶囊撑满 50 - 12 = 38 宽，文字区 38 - 8 - 14 = 16；最靠近 fillText 的那次裁剪（更早的是面板自己的）
      expect(ctx.calls[clipIndex - 1]?.args).toEqual([10, 3, 16, 14]);
    });

    it("draws just the arrow capsule for an empty value", () => {
      const { ctx } = render({ dropdown: { mode: "custom" }, value: "" });
      expect(ctx.calls.filter((call) => call.method === "fillText")).toEqual([]);
      expect(ctx.calls.filter((call) => call.method === "arc")).toHaveLength(4);
    });

    it("falls back to the plain arrow when the capsule does not fit", () => {
      const { ctx, painted } = render({ dropdown: { mode: "custom" }, width: 30 });
      expect(painted()).toBe(1);
      expect(ctx.calls.filter((call) => call.method === "arc")).toEqual([]);
      expect(ctx.calls.filter((call) => call.method === "moveTo")).toHaveLength(1);
    });

    it("keeps save/restore balanced", () => {
      const { ctx } = render({ dropdown: { mode: "custom", color: "#4285f4" } });
      expect(ctx.calls.filter((c) => c.method === "save")).toHaveLength(ctx.calls.filter((c) => c.method === "restore").length);
    });
  });
});
