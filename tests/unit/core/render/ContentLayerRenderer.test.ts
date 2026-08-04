import { describe, expect, it } from "vitest";

import {
  ContentLayerRenderer,
  CyExcelError,
  CyExcelErrorCode,
  LayoutEngine
} from "../../../../src/core";
import { RecordingCanvasContext } from "../../../helpers/RecordingCanvasContext";

function createLayout(rowCount: number, columnCount: number): LayoutEngine {
  return new LayoutEngine({
    rowCount,
    columnCount,
    rowHeight: 20,
    columnWidths: Array.from({ length: columnCount }, () => 50)
  });
}

describe("ContentLayerRenderer.paintContent — dirty regions", () => {
  it("only paints the cell in a single-cell dirty region", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(10, 10);
    const renderer = new ContentLayerRenderer({
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

  it("draws an empty string for the default painter when the value is null or undefined", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(1, 2);
    const values: Record<number, unknown> = { 0: undefined, 1: null };
    const renderer = new ContentLayerRenderer({
      ctx,
      layout,
      getCellValue: (_row, column) => values[column]
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
    const renderer = new ContentLayerRenderer({ ctx, layout, getCellValue: () => "x" });
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
    const renderer = new ContentLayerRenderer({ ctx, layout, getCellValue: () => "x" });
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

describe("ContentLayerRenderer error isolation", () => {
  it("falls back to the default painter and reports onError when a custom painter throws", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(3, 1);
    const errors: Array<{ error: unknown; cell: { row: number; column: number } }> = [];
    const renderer = new ContentLayerRenderer({
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
    const renderer = new ContentLayerRenderer({
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
    const renderer = new ContentLayerRenderer({
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

describe("ContentLayerRenderer no-op cases", () => {
  it("does nothing when nothing is visible", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(10, 10);
    const renderer = new ContentLayerRenderer({ ctx, layout, getCellValue: () => "x" });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 0, height: 0 });

    renderer.paintContent({ kind: "full" });

    expect(ctx.calls).toHaveLength(0);
  });

  it("does nothing when dirty.kind is none", () => {
    const ctx = new RecordingCanvasContext();
    const layout = createLayout(10, 10);
    const renderer = new ContentLayerRenderer({ ctx, layout, getCellValue: () => "x" });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 100 });

    renderer.paintContent({ kind: "none" });

    expect(ctx.calls).toHaveLength(0);
  });
});
