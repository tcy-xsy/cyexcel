import { describe, expect, it } from "vitest";

import { LayoutEngine, RowHeaderCanvasRenderer } from "../../../../src/core";
import { RecordingCanvasContext } from "../../../helpers/RecordingCanvasContext";

function createLayout(freeze?: { rows?: number; columns?: number }): LayoutEngine {
  return new LayoutEngine({
    rowCount: 20,
    columnCount: 3,
    rowHeight: 20,
    columnWidths: [50, 50, 50],
    ...(freeze ? { freeze } : {})
  });
}

function fillTexts(ctx: RecordingCanvasContext): unknown[] {
  return ctx.calls.filter((call) => call.method === "fillText").map((call) => call.args[0]);
}

describe("RowHeaderCanvasRenderer — row numbers", () => {
  it("draws 1-based row numbers for the visible rows", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new RowHeaderCanvasRenderer({
      ctx,
      layout: createLayout(),
      rowHeaderWidth: 40,
      headerHeight: 0
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 150, height: 100 });

    renderer.paint();

    const texts = fillTexts(ctx);
    expect(texts).toContain("1");
    expect(texts).toContain("2");
  });

  it("draws nothing when rowHeaderWidth is 0", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new RowHeaderCanvasRenderer({
      ctx,
      layout: createLayout(),
      rowHeaderWidth: 0,
      headerHeight: 0
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 150, height: 100 });

    renderer.paint();

    expect(ctx.calls).toHaveLength(0);
  });

  it("draws frozen rows regardless of scroll position", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new RowHeaderCanvasRenderer({
      ctx,
      layout: createLayout({ rows: 1 }),
      rowHeaderWidth: 40,
      headerHeight: 0
    });
    renderer.setViewport({ scrollTop: 200, scrollLeft: 0, width: 150, height: 100 });

    renderer.paint();

    expect(fillTexts(ctx)).toContain("1");
  });

  it("shifts non-frozen row numbers up by scrollTop", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new RowHeaderCanvasRenderer({
      ctx,
      layout: createLayout(),
      rowHeaderWidth: 40,
      headerHeight: 0
    });
    renderer.setViewport({ scrollTop: 40, scrollLeft: 0, width: 150, height: 100 });

    renderer.paint();

    // Row index 2 (label "3") sits at absolute y=40; after scrolling 40px it
    // should be painted at local y=0 (+ half the row height, for vertical centering).
    const call = ctx.calls.find((c) => c.method === "fillText" && c.args[0] === "3");
    expect(call).toBeDefined();
    expect(call?.args[2]).toBe(10); // rowHeight/2
  });
});

describe("RowHeaderCanvasRenderer — corner cell", () => {
  it("does not paint a corner cell when headerHeight is 0", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new RowHeaderCanvasRenderer({
      ctx,
      layout: createLayout(),
      rowHeaderWidth: 40,
      headerHeight: 0
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 150, height: 100 });

    renderer.paint();

    const fillRectCalls = ctx.calls.filter((c) => c.method === "fillRect");
    expect(fillRectCalls.some((c) => c.args[1] === 0 && c.args[3] === 0)).toBe(false);
  });

  it("paints a corner cell at (0, 0, rowHeaderWidth, headerHeight) when headerHeight > 0", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new RowHeaderCanvasRenderer({
      ctx,
      layout: createLayout(),
      rowHeaderWidth: 40,
      headerHeight: 24
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 150, height: 100 });

    renderer.paint();

    const fillRectCalls = ctx.calls.filter((c) => c.method === "fillRect");
    expect(fillRectCalls.some((c) => c.args[0] === 0 && c.args[1] === 0 && c.args[2] === 40 && c.args[3] === 24)).toBe(
      true
    );
  });

  it("offsets row numbers below the corner band when headerHeight > 0", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = new RowHeaderCanvasRenderer({
      ctx,
      layout: createLayout(),
      rowHeaderWidth: 40,
      headerHeight: 24
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 150, height: 100 });

    renderer.paint();

    const call = ctx.calls.find((c) => c.method === "fillText" && c.args[0] === "1");
    expect(call?.args[2]).toBe(24 + 10); // headerHeight + rowHeight / 2
  });
});
