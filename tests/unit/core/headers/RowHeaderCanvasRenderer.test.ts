import { describe, expect, it } from "vitest";

import { RowHeaderCanvasRenderer } from "../../../../src/core/headers";
import { LayoutEngine, SkeletonSnapshot, ViewportSnapshot } from "../../../../src/core/layout";
import type { ViewportState } from "../../../../src/core/layout";
import { SceneSnapshot } from "../../../../src/core/scene";
import { ROW_REORDER_INDICATOR_COLOR } from "../../../../src/core/render/theme";
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

const VIEWPORT: ViewportState = { scrollTop: 0, scrollLeft: 0, width: 150, height: 100 };

type RowHeaderRendererTestOptions = Omit<ConstructorParameters<typeof RowHeaderCanvasRenderer>[0], "getScene"> & {
  readonly layout: LayoutEngine;
  readonly rowHeaderWidth: number;
  readonly headerHeight: number;
};

function createScene(
  layout: LayoutEngine,
  viewport: ViewportState = VIEWPORT,
  surfaceMetrics = { columnHeaderHeight: 0, rowHeaderWidth: 0 }
): SceneSnapshot {
  return new SceneSnapshot({
    version: 1,
    skeleton: SkeletonSnapshot.fromLayout(layout, 1),
    viewport: new ViewportSnapshot({ version: 1, ...viewport }),
    surfaceMetrics
  });
}

function createRenderer(
  options: RowHeaderRendererTestOptions
): RowHeaderCanvasRenderer & { readonly setViewport: (viewport: ViewportState) => void } {
  const surfaceMetrics = {
    columnHeaderHeight: options.headerHeight,
    rowHeaderWidth: options.rowHeaderWidth
  };
  let scene = createScene(options.layout, VIEWPORT, surfaceMetrics);
  const renderer = new RowHeaderCanvasRenderer({
    ctx: options.ctx,
    getScene: () => scene,
    ...(options.getReorderIndicatorRow !== undefined ? { getReorderIndicatorRow: options.getReorderIndicatorRow } : {})
  });
  return Object.assign(renderer, {
    setViewport(viewport: ViewportState) {
      scene = createScene(options.layout, viewport, surfaceMetrics);
    }
  });
}

function fillTexts(ctx: RecordingCanvasContext): unknown[] {
  return ctx.calls.filter((call) => call.method === "fillText").map((call) => call.args[0]);
}

describe("RowHeaderCanvasRenderer — row numbers", () => {
  it("draws 1-based row numbers for the visible rows", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createRenderer({
      ctx,
      layout: createLayout(),
      rowHeaderWidth: 40,
      headerHeight: 0
    });
    renderer.setViewport(VIEWPORT);

    renderer.paint();

    const texts = fillTexts(ctx);
    expect(texts).toContain("1");
    expect(texts).toContain("2");
  });

  it("draws nothing when rowHeaderWidth is 0", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createRenderer({
      ctx,
      layout: createLayout(),
      rowHeaderWidth: 0,
      headerHeight: 0
    });
    renderer.setViewport(VIEWPORT);

    renderer.paint();

    expect(ctx.calls).toHaveLength(0);
  });

  it("draws frozen rows regardless of scroll position", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createRenderer({
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
    const renderer = createRenderer({
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

  it("clips a partially scrolled row header to its canonical scrollable pane", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createRenderer({
      ctx,
      layout: createLayout({ rows: 1 }),
      rowHeaderWidth: 40,
      headerHeight: 0
    });
    renderer.setViewport({ scrollTop: 7, scrollLeft: 0, width: 150, height: 100 });

    renderer.paint();

    // Row 1 has surface rect y=13..33, while its pane starts at y=20.
    expect(ctx.calls).toContainEqual({ method: "rect", args: [0, 20, 40, 13] });
  });
});

describe("RowHeaderCanvasRenderer — corner cell", () => {
  it("does not paint a corner cell when headerHeight is 0", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createRenderer({
      ctx,
      layout: createLayout(),
      rowHeaderWidth: 40,
      headerHeight: 0
    });
    renderer.setViewport(VIEWPORT);

    renderer.paint();

    const fillRectCalls = ctx.calls.filter((c) => c.method === "fillRect");
    expect(fillRectCalls.some((c) => c.args[1] === 0 && c.args[3] === 0)).toBe(false);
  });

  it("paints a corner cell at (0, 0, rowHeaderWidth, headerHeight) when headerHeight > 0", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createRenderer({
      ctx,
      layout: createLayout(),
      rowHeaderWidth: 40,
      headerHeight: 24
    });
    renderer.setViewport(VIEWPORT);

    renderer.paint();

    const fillRectCalls = ctx.calls.filter((c) => c.method === "fillRect");
    expect(fillRectCalls.some((c) => c.args[0] === 0 && c.args[1] === 0 && c.args[2] === 40 && c.args[3] === 24)).toBe(
      true
    );
  });

  it("offsets row numbers below the corner band when headerHeight > 0", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createRenderer({
      ctx,
      layout: createLayout(),
      rowHeaderWidth: 40,
      headerHeight: 24
    });
    renderer.setViewport(VIEWPORT);

    renderer.paint();

    const call = ctx.calls.find((c) => c.method === "fillText" && c.args[0] === "1");
    expect(call?.args[2]).toBe(24 + 10); // headerHeight + rowHeight / 2
  });
});

describe("RowHeaderCanvasRenderer — row-reorder drag indicator", () => {
  it("draws nothing extra when getReorderIndicatorRow is not provided", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createRenderer({
      ctx,
      layout: createLayout(),
      rowHeaderWidth: 40,
      headerHeight: 0
    });
    renderer.setViewport(VIEWPORT);

    renderer.paint();

    expect(ctx.calls.some((c) => c.args[0] === ROW_REORDER_INDICATOR_COLOR)).toBe(false);
  });

  it("draws nothing extra when getReorderIndicatorRow returns null", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createRenderer({
      ctx,
      layout: createLayout(),
      rowHeaderWidth: 40,
      headerHeight: 0,
      getReorderIndicatorRow: () => null
    });
    renderer.setViewport(VIEWPORT);

    renderer.paint();

    expect(ctx.calls.some((c) => c.args[0] === ROW_REORDER_INDICATOR_COLOR)).toBe(false);
  });

  it("draws a full-width bar at the boundary y when headerHeight is 0", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createRenderer({
      ctx,
      layout: createLayout(),
      rowHeaderWidth: 40,
      headerHeight: 0,
      getReorderIndicatorRow: () => 3 // boundary before row 3 -> absolute y=60
    });
    renderer.setViewport(VIEWPORT);

    renderer.paint();

    const styleIndex = ctx.calls.findIndex(
      (c) => c.method === "set:fillStyle" && c.args[0] === ROW_REORDER_INDICATOR_COLOR
    );
    expect(styleIndex).toBeGreaterThanOrEqual(0);
    const indicatorRect = ctx.calls
      .slice(styleIndex)
      .find((c) => c.method === "fillRect") as { args: readonly number[] } | undefined;
    expect(indicatorRect).toBeDefined();
    expect(indicatorRect?.args[0]).toBe(0);
    expect(indicatorRect?.args[2]).toBe(40); // full rowHeaderWidth
    // Centered on y=60.
    const [, y, , height] = indicatorRect!.args;
    expect(y! + height! / 2).toBe(60);
  });

  it("offsets the indicator by headerHeight, like the row numbers themselves", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createRenderer({
      ctx,
      layout: createLayout(),
      rowHeaderWidth: 40,
      headerHeight: 24,
      getReorderIndicatorRow: () => 3 // absolute y=60, +24 header band
    });
    renderer.setViewport(VIEWPORT);

    renderer.paint();

    // The indicator is always painted last, after every row/corner fillRect
    // — find the LAST fillRect following our color, since it's the only one
    // still in the recording after that final set:fillStyle.
    const matchingStyleIndices = ctx.calls
      .map((c, index) => ({ c, index }))
      .filter(({ c }) => c.method === "set:fillStyle" && c.args[0] === ROW_REORDER_INDICATOR_COLOR)
      .map(({ index }) => index);
    const styleIndex = matchingStyleIndices[matchingStyleIndices.length - 1] ?? -1;
    const indicatorRect = ctx.calls
      .slice(styleIndex)
      .find((c) => c.method === "fillRect") as { args: readonly number[] } | undefined;
    expect(indicatorRect).toBeDefined();
    const [, y, , height] = indicatorRect!.args;
    expect(y! + height! / 2).toBe(60 + 24);
  });

  it("accounts for scroll when the boundary is outside the frozen band", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createRenderer({
      ctx,
      layout: createLayout({ rows: 1 }),
      rowHeaderWidth: 40,
      headerHeight: 0,
      getReorderIndicatorRow: () => 5 // absolute y=100
    });
    renderer.setViewport({ scrollTop: 40, scrollLeft: 0, width: 150, height: 100 });

    renderer.paint();

    const matchingStyleIndices = ctx.calls
      .map((c, index) => ({ c, index }))
      .filter(({ c }) => c.method === "set:fillStyle" && c.args[0] === ROW_REORDER_INDICATOR_COLOR)
      .map(({ index }) => index);
    const styleIndex = matchingStyleIndices[matchingStyleIndices.length - 1] ?? -1;
    const indicatorRect = ctx.calls
      .slice(styleIndex)
      .find((c) => c.method === "fillRect") as { args: readonly number[] } | undefined;
    expect(indicatorRect).toBeDefined();
    const [, y, , height] = indicatorRect!.args;
    expect(y! + height! / 2).toBe(60); // 100 - scrollTop(40)
  });
});
