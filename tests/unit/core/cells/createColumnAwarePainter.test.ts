import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { createColumnAwarePainter } from "../../../../src/core/cells";
import type { CanvasCellRenderer, RuntimeColumn } from "../../../../src/core/cells";
import { LayoutEngine, SkeletonSnapshot, ViewportSnapshot } from "../../../../src/core/layout";
import { ContentLayerRenderer } from "../../../../src/core/render";
import { SceneSnapshot } from "../../../../src/core/scene";
import { RecordingCanvasContext } from "../../../helpers/RecordingCanvasContext";

function textColumn(overrides: Partial<RuntimeColumn> = {}): RuntimeColumn {
  return {
    
    
    width: 100,
    typeName: "text",
    format: (value) => (value === undefined || value === null ? "" : String(value)),
    clipboard: { copy: true, paste: true },
    wrap: false,
    sortable: true,
    filterable: true,
    version: 1,
    ...overrides
  };
}

const rect = { x: 10, y: 20, width: 100, height: 24 };

function createScene(layout: LayoutEngine): SceneSnapshot {
  return new SceneSnapshot({
    version: 1,
    skeleton: SkeletonSnapshot.fromLayout(layout, 1),
    viewport: new ViewportSnapshot({ version: 1, scrollTop: 0, scrollLeft: 0, width: 100, height: 20 })
  });
}

describe("createColumnAwarePainter — no custom renderer", () => {
  it("draws the column's format() output as text", () => {
    const ctx = new RecordingCanvasContext();
    const column = textColumn({ format: () => "formatted-value" });
    const painter = createColumnAwarePainter(() => [column]);

    painter({ ctx, rect, row: 0, column: 0, value: "raw" });

    const fillTextCalls = ctx.calls.filter((call) => call.method === "fillText");
    expect(fillTextCalls).toHaveLength(1);
    expect(fillTextCalls[0]?.args[0]).toBe("formatted-value");
  });

  it("sets ctx.font explicitly before drawing, instead of relying on whatever font the canvas state happens to already be in", () => {
    const ctx = new RecordingCanvasContext();
    const column = textColumn({ format: () => "x" });
    const painter = createColumnAwarePainter(() => [column]);

    painter({ ctx, rect, row: 0, column: 0, value: "raw" });

    const fontCall = ctx.calls.find((call) => call.method === "set:font");
    expect(fontCall).toBeDefined();
    const fontIndex = ctx.calls.indexOf(fontCall!);
    const fillTextIndex = ctx.calls.findIndex((call) => call.method === "fillText");
    expect(fontIndex).toBeLessThan(fillTextIndex);
  });

  it("wraps long text into top-aligned lines and stops at the cell bottom", () => {
    const ctx = new RecordingCanvasContext();
    const column = textColumn({ wrap: true, format: () => "abcd efgh" });
    const painter = createColumnAwarePainter(() => [column]);

    painter({ ctx, rect: { ...rect, width: 36, height: 44 }, row: 0, column: 0, value: "raw" });

    const fillTextCalls = ctx.calls.filter((call) => call.method === "fillText");
    expect(fillTextCalls.map((call) => call.args)).toEqual([
      ["abcd", 14, 32],
      ["efgh", 14, 48]
    ]);
  });
});

describe("createColumnAwarePainter — reads the columns list live, not a one-time snapshot", () => {
  it("picks up a column root published after the painter was created", () => {
    const ctx = new RecordingCanvasContext();
    let columns: readonly RuntimeColumn[] = [textColumn({ format: () => "before" })];
    const painter = createColumnAwarePainter(() => columns);

    painter({ ctx, rect, row: 0, column: 0, value: "x" });
    columns = [textColumn({ format: () => "after" })];
    painter({ ctx, rect, row: 0, column: 0, value: "x" });

    const fillTextCalls = ctx.calls.filter((call) => call.method === "fillText");
    expect(fillTextCalls.map((call) => call.args[0])).toEqual(["before", "after"]);
  });
});

describe("createColumnAwarePainter — out-of-range column", () => {
  it("draws nothing when the context's column index has no resolved RuntimeColumn", () => {
    const ctx = new RecordingCanvasContext();
    const painter = createColumnAwarePainter(() => [textColumn()]);

    expect(() => painter({ ctx, rect, row: 0, column: 5, value: "x" })).not.toThrow();
    expect(ctx.calls).toHaveLength(0);
  });
});

describe("createColumnAwarePainter — canvas renderer", () => {
  it("delegates painting to the column's CanvasCellRenderer", () => {
    const ctx = new RecordingCanvasContext();
    const paintedWith: unknown[] = [];
    const renderer: CanvasCellRenderer = {
      kind: "canvas",
      paint: (context) => {
        paintedWith.push(context);
      }
    };
    const column = textColumn({ renderer });
    const painter = createColumnAwarePainter(() => [column]);

    painter({ ctx, rect, row: 3, column: 0, value: "raw" });

    expect(paintedWith).toHaveLength(1);
    // Never falls through to the default text painter once a custom
    // renderer is present.
    expect(ctx.calls.filter((call) => call.method === "fillText")).toHaveLength(0);
  });

  it("hands the renderer a read-only context — no store/commit access, only ctx/rect/row/column/value", () => {
    const ctx = new RecordingCanvasContext();
    let received: unknown;
    const renderer: CanvasCellRenderer = {
      kind: "canvas",
      paint: (context) => {
        received = context;
      }
    };
    const column = textColumn({ renderer });
    const painter = createColumnAwarePainter(() => [column]);

    painter({ ctx, rect, row: 1, column: 0, value: 42 });

    expect(Object.keys(received as object).sort()).toEqual(["column", "ctx", "rect", "row", "value"]);
    expect((received as { column: RuntimeColumn }).column).toBe(column);
    expect((received as { value: unknown }).value).toBe(42);
  });
});

describe("createColumnAwarePainter — dom renderer", () => {
  it("draws nothing on the canvas — the DOM overlay owns the cell entirely", () => {
    const ctx = new RecordingCanvasContext();
    const column = textColumn({
      renderer: {
        kind: "dom",
        pointerOwnership: "surface",
        create: () => ({ mount() {}, update() {}, destroy() {} })
      }
    });
    const painter = createColumnAwarePainter(() => [column]);

    painter({ ctx, rect, row: 0, column: 0, value: "x" });

    expect(ctx.calls).toHaveLength(0);
  });
});

describe("createColumnAwarePainter — error isolation through ContentLayerRenderer", () => {
  it("falls back to the default painter for a throwing renderer without affecting other cells", () => {
    const ctx = new RecordingCanvasContext();
    const layout = new LayoutEngine({
      rowCount: 1,
      columnCount: 2,
      rowHeight: 20,
      columnWidths: [50, 50]
    });
    const throwingRenderer: CanvasCellRenderer = {
      kind: "canvas",
      paint: () => {
        throw new Error("boom");
      }
    };
    const columns: readonly RuntimeColumn[] = [
      textColumn({ renderer: throwingRenderer }),
      textColumn({ format: () => "still-fine" })
    ];
    const errors: unknown[] = [];
    const renderer = new ContentLayerRenderer({
      ctx,
      getScene: () => createScene(layout),
      getCellValue: () => "x",
      painter: createColumnAwarePainter(() => columns),
      onError: (error) => errors.push(error)
    });

    renderer.paintContent({ kind: "full" });

    expect(errors).toHaveLength(1);
    const fillTextCalls = ctx.calls.filter((call) => call.method === "fillText");
    // Column 0 ("bad") falls back to ContentLayerRenderer's own default
    // painter (draws the raw cell value); column 1 ("good") is unaffected
    // and still goes through createColumnAwarePainter's format() path.
    expect(fillTextCalls.map((call) => call.args[0])).toEqual(["x", "still-fine"]);
  });
});

describe("createColumnAwarePainter — auto-linkified text", () => {
  it("paints a cell whose full value is a URL in link color with an underline", () => {
    const ctx = new RecordingCanvasContext();
    const column = textColumn({ format: (value) => String(value) });
    const painter = createColumnAwarePainter(() => [column]);

    painter({ ctx, rect, row: 0, column: 0, value: "https://example.com" });

    const fillStyleCalls = ctx.calls.filter((call) => call.method === "set:fillStyle");
    expect(fillStyleCalls.at(-1)?.args[0]).toBe("#1a73e8");
    expect(ctx.calls.some((call) => call.method === "stroke")).toBe(true);
  });

  it("does not linkify ordinary text", () => {
    const ctx = new RecordingCanvasContext();
    const column = textColumn({ format: (value) => String(value) });
    const painter = createColumnAwarePainter(() => [column]);

    painter({ ctx, rect, row: 0, column: 0, value: "工业温度传感器" });

    const fillStyleCalls = ctx.calls.filter((call) => call.method === "set:fillStyle");
    expect(fillStyleCalls.at(-1)?.args[0]).not.toBe("#1a73e8");
    expect(ctx.calls.some((call) => call.method === "stroke")).toBe(false);
  });
});

describe("createColumnAwarePainter — cell-level image (context.image)", () => {
  // createImageCellRenderer() needs a global `Image` to construct — this
  // environment runs under Node, not jsdom, so it must be stubbed (same
  // stub as imageCellRenderer.test.ts).
  class FakeImage {
    width = 10;
    height = 10;
    onload: (() => void) | null = null;
    onerror: (() => void) | null = null;
    #src = "";
    set src(value: string) {
      this.#src = value;
    }
    get src(): string {
      return this.#src;
    }
  }

  beforeEach(() => {
    vi.stubGlobal("Image", FakeImage);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("paints the cell's own image instead of the column's default text formatting", () => {
    const ctx = new RecordingCanvasContext();
    const column = textColumn({ format: () => "should-not-appear" });
    const painter = createColumnAwarePainter(() => [column]);

    painter({ ctx, rect, row: 0, column: 0, value: "not-a-url", image: "https://example.com/a.png" });

    expect(ctx.calls.filter((call) => call.method === "fillText")).toHaveLength(0);
    // Placeholder (or the decoded image) always fills a rect — the image path
    // ran, the text path didn't.
    expect(ctx.calls.some((call) => call.method === "fillRect")).toBe(true);
  });

  it("takes priority over the column's own canvas renderer — the cell decides, not the column", () => {
    const ctx = new RecordingCanvasContext();
    const paintedWith: unknown[] = [];
    const renderer: CanvasCellRenderer = {
      kind: "canvas",
      paint: (context) => {
        paintedWith.push(context);
      }
    };
    const column = textColumn({ renderer });
    const painter = createColumnAwarePainter(() => [column]);

    painter({ ctx, rect, row: 0, column: 0, value: "raw", image: "https://example.com/b.png" });

    expect(paintedWith).toHaveLength(0);
    expect(ctx.calls.some((call) => call.method === "fillRect")).toBe(true);
  });

  it("falls back to the column's normal dispatch when the cell has no image", () => {
    const ctx = new RecordingCanvasContext();
    const column = textColumn({ format: () => "plain-text" });
    const painter = createColumnAwarePainter(() => [column]);

    painter({ ctx, rect, row: 0, column: 0, value: "raw" });

    const fillTextCalls = ctx.calls.filter((call) => call.method === "fillText");
    expect(fillTextCalls.map((call) => call.args[0])).toEqual(["plain-text"]);
  });
});

