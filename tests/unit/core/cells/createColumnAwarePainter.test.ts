import { describe, expect, it } from "vitest";

import { ContentLayerRenderer, LayoutEngine, createColumnAwarePainter } from "../../../../src/core";
import type { CanvasCellRenderer, RuntimeColumn } from "../../../../src/core";
import { RecordingCanvasContext } from "../../../helpers/RecordingCanvasContext";

function textColumn(overrides: Partial<RuntimeColumn> = {}): RuntimeColumn {
  return {
    id: "c",
    field: "c",
    width: 100,
    typeName: "text",
    format: (value) => (value === undefined || value === null ? "" : String(value)),
    clipboard: { copy: true, paste: true },
    version: 1,
    ...overrides
  };
}

const rect = { x: 10, y: 20, width: 100, height: 24 };

describe("createColumnAwarePainter — no custom renderer", () => {
  it("draws the column's format() output as text", () => {
    const ctx = new RecordingCanvasContext();
    const column = textColumn({ format: () => "formatted-value" });
    const painter = createColumnAwarePainter([column]);

    painter({ ctx, rect, row: 0, column: 0, value: "raw" });

    const fillTextCalls = ctx.calls.filter((call) => call.method === "fillText");
    expect(fillTextCalls).toHaveLength(1);
    expect(fillTextCalls[0]?.args[0]).toBe("formatted-value");
  });

  it("sets ctx.font explicitly before drawing, instead of relying on whatever font the canvas state happens to already be in", () => {
    const ctx = new RecordingCanvasContext();
    const column = textColumn({ format: () => "x" });
    const painter = createColumnAwarePainter([column]);

    painter({ ctx, rect, row: 0, column: 0, value: "raw" });

    const fontCall = ctx.calls.find((call) => call.method === "set:font");
    expect(fontCall).toBeDefined();
    const fontIndex = ctx.calls.indexOf(fontCall!);
    const fillTextIndex = ctx.calls.findIndex((call) => call.method === "fillText");
    expect(fontIndex).toBeLessThan(fillTextIndex);
  });
});

describe("createColumnAwarePainter — out-of-range column", () => {
  it("draws nothing when the context's column index has no resolved RuntimeColumn", () => {
    const ctx = new RecordingCanvasContext();
    const painter = createColumnAwarePainter([textColumn()]);

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
    const painter = createColumnAwarePainter([column]);

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
    const painter = createColumnAwarePainter([column]);

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
      renderer: { kind: "dom", create: () => ({ mount() {}, update() {}, destroy() {} }) }
    });
    const painter = createColumnAwarePainter([column]);

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
      textColumn({ id: "bad", renderer: throwingRenderer }),
      textColumn({ id: "good", format: () => "still-fine" })
    ];
    const errors: unknown[] = [];
    const renderer = new ContentLayerRenderer({
      ctx,
      layout,
      getCellValue: () => "x",
      painter: createColumnAwarePainter(columns),
      onError: (error) => errors.push(error)
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 20 });

    renderer.paintContent({ kind: "full" });

    expect(errors).toHaveLength(1);
    const fillTextCalls = ctx.calls.filter((call) => call.method === "fillText");
    // Column 0 ("bad") falls back to ContentLayerRenderer's own default
    // painter (draws the raw cell value); column 1 ("good") is unaffected
    // and still goes through createColumnAwarePainter's format() path.
    expect(fillTextCalls.map((call) => call.args[0])).toEqual(["x", "still-fine"]);
  });
});
