import { describe, expect, it } from "vitest";

import { HeaderCanvasRenderer, HeaderManager, LayoutEngine } from "../../../../src/core";
import type { CanvasHeaderRenderer } from "../../../../src/core";
import { RecordingCanvasContext } from "../../../helpers/RecordingCanvasContext";

function createLayout(columnCount: number, freeze?: { rows?: number; columns?: number }): LayoutEngine {
  return new LayoutEngine({
    rowCount: 10,
    columnCount,
    rowHeight: 20,
    columnWidths: Array.from({ length: columnCount }, () => 50),
    ...(freeze ? { freeze } : {})
  });
}

describe("HeaderCanvasRenderer — leaf headers (先写测试 #1: title/sort/filter state)", () => {
  it("draws each visible column's title by default", () => {
    const ctx = new RecordingCanvasContext();
    const headerManager = new HeaderManager({
      columns: [{ id: "a", header: "A" }, { id: "b", header: "B" }]
    });
    const renderer = new HeaderCanvasRenderer({
      ctx,
      layout: createLayout(2),
      headerManager,
      columnHeaderHeight: 24
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 200, height: 100 });

    renderer.paint();

    const titles = ctx.calls.filter((call) => call.method === "fillText").map((call) => call.args[0]);
    expect(titles).toContain("A");
    expect(titles).toContain("B");
  });

  it("reflects sorted state in the default painter's output", () => {
    const ctx = new RecordingCanvasContext();
    const headerManager = new HeaderManager({ columns: [{ id: "a", header: "A" }] });
    const renderer = new HeaderCanvasRenderer({
      ctx,
      layout: createLayout(1),
      headerManager,
      columnHeaderHeight: 24,
      getHeaderState: () => ({ sorted: "asc" })
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 200, height: 100 });

    renderer.paint();

    const titles = ctx.calls.filter((call) => call.method === "fillText").map((call) => call.args[0]);
    expect(titles.some((title) => typeof title === "string" && title.includes("A") && title !== "A")).toBe(
      true
    );
  });

  it("reflects filtered state in the default painter's output", () => {
    const ctx = new RecordingCanvasContext();
    const headerManager = new HeaderManager({ columns: [{ id: "a", header: "A" }] });
    const renderer = new HeaderCanvasRenderer({
      ctx,
      layout: createLayout(1),
      headerManager,
      columnHeaderHeight: 24,
      getHeaderState: () => ({ filtered: true })
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 200, height: 100 });

    renderer.paint();

    const titles = ctx.calls.filter((call) => call.method === "fillText").map((call) => call.args[0]);
    expect(titles.some((title) => typeof title === "string" && title !== "A" && title.includes("A"))).toBe(
      true
    );
  });

  it("delegates to a custom CanvasHeaderRenderer instead of the default painter", () => {
    const ctx = new RecordingCanvasContext();
    const paintedWith: unknown[] = [];
    const custom: CanvasHeaderRenderer = {
      kind: "canvas",
      paint: (context) => paintedWith.push(context)
    };
    const headerManager = new HeaderManager({
      columns: [{ id: "a", header: { title: "A", renderer: custom } }]
    });
    const renderer = new HeaderCanvasRenderer({
      ctx,
      layout: createLayout(1),
      headerManager,
      columnHeaderHeight: 24
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 200, height: 100 });

    renderer.paint();

    expect(paintedWith).toHaveLength(1);
    expect(ctx.calls.filter((call) => call.method === "fillText")).toHaveLength(0);
  });

  it("only paints columns currently in the frozen or visible set (virtualized)", () => {
    const ctx = new RecordingCanvasContext();
    const headerManager = new HeaderManager({
      columns: Array.from({ length: 5 }, (_unused, index) => ({ id: `c${index}`, header: `C${index}` }))
    });
    const renderer = new HeaderCanvasRenderer({
      ctx,
      layout: createLayout(5),
      headerManager,
      columnHeaderHeight: 24
    });
    // Only 2 columns (100px) fit in a 100px-wide viewport with 50px columns.
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 100, height: 100 });

    renderer.paint();

    const titles = ctx.calls.filter((call) => call.method === "fillText").map((call) => call.args[0]);
    expect(titles).not.toContain("C4");
  });
});

describe("HeaderCanvasRenderer — two-level groups (先写测试 #4: split at frozen boundary)", () => {
  it("draws one group title when it doesn't straddle the frozen boundary", () => {
    const ctx = new RecordingCanvasContext();
    const headerManager = new HeaderManager({
      columns: [{ id: "a" }, { id: "b" }],
      groups: [{ id: "g1", title: "Group", columns: [0, 1] }]
    });
    const renderer = new HeaderCanvasRenderer({
      ctx,
      layout: createLayout(2),
      headerManager,
      columnHeaderHeight: 24,
      groupHeaderHeight: 20
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 200, height: 100 });

    renderer.paint();

    const titles = ctx.calls.filter((call) => call.method === "fillText").map((call) => call.args[0]);
    expect(titles.filter((title) => title === "Group")).toHaveLength(1);
  });

  it("draws a group straddling the frozen boundary as two fragments, both showing the group's title", () => {
    const ctx = new RecordingCanvasContext();
    const headerManager = new HeaderManager({
      columns: [{ id: "a" }, { id: "b" }, { id: "c" }],
      groups: [{ id: "g1", title: "Group", columns: [0, 1, 2] }]
    });
    const renderer = new HeaderCanvasRenderer({
      ctx,
      layout: createLayout(3, { columns: 1 }),
      headerManager,
      columnHeaderHeight: 24,
      groupHeaderHeight: 20
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 200, height: 100 });

    renderer.paint();

    const titles = ctx.calls.filter((call) => call.method === "fillText").map((call) => call.args[0]);
    expect(titles.filter((title) => title === "Group")).toHaveLength(2);
  });

  it("keeps a frozen group fragment's x unshifted while the scrollable fragment scrolls", () => {
    const ctx = new RecordingCanvasContext();
    const headerManager = new HeaderManager({
      columns: [{ id: "a" }, { id: "b" }, { id: "c" }],
      groups: [{ id: "g1", title: "Group", columns: [0, 1, 2] }]
    });
    const renderer = new HeaderCanvasRenderer({
      ctx,
      layout: createLayout(3, { columns: 1 }),
      headerManager,
      columnHeaderHeight: 24,
      groupHeaderHeight: 20
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 20, width: 200, height: 100 });

    renderer.paint();

    const clearRectCalls = ctx.calls.filter((call) => call.method === "clearRect");
    // Frozen fragment (column 0): absolute x=0, unshifted despite scrollLeft.
    expect(clearRectCalls.some((call) => call.args[0] === 0)).toBe(true);
  });
});

describe("HeaderCanvasRenderer — background/font/border styling", () => {
  it("fills a background rect behind each leaf header before drawing its title", () => {
    const ctx = new RecordingCanvasContext();
    const headerManager = new HeaderManager({ columns: [{ id: "a", header: "A" }] });
    const renderer = new HeaderCanvasRenderer({
      ctx,
      layout: createLayout(1),
      headerManager,
      columnHeaderHeight: 24
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 200, height: 100 });

    renderer.paint();

    const fillRectCalls = ctx.calls.filter((call) => call.method === "fillRect");
    expect(fillRectCalls).toContainEqual({ method: "fillRect", args: [0, 0, 50, 24] });
    const fillRectIndex = ctx.calls.findIndex((call) => call.method === "fillRect");
    const fillTextIndex = ctx.calls.findIndex((call) => call.method === "fillText");
    expect(fillRectIndex).toBeLessThan(fillTextIndex);
  });

  it("uses a bold font for the default leaf-header painter", () => {
    const ctx = new RecordingCanvasContext();
    const headerManager = new HeaderManager({ columns: [{ id: "a", header: "A" }] });
    const renderer = new HeaderCanvasRenderer({
      ctx,
      layout: createLayout(1),
      headerManager,
      columnHeaderHeight: 24
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 200, height: 100 });

    renderer.paint();

    const fontCall = ctx.calls.find((call) => call.method === "set:font");
    expect(typeof fontCall?.args[0]).toBe("string");
    expect((fontCall?.args[0] as string).startsWith("600 ")).toBe(true);
  });

  it("draws a bottom border spanning the viewport width after the header band", () => {
    const ctx = new RecordingCanvasContext();
    const headerManager = new HeaderManager({ columns: [{ id: "a", header: "A" }] });
    const renderer = new HeaderCanvasRenderer({
      ctx,
      layout: createLayout(1),
      headerManager,
      columnHeaderHeight: 24
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 200, height: 100 });

    renderer.paint();

    const fillRectCalls = ctx.calls.filter((call) => call.method === "fillRect");
    // totalHeight (24, no group row) - 1 (half the 2px border thickness).
    expect(fillRectCalls).toContainEqual({ method: "fillRect", args: [0, 23, 200, 2] });
  });

  it("draws a background fill for a group fragment too", () => {
    const ctx = new RecordingCanvasContext();
    const headerManager = new HeaderManager({
      columns: [{ id: "a" }, { id: "b" }],
      groups: [{ id: "g1", title: "Group", columns: [0, 1] }]
    });
    const renderer = new HeaderCanvasRenderer({
      ctx,
      layout: createLayout(2),
      headerManager,
      columnHeaderHeight: 24,
      groupHeaderHeight: 20
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 200, height: 100 });

    renderer.paint();

    const fillRectCalls = ctx.calls.filter((call) => call.method === "fillRect");
    // Group fragment spans both columns: x=0, y=0, width=100, height=20.
    expect(fillRectCalls).toContainEqual({ method: "fillRect", args: [0, 0, 100, 20] });
  });
});

describe("HeaderCanvasRenderer — clipping", () => {
  it("wraps each painted region in save/clip/restore", () => {
    const ctx = new RecordingCanvasContext();
    const headerManager = new HeaderManager({ columns: [{ id: "a", header: "A" }] });
    const renderer = new HeaderCanvasRenderer({
      ctx,
      layout: createLayout(1),
      headerManager,
      columnHeaderHeight: 24
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 0, width: 200, height: 100 });

    renderer.paint();

    const saves = ctx.calls.filter((call) => call.method === "save").length;
    expect(saves).toBeGreaterThan(0);
    expect(ctx.calls.filter((call) => call.method === "restore")).toHaveLength(saves);
    expect(ctx.calls.filter((call) => call.method === "clip")).toHaveLength(saves);
  });
});
