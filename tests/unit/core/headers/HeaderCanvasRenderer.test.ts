import { describe, expect, it } from "vitest";

import type { CanvasHeaderRenderer, HeaderState } from "../../../../src/core/cells";
import {
  DEFAULT_HEADER_SORT_ZONE_MIN_WIDTH,
  DEFAULT_HEADER_SORT_ZONE_WIDTH,
  defaultHeaderHitTest,
  HeaderCanvasRenderer,
  HeaderManager
} from "../../../../src/core/headers";
import { LayoutEngine, SkeletonSnapshot, ViewportSnapshot } from "../../../../src/core/layout";
import {
  HEADER_SORT_INDICATOR_ACTIVE_COLOR,
  HEADER_SORT_INDICATOR_COLOR
} from "../../../../src/core/render/theme";
import type { ViewportState } from "../../../../src/core/layout";
import { SceneSnapshot } from "../../../../src/core/scene";
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

const VIEWPORT: ViewportState = { scrollTop: 0, scrollLeft: 0, width: 200, height: 100 };

type HeaderRendererTestOptions = Omit<ConstructorParameters<typeof HeaderCanvasRenderer>[0], "getScene"> & {
  readonly layout: LayoutEngine;
  readonly columnHeaderHeight: number;
  readonly groupHeaderHeight?: number;
};

function createScene(
  layout: LayoutEngine,
  viewport: ViewportState = VIEWPORT,
  columnHeaderHeight = 0,
  groupHeaderHeight = 0,
  groupDepth = 0
): SceneSnapshot {
  const totalHeight = columnHeaderHeight + groupHeaderHeight * groupDepth;
  return new SceneSnapshot({
    version: 1,
    skeleton: SkeletonSnapshot.fromLayout(layout, 1),
    viewport: new ViewportSnapshot({ version: 1, ...viewport }),
    surfaceMetrics: {
      columnHeaderHeight: totalHeight,
      columnHeaderBands: [
        ...Array.from({ length: groupDepth }, (_unused, index) => ({
          role: "group" as const,
          level: groupDepth - 1 - index,
          y: index * groupHeaderHeight,
          height: groupHeaderHeight
        })),
        {
          role: "leaf" as const,
          y: groupHeaderHeight * groupDepth,
          height: columnHeaderHeight
        }
      ],
      rowHeaderWidth: 0
    }
  });
}

function createRenderer(
  options: HeaderRendererTestOptions
): HeaderCanvasRenderer & {
  readonly getScene: () => SceneSnapshot;
  readonly setViewport: (viewport: ViewportState) => void;
} {
  let scene = createScene(
    options.layout,
    VIEWPORT,
    options.columnHeaderHeight,
    options.groupHeaderHeight,
    options.headerManager.groupDepth
  );
  const renderer = new HeaderCanvasRenderer({
    ctx: options.ctx,
    getScene: () => scene,
    headerManager: options.headerManager,
    ...(options.getHeaderState !== undefined ? { getHeaderState: options.getHeaderState } : {})
  });
  return Object.assign(renderer, {
    getScene: () => scene,
    setViewport(viewport: ViewportState) {
      scene = createScene(
        options.layout,
        viewport,
        options.columnHeaderHeight,
        options.groupHeaderHeight,
        options.headerManager.groupDepth
      );
    }
  });
}

describe("HeaderCanvasRenderer — leaf headers (先写测试 #1: title/sort/filter state)", () => {
  it("draws each visible column's title by default", () => {
    const ctx = new RecordingCanvasContext();
    const headerManager = new HeaderManager({
      columns: [{ id: "a", header: "A" }, { id: "b", header: "B" }]
    });
    const renderer = createRenderer({
      ctx,
      layout: createLayout(2),
      headerManager,
      columnHeaderHeight: 24
    });
    renderer.setViewport(VIEWPORT);

    renderer.paint();

    const titles = ctx.calls.filter((call) => call.method === "fillText").map((call) => call.args[0]);
    expect(titles).toContain("A");
    expect(titles).toContain("B");
  });

  it("reflects sorted state in the default painter's output", () => {
    const ctx = new RecordingCanvasContext();
    const headerManager = new HeaderManager({ columns: [{ id: "a", header: "A" }] });
    const renderer = createRenderer({
      ctx,
      layout: createLayout(1),
      headerManager,
      columnHeaderHeight: 24,
      getHeaderState: () => ({ sorted: "asc" })
    });
    renderer.setViewport(VIEWPORT);

    renderer.paint();

    // The direction is a painted caret now, not a suffix fused into the title:
    // it has to land in the trailing zone the click happens in, and a fused
    // string is left-aligned with the label at the other end of the cell.
    const drawn = ctx.calls.filter((call) => call.method === "fillText").map((call) => call.args[0]);
    expect(drawn).toEqual(["A"]);
    expect(ctx.calls.some((call) => call.method === "fill")).toBe(true);
  });

  it("reflects filtered state in the default painter's output", () => {
    const ctx = new RecordingCanvasContext();
    const headerManager = new HeaderManager({ columns: [{ id: "a", header: "A" }] });
    const renderer = createRenderer({
      ctx,
      layout: createLayout(1),
      headerManager,
      columnHeaderHeight: 24,
      getHeaderState: () => ({ filtered: true })
    });
    renderer.setViewport(VIEWPORT);

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
    const renderer = createRenderer({
      ctx,
      layout: createLayout(1),
      headerManager,
      columnHeaderHeight: 24
    });
    renderer.setViewport(VIEWPORT);

    renderer.paint();

    expect(paintedWith).toHaveLength(1);
    expect(ctx.calls.filter((call) => call.method === "fillText")).toHaveLength(0);
  });

  it("only paints columns currently in the frozen or visible set (virtualized)", () => {
    const ctx = new RecordingCanvasContext();
    const headerManager = new HeaderManager({
      columns: Array.from({ length: 5 }, (_unused, index) => ({ id: `c${index}`, header: `C${index}` }))
    });
    const renderer = createRenderer({
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

  it("clips a partially scrolled leaf header to its canonical scrollable pane", () => {
    const ctx = new RecordingCanvasContext();
    const headerManager = new HeaderManager({
      columns: [{ id: "frozen", header: "Frozen" }, { id: "scroll", header: "Scroll" }]
    });
    const renderer = createRenderer({
      ctx,
      layout: createLayout(2, { columns: 1 }),
      headerManager,
      columnHeaderHeight: 24
    });
    renderer.setViewport({ scrollTop: 0, scrollLeft: 20, width: 100, height: 100 });

    renderer.paint();

    // Column 1 has surface rect x=30..80, while its pane starts at x=50.
    expect(ctx.calls).toContainEqual({ method: "rect", args: [50, 0, 30, 24] });
  });
});

describe("HeaderCanvasRenderer — two-level groups (先写测试 #4: split at frozen boundary)", () => {
  it("draws one group title when it doesn't straddle the frozen boundary", () => {
    const ctx = new RecordingCanvasContext();
    const headerManager = new HeaderManager({
      columns: [{ id: "a" }, { id: "b" }],
      groups: [{ id: "g1", title: "Group", columns: [0, 1] }]
    });
    const renderer = createRenderer({
      ctx,
      layout: createLayout(2),
      headerManager,
      columnHeaderHeight: 24,
      groupHeaderHeight: 20
    });
    renderer.setViewport(VIEWPORT);

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
    const renderer = createRenderer({
      ctx,
      layout: createLayout(3, { columns: 1 }),
      headerManager,
      columnHeaderHeight: 24,
      groupHeaderHeight: 20
    });
    renderer.setViewport(VIEWPORT);

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
    const renderer = createRenderer({
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

describe("HeaderCanvasRenderer — arbitrary-depth nested groups (Step 19)", () => {
  function threeLevelHeaderManager() {
    // Same tree HeaderManager.test.ts uses: level 0 "cat-a"[0,1]/"cat-b"[2,3],
    // level 1 "dept"[0,3], level 2 "root"[0,3].
    return new HeaderManager({
      columns: [{ id: "a" }, { id: "b" }, { id: "c" }, { id: "d" }],
      groups: [
        {
          id: "root",
          title: "Root",
          children: [
            {
              id: "dept",
              title: "Dept",
              children: [
                { id: "cat-a", title: "Cat A", columns: [0, 1] },
                { id: "cat-b", title: "Cat B", columns: [2, 3] }
              ]
            }
          ]
        }
      ]
    });
  }

  it("reads total header height from the Scene surface", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createRenderer({
      ctx,
      layout: createLayout(4),
      headerManager: threeLevelHeaderManager(),
      columnHeaderHeight: 24,
      groupHeaderHeight: 20
    });

    expect(renderer.getScene().getSurface("column-header").rootRect.height).toBe(24 + 20 * 3);
  });

  it("draws every level's group titles, the right number of times", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createRenderer({
      ctx,
      layout: createLayout(4),
      headerManager: threeLevelHeaderManager(),
      columnHeaderHeight: 24,
      groupHeaderHeight: 20
    });
    renderer.setViewport(VIEWPORT);

    renderer.paint();

    const titles = ctx.calls.filter((call) => call.method === "fillText").map((call) => call.args[0]);
    expect(titles.filter((title) => title === "Root")).toHaveLength(1);
    expect(titles.filter((title) => title === "Dept")).toHaveLength(1);
    expect(titles.filter((title) => title === "Cat A")).toHaveLength(1);
    expect(titles.filter((title) => title === "Cat B")).toHaveLength(1);
  });

  it("stacks levels outward-to-inward: outermost (level 2) at y=0, leaf-adjacent (level 0) just above the leaf row", () => {
    const ctx = new RecordingCanvasContext();
    const renderer = createRenderer({
      ctx,
      layout: createLayout(4),
      headerManager: threeLevelHeaderManager(),
      columnHeaderHeight: 24,
      groupHeaderHeight: 20
    });
    renderer.setViewport(VIEWPORT);

    renderer.paint();

    const fillRectYs = ctx.calls.filter((call) => call.method === "fillRect").map((call) => call.args[1]);
    // Root (level 2, outermost): y=0. Dept (level 1): y=20. Cat A/B (level 0): y=40.
    // Leaf row: y=60 (groupHeaderHeight * groupDepth).
    expect(fillRectYs).toContain(0);
    expect(fillRectYs).toContain(20);
    expect(fillRectYs).toContain(40);
    expect(fillRectYs).toContain(60);
  });

  it("a shallower branch leaves its higher levels blank (no fillText for it there)", () => {
    const ctx = new RecordingCanvasContext();
    const headerManager = new HeaderManager({
      columns: [{ id: "a" }, { id: "b" }, { id: "c" }, { id: "d" }],
      groups: [
        { id: "shallow", title: "Shallow", columns: [0, 1] },
        { id: "wrapper", title: "Wrapper", children: [{ id: "deep", title: "Deep", columns: [2, 3] }] }
      ]
    });
    const renderer = createRenderer({
      ctx,
      layout: createLayout(4),
      headerManager,
      columnHeaderHeight: 24,
      groupHeaderHeight: 20
    });
    renderer.setViewport(VIEWPORT);

    renderer.paint();

    const titles = ctx.calls.filter((call) => call.method === "fillText").map((call) => call.args[0]);
    expect(titles).toContain("Shallow");
    expect(titles).toContain("Wrapper");
    // "Shallow" only ever appears once — it has nothing painted at level 1.
    expect(titles.filter((title) => title === "Shallow")).toHaveLength(1);
  });

  it("depth 0 (no groups) and depth 1 (Step 9's flat case) behave exactly as before", () => {
    const ctxNoGroups = new RecordingCanvasContext();
    const noGroups = createRenderer({
      ctx: ctxNoGroups,
      layout: createLayout(2),
      headerManager: new HeaderManager({ columns: [{ id: "a" }, { id: "b" }] }),
      columnHeaderHeight: 24,
      groupHeaderHeight: 20
    });
    expect(noGroups.getScene().getSurface("column-header").rootRect.height).toBe(24);

    const ctxFlat = new RecordingCanvasContext();
    const flat = createRenderer({
      ctx: ctxFlat,
      layout: createLayout(2),
      headerManager: new HeaderManager({
        columns: [{ id: "a" }, { id: "b" }],
        groups: [{ id: "g", title: "G", columns: [0, 1] }]
      }),
      columnHeaderHeight: 24,
      groupHeaderHeight: 20
    });
    expect(flat.getScene().getSurface("column-header").rootRect.height).toBe(24 + 20);
  });
});

describe("HeaderCanvasRenderer — background/font/border styling", () => {
  it("fills a background rect behind each leaf header before drawing its title", () => {
    const ctx = new RecordingCanvasContext();
    const headerManager = new HeaderManager({ columns: [{ id: "a", header: "A" }] });
    const renderer = createRenderer({
      ctx,
      layout: createLayout(1),
      headerManager,
      columnHeaderHeight: 24
    });
    renderer.setViewport(VIEWPORT);

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
    const renderer = createRenderer({
      ctx,
      layout: createLayout(1),
      headerManager,
      columnHeaderHeight: 24
    });
    renderer.setViewport(VIEWPORT);

    renderer.paint();

    const fontCall = ctx.calls.find((call) => call.method === "set:font");
    expect(typeof fontCall?.args[0]).toBe("string");
    expect((fontCall?.args[0] as string).startsWith("600 ")).toBe(true);
  });

  it("draws a bottom border spanning the viewport width after the header band", () => {
    const ctx = new RecordingCanvasContext();
    const headerManager = new HeaderManager({ columns: [{ id: "a", header: "A" }] });
    const renderer = createRenderer({
      ctx,
      layout: createLayout(1),
      headerManager,
      columnHeaderHeight: 24
    });
    renderer.setViewport(VIEWPORT);

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
    const renderer = createRenderer({
      ctx,
      layout: createLayout(2),
      headerManager,
      columnHeaderHeight: 24,
      groupHeaderHeight: 20
    });
    renderer.setViewport(VIEWPORT);

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
    const renderer = createRenderer({
      ctx,
      layout: createLayout(1),
      headerManager,
      columnHeaderHeight: 24
    });
    renderer.setViewport(VIEWPORT);

    renderer.paint();

    const saves = ctx.calls.filter((call) => call.method === "save").length;
    expect(saves).toBeGreaterThan(0);
    expect(ctx.calls.filter((call) => call.method === "restore")).toHaveLength(saves);
    expect(ctx.calls.filter((call) => call.method === "clip")).toHaveLength(saves);
  });
});

describe("defaultHeaderHitTest", () => {
  const rect = (width: number) => ({ x: 0, y: 0, width, height: 24 });

  it("claims the trailing strip of a wide-enough header for toggle-sort", () => {
    const wide = rect(120);

    expect(defaultHeaderHitTest({ rect: wide, point: { x: 119, y: 10 } }))
      .toEqual({ type: "toggle-sort" });
    expect(defaultHeaderHitTest({
      rect: wide,
      point: { x: 120 - DEFAULT_HEADER_SORT_ZONE_WIDTH, y: 10 }
    })).toEqual({ type: "toggle-sort" });
    // ...and leaves the rest of the cell to column selection.
    expect(defaultHeaderHitTest({
      rect: wide,
      point: { x: 120 - DEFAULT_HEADER_SORT_ZONE_WIDTH - 1, y: 10 }
    })).toBeUndefined();
  });

  it("gives a column too narrow to hold both zones no sort zone at all", () => {
    // Without a width floor, `x >= width - 18` is true for *every* point in
    // a column narrower than 18px: the whole header becomes a sort zone and
    // the column can never be selected by clicking it again.
    for (const width of [1, 6, 17, DEFAULT_HEADER_SORT_ZONE_MIN_WIDTH - 1]) {
      for (const x of [0, width / 2, width - 0.5]) {
        expect(defaultHeaderHitTest({ rect: rect(width), point: { x, y: 10 } }))
          .toBeUndefined();
      }
    }
    // Exactly at the floor both zones still fit, so the strip comes back.
    const floor = DEFAULT_HEADER_SORT_ZONE_MIN_WIDTH;
    expect(defaultHeaderHitTest({ rect: rect(floor), point: { x: floor - 1, y: 10 } }))
      .toEqual({ type: "toggle-sort" });
    expect(defaultHeaderHitTest({ rect: rect(floor), point: { x: 0, y: 10 } }))
      .toBeUndefined();
  });
});

describe("HeaderCanvasRenderer — the sort carets live in the zone you click", () => {
  const SORTABLE = { sortable: true } as const;

  function paintOne(state: HeaderState): RecordingCanvasContext {
    const ctx = new RecordingCanvasContext();
    const headerManager = new HeaderManager({ columns: [{ id: "a", header: "A" }] });
    const renderer = createRenderer({
      ctx,
      layout: createLayout(1),
      headerManager,
      columnHeaderHeight: 24,
      getHeaderState: () => state
    });
    renderer.setViewport(VIEWPORT);
    renderer.paint();
    return ctx;
  }

  /**
   * The colour each filled triangle was painted in, in paint order — the sort
   * carets are the only `fill()`ed paths the header draws, and `set:fillStyle`
   * immediately precedes each one.
   */
  function caretColors(ctx: RecordingCanvasContext): string[] {
    const colors: string[] = [];
    let pending: string | undefined;
    for (const call of ctx.calls) {
      if (call.method === "set:fillStyle") {
        pending = String(call.args[0]);
      } else if (call.method === "fill" && pending !== undefined) {
        colors.push(pending);
      }
    }
    return colors;
  }

  it("draws both carets on an unsorted sortable column, so the column advertises sorting", () => {
    // The bug this pins: previously nothing at all was drawn until *after* a
    // sort, so a user had no way to discover the column was sortable — and the
    // clickable zone was at the trailing edge while the only feedback appeared
    // next to the title.
    expect(caretColors(paintOne(SORTABLE)))
      .toEqual([HEADER_SORT_INDICATOR_COLOR, HEADER_SORT_INDICATOR_COLOR]);
  });

  it("draws no carets at all on a column that is not sortable", () => {
    expect(caretColors(paintOne({ sortable: false }))).toEqual([]);
  });

  it("colours only the ascending caret when sorted ascending, keeping the other visible", () => {
    // Both stay drawn: the inactive caret is what says the opposite direction
    // is one more click away. Removing it would lose that.
    expect(caretColors(paintOne({ ...SORTABLE, sorted: "asc" })))
      .toEqual([HEADER_SORT_INDICATOR_ACTIVE_COLOR, HEADER_SORT_INDICATOR_COLOR]);
  });

  it("colours only the descending caret when sorted descending", () => {
    expect(caretColors(paintOne({ ...SORTABLE, sorted: "desc" })))
      .toEqual([HEADER_SORT_INDICATOR_COLOR, HEADER_SORT_INDICATOR_ACTIVE_COLOR]);
  });

  it("keeps the direction out of the title string", () => {
    // It used to be appended to the label and left-aligned, which is why it
    // never appeared where the click had to happen.
    const drawn = paintOne({ ...SORTABLE, sorted: "desc" }).calls
      .filter((call) => call.method === "fillText")
      .map((call) => String(call.args[0]));

    expect(drawn).toEqual(["A"]);
  });

  it("paints the carets inside the trailing sort zone, not beside the title", () => {
    const ctx = paintOne({ ...SORTABLE, sorted: "asc" });
    const columnWidth = 50; // createLayout's fixed column width
    const apexes = ctx.calls
      .filter((call) => call.method === "moveTo")
      .map((call) => Number(call.args[0]));

    expect(apexes).toHaveLength(2);
    for (const x of apexes) {
      expect(x).toBeGreaterThan(columnWidth - DEFAULT_HEADER_SORT_ZONE_WIDTH);
    }
  });

  it("stacks the two carets, ascending above descending", () => {
    const ctx = paintOne(SORTABLE);
    const apexY = ctx.calls
      .filter((call) => call.method === "moveTo")
      .map((call) => Number(call.args[1]));

    expect(apexY[0]).toBeLessThan(apexY[1] as number);
  });
});

describe("defaultHeaderHitTest — sortable columns only", () => {
  it("refuses the sort zone on a column that is not sortable", () => {
    expect(defaultHeaderHitTest({
      rect: { x: 0, y: 0, width: 120, height: 24 },
      point: { x: 119, y: 10 },
      sortable: false
    })).toBeUndefined();
  });
});
