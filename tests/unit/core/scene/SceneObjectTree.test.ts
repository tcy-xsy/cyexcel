import { describe, expect, it } from "vitest";

import { LayoutEngine, SkeletonSnapshot, ViewportSnapshot } from "../../../../src/core/layout";
import { SceneSnapshot } from "../../../../src/core/scene";

type AxisBehavior = "fixed" | "scroll";

interface PaneContract {
  readonly id: string;
  readonly surface: string;
  readonly clipRect: { readonly x: number; readonly y: number; readonly width: number; readonly height: number };
  readonly rows: { readonly start: number; readonly count: number };
  readonly columns: { readonly start: number; readonly count: number };
  readonly frozen: { readonly row: boolean; readonly column: boolean };
  readonly scroll: { readonly x: AxisBehavior; readonly y: AxisBehavior };
  readonly translation: { readonly x: number; readonly y: number };
}

interface SurfaceContract {
  readonly id: string;
  readonly rootRect: { readonly x: number; readonly y: number; readonly width: number; readonly height: number };
  readonly layers: readonly { readonly id: string; readonly nodes: readonly unknown[] }[];
  readonly panes: readonly PaneContract[];
}

interface SceneRootContract {
  readonly surfaces: readonly SurfaceContract[];
}

type SceneWithObjectTree = Omit<SceneSnapshot, "root"> & { readonly root: SceneRootContract };

interface FragmentWithPane {
  readonly pane: PaneContract;
  readonly surface: string;
  readonly clipRect: PaneContract["clipRect"];
  readonly rows: PaneContract["rows"];
  readonly columns: PaneContract["columns"];
  readonly frozenRow: boolean;
  readonly frozenColumn: boolean;
}

function createScene(): SceneWithObjectTree {
  const layout = new LayoutEngine({
    rowCount: 10,
    columnCount: 5,
    rowHeight: 20,
    columnWidths: [50, 50, 50, 50, 50],
    freeze: { rows: 1, columns: 1 }
  });
  return new SceneSnapshot({
    version: 1,
    skeleton: SkeletonSnapshot.fromLayout(layout, 1),
    viewport: new ViewportSnapshot({
      version: 1,
      scrollTop: 7,
      scrollLeft: 13,
      width: 180,
      height: 90
    }),
    surfaceMetrics: {
      columnHeaderHeight: 60,
      rowHeaderWidth: 40
    }
  }) as SceneWithObjectTree;
}

function indexById<T extends { readonly id: string }>(items: readonly T[]): Map<string, T> {
  expect(new Set(items.map((item) => item.id)).size).toBe(items.length);
  return new Map(items.map((item) => [item.id, item]));
}

describe("SceneSnapshot object tree contract", () => {
  it("publishes the five physical surfaces with root-space bounds and virtual grid layers", () => {
    const scene = createScene();
    const surfaces = indexById(scene.root.surfaces);

    expect([...surfaces.keys()]).toEqual([
      "body",
      "column-header",
      "row-header",
      "corner",
      "menu-overlay"
    ]);

    expect(surfaces.get("body")?.rootRect).toEqual({ x: 40, y: 60, width: 180, height: 90 });
    expect(surfaces.get("column-header")?.rootRect).toEqual({ x: 40, y: 0, width: 180, height: 60 });
    expect(surfaces.get("row-header")?.rootRect).toEqual({ x: 0, y: 60, width: 40, height: 90 });
    expect(surfaces.get("corner")?.rootRect).toEqual({ x: 0, y: 0, width: 40, height: 60 });
    expect(surfaces.get("menu-overlay")?.rootRect).toEqual({ x: 0, y: 0, width: 220, height: 150 });

    const body = surfaces.get("body");
    expect(body).toBeDefined();
    const layers = indexById(body?.layers ?? []);
    expect([...layers.keys()]).toEqual(expect.arrayContaining([
      "grid",
      "virtual-cell",
      "selection",
      "drawing-object",
      "dom-cell",
      "editor"
    ]));
    expect(indexById(surfaces.get("menu-overlay")?.layers ?? []).has("menu")).toBe(true);
    const headerLayer = indexById(surfaces.get("column-header")?.layers ?? []).get("column-header");
    expect(headerLayer?.nodes).toMatchObject([
      { kind: "column-header-band", id: "leaf", role: "leaf", level: null, rect: { y: 0, height: 60 } }
    ]);
    expect([...surfaces.values()].flatMap((surface) => surface.layers)
      .filter((layer) => layer.id !== "column-header")
      .every((layer) => layer.nodes.length === 0)).toBe(true);
  });

  it("publishes canonical header, corner, and menu panes in surface-local coordinates", () => {
    const surfaces = indexById(createScene().root.surfaces);
    const columnPanes = indexById(surfaces.get("column-header")?.panes ?? []);
    const rowPanes = indexById(surfaces.get("row-header")?.panes ?? []);

    expect(columnPanes.get("frozen")).toMatchObject({
      surface: "column-header",
      clipRect: { x: 0, y: 0, width: 50, height: 60 },
      columns: { start: 0, count: 1 },
      scroll: { x: "fixed", y: "fixed" },
      translation: { x: 0, y: 0 }
    });
    expect(columnPanes.get("scroll")).toMatchObject({
      surface: "column-header",
      clipRect: { x: 50, y: 0, width: 130, height: 60 },
      columns: { start: 1, count: 3 },
      scroll: { x: "scroll", y: "fixed" },
      translation: { x: -13, y: 0 }
    });
    expect(rowPanes.get("frozen")).toMatchObject({
      surface: "row-header",
      clipRect: { x: 0, y: 0, width: 40, height: 20 },
      rows: { start: 0, count: 1 },
      scroll: { x: "fixed", y: "fixed" },
      translation: { x: 0, y: 0 }
    });
    expect(rowPanes.get("scroll")).toMatchObject({
      surface: "row-header",
      clipRect: { x: 0, y: 20, width: 40, height: 70 },
      rows: { start: 1, count: 4 },
      scroll: { x: "fixed", y: "scroll" },
      translation: { x: 0, y: -7 }
    });
    expect(surfaces.get("corner")?.panes).toMatchObject([{
      id: "surface",
      surface: "corner",
      clipRect: { x: 0, y: 0, width: 40, height: 60 }
    }]);
    expect(surfaces.get("menu-overlay")?.panes).toMatchObject([{
      id: "surface",
      surface: "menu-overlay",
      clipRect: { x: 0, y: 0, width: 220, height: 150 }
    }]);
  });

  it("describes the four body panes with one shared clip and scrolling truth", () => {
    const scene = createScene();
    const body = indexById(scene.root.surfaces).get("body");
    const panes = indexById(body?.panes ?? []);

    expect([...panes.keys()].sort()).toEqual([
      "body",
      "frozen-columns",
      "frozen-corner",
      "frozen-rows"
    ]);
    expect(panes.get("frozen-corner")).toMatchObject({
      surface: "body",
      clipRect: { x: 0, y: 0, width: 50, height: 20 },
      rows: { start: 0, count: 1 },
      columns: { start: 0, count: 1 },
      frozen: { row: true, column: true },
      scroll: { x: "fixed", y: "fixed" },
      translation: { x: 0, y: 0 }
    });
    expect(panes.get("frozen-rows")).toMatchObject({
      surface: "body",
      clipRect: { x: 50, y: 0, width: 130, height: 20 },
      rows: { start: 0, count: 1 },
      columns: { start: 1, count: 3 },
      frozen: { row: true, column: false },
      scroll: { x: "scroll", y: "fixed" },
      translation: { x: -13, y: 0 }
    });
    expect(panes.get("frozen-columns")).toMatchObject({
      surface: "body",
      clipRect: { x: 0, y: 20, width: 50, height: 70 },
      rows: { start: 1, count: 4 },
      columns: { start: 0, count: 1 },
      frozen: { row: false, column: true },
      scroll: { x: "fixed", y: "scroll" },
      translation: { x: 0, y: -7 }
    });
    expect(panes.get("body")).toMatchObject({
      surface: "body",
      clipRect: { x: 50, y: 20, width: 130, height: 70 },
      rows: { start: 1, count: 4 },
      columns: { start: 1, count: 3 },
      frozen: { row: false, column: false },
      scroll: { x: "scroll", y: "scroll" },
      translation: { x: -13, y: -7 }
    });
  });

  it("links visible fragments to the canonical pane descriptors", () => {
    const scene = createScene();
    const body = indexById(scene.root.surfaces).get("body");
    const panes = indexById(body?.panes ?? []);
    const wholeSheet = { startRow: 0, startColumn: 0, endRow: 9, endColumn: 4 };
    const fragmentSets = [
      scene.getVisibleFragments(),
      scene.getVisibleFragmentsForRange(wholeSheet),
      scene.getFragmentsForRange(wholeSheet)
    ] as unknown as readonly (readonly FragmentWithPane[])[];

    for (const fragments of fragmentSets) {
      expect(fragments).toHaveLength(4);
      expect(new Set(fragments.map((fragment) => fragment.pane))).toEqual(new Set(panes.values()));
      for (const fragment of fragments) {
        expect(fragment.pane).toBe(panes.get(fragment.pane.id));
        expect(fragment.surface).toBe(fragment.pane.surface);
        expect(fragment.clipRect).toEqual(fragment.pane.clipRect);
        expect(fragment.frozenRow).toBe(fragment.pane.frozen.row);
        expect(fragment.frozenColumn).toBe(fragment.pane.frozen.column);
      }
    }

    for (const fragment of fragmentSets[0] ?? []) {
      expect(fragment.rows).toEqual(fragment.pane.rows);
      expect(fragment.columns).toEqual(fragment.pane.columns);
    }
  });

  it("creates canonical virtual header fragments from structural band nodes", () => {
    const scene = createScene() as unknown as SceneSnapshot;
    const columnPanes = new Map(scene.getPanes("column-header").map((pane) => [pane.id, pane]));
    const rowPanes = new Map(scene.getPanes("row-header").map((pane) => [pane.id, pane]));
    const leafBand = scene.getColumnHeaderLeafBand();
    const columnFragments = scene.getColumnHeaderBandFragments(leafBand, 0, 4);
    const rowFragments = scene.getRowHeaderFragments(0, 9);

    expect(columnFragments).toMatchObject([
      {
        surface: "column-header",
        columns: { start: 0, count: 1 },
        clipRect: { x: 0, y: 0, width: 50, height: 60 },
        rect: { x: 0, y: 0, width: 50, height: 60 }
      },
      {
        surface: "column-header",
        columns: { start: 1, count: 3 },
        clipRect: { x: 50, y: 0, width: 130, height: 60 },
        rect: { x: 37, y: 0, width: 150, height: 60 }
      }
    ]);
    expect(columnFragments[0]?.pane).toBe(columnPanes.get("frozen"));
    expect(columnFragments[1]?.pane).toBe(columnPanes.get("scroll"));
    expect(columnFragments[1]?.clipRect).toBe(columnPanes.get("scroll")?.clipRect);
    expect(columnFragments.every((fragment) => fragment.band === leafBand)).toBe(true);

    expect(rowFragments).toMatchObject([
      {
        surface: "row-header",
        rows: { start: 0, count: 1 },
        clipRect: { x: 0, y: 0, width: 40, height: 20 },
        rect: { x: 0, y: 0, width: 40, height: 20 }
      },
      {
        surface: "row-header",
        rows: { start: 1, count: 4 },
        clipRect: { x: 0, y: 20, width: 40, height: 70 },
        rect: { x: 0, y: 13, width: 40, height: 80 }
      }
    ]);
    expect(rowFragments[0]?.pane).toBe(rowPanes.get("frozen"));
    expect(rowFragments[1]?.pane).toBe(rowPanes.get("scroll"));
    expect(rowFragments[1]?.clipRect).toBe(rowPanes.get("scroll")?.clipRect);
    expect(Object.isFrozen(columnFragments)).toBe(true);
    expect(Object.isFrozen(columnFragments[0])).toBe(true);
    expect(Object.isFrozen(rowFragments)).toBe(true);
    expect(Object.isFrozen(rowFragments[0])).toBe(true);
  });

  it("publishes fill-handle geometry with the canonical pane clip", () => {
    const layout = new LayoutEngine({
      rowCount: 10,
      columnCount: 5,
      rowHeight: 20,
      columnWidths: [50, 50, 50, 50, 50],
      freeze: { rows: 1, columns: 1 }
    });
    const scene = new SceneSnapshot({
      version: 1,
      skeleton: SkeletonSnapshot.fromLayout(layout, 1),
      viewport: new ViewportSnapshot({
        version: 1,
        scrollTop: 7,
        scrollLeft: 21,
        width: 180,
        height: 90
      })
    });

    const handle = scene.getSelectionHandleFragment(
      { startRow: 3, startColumn: 3, endRow: 3, endColumn: 3 },
      6
    );
    const bodyPane = scene.getPanes().find((pane) => pane.id === "body");

    expect(handle).toMatchObject({
      surface: "body",
      rect: { x: 176, y: 70, width: 6, height: 6 },
      clipRect: { x: 50, y: 20, width: 130, height: 70 }
    });
    expect(handle?.pane).toBe(bodyPane);
    expect(handle?.clipRect).toBe(bodyPane?.clipRect);
    expect(Object.isFrozen(handle)).toBe(true);
    expect(Object.isFrozen(handle?.rect)).toBe(true);
  });

  it("omits fill handles fully hidden behind a frozen pane or outside the viewport", () => {
    const layout = new LayoutEngine({
      rowCount: 10,
      columnCount: 5,
      rowHeight: 20,
      columnWidths: [50, 50, 50, 50, 50],
      freeze: { columns: 1 }
    });
    const hiddenByFreeze = new SceneSnapshot({
      version: 1,
      skeleton: SkeletonSnapshot.fromLayout(layout, 1),
      viewport: new ViewportSnapshot({
        version: 1,
        scrollTop: 0,
        scrollLeft: 54,
        width: 180,
        height: 90
      })
    });
    const outsideViewport = new SceneSnapshot({
      version: 2,
      skeleton: SkeletonSnapshot.fromLayout(layout, 1),
      viewport: new ViewportSnapshot({
        version: 2,
        scrollTop: 0,
        scrollLeft: 0,
        width: 100,
        height: 90
      })
    });

    expect(hiddenByFreeze.getSelectionHandleFragment(
      { startRow: 0, startColumn: 1, endRow: 0, endColumn: 1 },
      6
    )).toBeNull();
    expect(outsideViewport.getSelectionHandleFragment(
      { startRow: 0, startColumn: 4, endRow: 0, endColumn: 4 },
      6
    )).toBeNull();
  });
});
