import { describe, expect, it } from "vitest";

import { LayoutEngine, SkeletonSnapshot, ViewportSnapshot } from "../../../../src/core/layout";
import type { ViewportState } from "../../../../src/core/layout";
import { SceneSnapshot } from "../../../../src/core/scene";
import type { CellRange } from "../../../../src/core/types";

const VIEWPORT: ViewportState = { scrollTop: 0, scrollLeft: 0, width: 300, height: 200 };

function createLayout(freeze?: { readonly rows?: number; readonly columns?: number }): LayoutEngine {
  return new LayoutEngine({
    rowCount: 10,
    columnCount: 5,
    rowHeight: 20,
    columnWidths: [50, 50, 50, 50, 50],
    ...(freeze ? { freeze } : {})
  });
}

function createScene(layout: LayoutEngine, viewport: ViewportState = VIEWPORT): SceneSnapshot {
  return new SceneSnapshot({
    version: 1,
    skeleton: SkeletonSnapshot.fromLayout(layout, 1),
    viewport: new ViewportSnapshot({ version: 1, ...viewport })
  });
}

function createOverscannedScene(): SceneSnapshot {
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
      scrollTop: 47,
      scrollLeft: 13,
      width: 180,
      height: 90,
      overscanRows: 2,
      overscanColumns: 1
    })
  });
}

describe("SceneSnapshot.hitTestCellGeometry", () => {
  it("resolves a point inside the body surface to a cell", () => {
    const scene = createScene(createLayout());

    expect(scene.hitTestCellGeometry({ x: 55, y: 45 })?.cell).toEqual({ row: 2, column: 1 });
  });

  it("returns null for negative points and points past the data extent", () => {
    const scene = createScene(createLayout());

    expect(scene.hitTestCellGeometry({ x: -1, y: 0 })).toBeNull();
    expect(scene.hitTestCellGeometry({ x: 0, y: -1 })).toBeNull();
    expect(scene.hitTestCellGeometry({ x: 0, y: 500 })).toBeNull();
    expect(scene.hitTestCellGeometry({ x: 500, y: 0 })).toBeNull();
  });

  it("applies scroll to the scrollable region", () => {
    const scene = createScene(createLayout(), { ...VIEWPORT, scrollTop: 40, scrollLeft: 50 });

    expect(scene.hitTestCellGeometry({ x: 10, y: 10 })?.cell).toEqual({ row: 2, column: 1 });
  });

  it("keeps frozen rows and columns pinned while scrollable cells use scroll", () => {
    const scene = createScene(createLayout({ rows: 1, columns: 1 }), {
      ...VIEWPORT,
      scrollTop: 40,
      scrollLeft: 50
    });

    expect(scene.hitTestCellGeometry({ x: 10, y: 10 })?.cell).toEqual({ row: 0, column: 0 });
    expect(scene.hitTestCellGeometry({ x: 60, y: 30 })?.cell).toEqual({ row: 3, column: 2 });
  });

  it("treats the viewport right and bottom edges as out of bounds", () => {
    const scene = createScene(createLayout({ rows: 1, columns: 1 }), {
      scrollTop: 7,
      scrollLeft: 13,
      width: 180,
      height: 90
    });

    expect(scene.hitTestCellGeometry({ x: 179, y: 89 })?.cell).toEqual({ row: 4, column: 3 });
    expect(scene.hitTestCellGeometry({ x: 180, y: 89 })).toBeNull();
    expect(scene.hitTestCellGeometry({ x: 179, y: 90 })).toBeNull();
    expect(scene.hitTestCellGeometry({ x: 180, y: 90 })).toBeNull();
  });

  it("keeps frozen hit-test ownership at the pane boundary during partial-cell scroll", () => {
    const scene = createScene(createLayout({ rows: 1, columns: 1 }), {
      scrollTop: 7,
      scrollLeft: 13,
      width: 180,
      height: 90
    });

    expect(scene.hitTestCellGeometry({ x: 49, y: 19 })?.cell).toEqual({ row: 0, column: 0 });
    expect(scene.hitTestCellGeometry({ x: 50, y: 19 })?.cell).toEqual({ row: 0, column: 1 });
    expect(scene.hitTestCellGeometry({ x: 49, y: 20 })?.cell).toEqual({ row: 1, column: 0 });
    expect(scene.hitTestCellGeometry({ x: 50, y: 20 })?.cell).toEqual({ row: 1, column: 1 });
  });
});

describe("SceneSnapshot.hitTestCellGeometry merge-anchor normalization", () => {
  // A 3-row-by-3-column merge whose anchor is not the origin — exercising
  // both "not at (0,0)" and "spans more than 2 cells" in the same fixture.
  const MERGE: CellRange = { startRow: 2, endRow: 4, startColumn: 1, endColumn: 3 };
  const findMergeAt = (row: number, column: number): CellRange | undefined =>
    row >= MERGE.startRow && row <= MERGE.endRow && column >= MERGE.startColumn && column <= MERGE.endColumn
      ? MERGE
      : undefined;

  function createMergeAwareScene(): SceneSnapshot {
    return new SceneSnapshot({
      version: 1,
      skeleton: SkeletonSnapshot.fromLayout(createLayout(), 1),
      viewport: new ViewportSnapshot({ version: 1, ...VIEWPORT }),
      findMergeAt
    });
  }

  it("normalizes a click on a non-anchor cell of the merge to the merge's anchor", () => {
    const scene = createMergeAwareScene();

    // row 3, column 2 — the merge's interior, not its anchor.
    const hit = scene.hitTestCellGeometry({ x: 120, y: 70 });

    expect(hit?.cell).toEqual({ row: 2, column: 1 });
  });

  it("normalizes a click on the merge's far corner (not just an edge cell) to the anchor", () => {
    const scene = createMergeAwareScene();

    // row 4, column 3 — the merge's bottom-right corner.
    const hit = scene.hitTestCellGeometry({ x: 170, y: 90 });

    expect(hit?.cell).toEqual({ row: 2, column: 1 });
  });

  it("keeps rect/localPoint describing the physically clicked cell, not the whole merge", () => {
    const scene = createMergeAwareScene();

    const hit = scene.hitTestCellGeometry({ x: 120, y: 70 });

    // Physical cell (row 3, column 2) rect: x=100..150, y=60..80.
    expect(hit?.rect).toEqual({ x: 100, y: 60, width: 50, height: 20 });
    expect(hit?.localPoint).toEqual({ x: 20, y: 10 });
  });

  it("leaves a click outside any merge unaffected", () => {
    const scene = createMergeAwareScene();

    expect(scene.hitTestCellGeometry({ x: 0, y: 0 })?.cell).toEqual({ row: 0, column: 0 });
    expect(scene.hitTestCellGeometry({ x: 210, y: 145 })?.cell).toEqual({ row: 7, column: 4 });
  });

  it("defaults to no merge-anchor normalization when findMergeAt is not provided", () => {
    const scene = createScene(createLayout());

    expect(scene.hitTestCellGeometry({ x: 120, y: 70 })?.cell).toEqual({ row: 3, column: 2 });
  });
});

describe("SceneSnapshot pane clips", () => {
  it("publishes overscanned row and column traversal through canonical panes and fragments", () => {
    const scene = createOverscannedScene();
    const panes = new Map(scene.getPanes().map((pane) => [pane.id, pane]));

    expect(scene.overscanRows).toBe(2);
    expect(scene.overscanColumns).toBe(1);
    expect(scene.getVisibleRows()).toEqual({ start: 1, count: 8 });
    expect(scene.getVisibleColumns()).toEqual({ start: 1, count: 4 });
    expect(panes.get("frozen-columns")?.rows).toEqual({ start: 1, count: 8 });
    expect(panes.get("frozen-rows")?.columns).toEqual({ start: 1, count: 4 });
    expect(panes.get("body")).toMatchObject({
      rows: { start: 1, count: 8 },
      columns: { start: 1, count: 4 }
    });
    expect(scene.getVisibleFragments().at(-1)).toMatchObject({
      rows: { start: 1, count: 8 },
      columns: { start: 1, count: 4 }
    });
  });

  it("keeps four fixed viewport clips when partial-cell scroll moves fragment rects under frozen bands", () => {
    const scene = createScene(createLayout({ rows: 1, columns: 1 }), {
      scrollTop: 7,
      scrollLeft: 13,
      width: 180,
      height: 90
    });

    expect(scene.getVisibleFragments()).toMatchObject([
      {
        rows: { start: 0, count: 1 },
        columns: { start: 0, count: 1 },
        frozenRow: true,
        frozenColumn: true,
        clipRect: { x: 0, y: 0, width: 50, height: 20 },
        rect: { x: 0, y: 0, width: 50, height: 20 }
      },
      {
        rows: { start: 0, count: 1 },
        columns: { start: 1, count: 3 },
        frozenRow: true,
        frozenColumn: false,
        clipRect: { x: 50, y: 0, width: 130, height: 20 },
        rect: { x: 37, y: 0, width: 150, height: 20 }
      },
      {
        rows: { start: 1, count: 4 },
        columns: { start: 0, count: 1 },
        frozenRow: false,
        frozenColumn: true,
        clipRect: { x: 0, y: 20, width: 50, height: 70 },
        rect: { x: 0, y: 13, width: 50, height: 80 }
      },
      {
        rows: { start: 1, count: 4 },
        columns: { start: 1, count: 3 },
        frozenRow: false,
        frozenColumn: false,
        clipRect: { x: 50, y: 20, width: 130, height: 70 },
        rect: { x: 37, y: 13, width: 150, height: 80 }
      }
    ]);

    const fragments = scene.getVisibleFragments();
    expect(fragments[1]?.rect.x).toBeLessThan(fragments[1]?.clipRect.x ?? 0);
    expect(fragments[2]?.rect.y).toBeLessThan(fragments[2]?.clipRect.y ?? 0);
    expect(fragments[3]?.rect.x).toBeLessThan(fragments[3]?.clipRect.x ?? 0);
    expect(fragments[3]?.rect.y).toBeLessThan(fragments[3]?.clipRect.y ?? 0);
  });
});

describe("SceneSnapshot.getScrollTargetForCell", () => {
  it("computes the nearest scroll offset from canonical pane clips", () => {
    const scene = createScene(createLayout({ rows: 1, columns: 1 }), {
      scrollTop: 40,
      scrollLeft: 50,
      width: 180,
      height: 90
    });

    expect(scene.getScrollTargetForCell({ row: 7, column: 4 })).toEqual({
      logicalTop: 70,
      logicalLeft: 70
    });
    expect(scene.getScrollTargetForCell({ row: 1, column: 1 })).toEqual({
      logicalTop: 0,
      logicalLeft: 0
    });
  });

  it("does not scroll either fixed axis for a frozen cell", () => {
    const scene = createScene(createLayout({ rows: 1, columns: 1 }), {
      scrollTop: 40,
      scrollLeft: 50,
      width: 180,
      height: 90
    });

    expect(scene.getScrollTargetForCell({ row: 0, column: 0 })).toEqual({
      logicalTop: 40,
      logicalLeft: 50
    });
    expect(scene.getScrollTargetForCell({ row: -1, column: 0 })).toBeNull();
  });
});
