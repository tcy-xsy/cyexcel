import { describe, expect, it } from "vitest";

import { HeaderManager, hitTestHeaderGroup } from "../../../../src/core/headers";
import { LayoutEngine, SkeletonSnapshot, ViewportSnapshot } from "../../../../src/core/layout";
import type { ViewportState } from "../../../../src/core/layout";
import { SceneSnapshot } from "../../../../src/core/scene";

const VIEWPORT: ViewportState = { scrollTop: 0, scrollLeft: 0, width: 200, height: 100 };
const GROUP_HEADER_HEIGHT = 20;

function createScene(layout: LayoutEngine, viewport: ViewportState = VIEWPORT): SceneSnapshot {
  return new SceneSnapshot({
    version: 1,
    skeleton: SkeletonSnapshot.fromLayout(layout, 1),
    viewport: new ViewportSnapshot({ version: 1, ...viewport }),
    surfaceMetrics: {
      columnHeaderHeight: 80,
      columnHeaderBands: [
        { role: "group", level: 2, y: 0, height: GROUP_HEADER_HEIGHT },
        { role: "group", level: 1, y: 20, height: GROUP_HEADER_HEIGHT },
        { role: "group", level: 0, y: 40, height: GROUP_HEADER_HEIGHT },
        { role: "leaf", y: 60, height: 20 }
      ],
      rowHeaderWidth: 0
    }
  });
}

function createLayout(columnCount: number, freeze?: { columns?: number }): LayoutEngine {
  return new LayoutEngine({
    rowCount: 10,
    columnCount,
    rowHeight: 20,
    columnWidths: Array.from({ length: columnCount }, () => 50),
    ...(freeze ? { freeze } : {})
  });
}

function threeLevelHeaderManager(): HeaderManager {
  // Level 0: "cat-a"[0,1]/"cat-b"[2,3]. Level 1: "dept"[0,3]. Level 2: "root"[0,3].
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

describe("hitTestHeaderGroup", () => {
  it("resolves a point in the middle level (level 1) to the correct group", () => {
    const layout = createLayout(4);
    const headerManager = threeLevelHeaderManager();

    // y=20..40 is level 1 ("dept"); x=75 is column 1 (columns are 50px wide).
    const result = hitTestHeaderGroup({ x: 75, y: 25 }, createScene(layout), headerManager);

    expect(result).toEqual({ level: 1, group: headerManager.getGroups(1)[0] });
  });

  it("resolves a point at the outermost level (level 2, y=0) to the root group", () => {
    const layout = createLayout(4);
    const headerManager = threeLevelHeaderManager();

    const result = hitTestHeaderGroup({ x: 10, y: 5 }, createScene(layout), headerManager);

    expect(result).toEqual({ level: 2, group: headerManager.getGroups(2)[0] });
  });

  it("resolves a point at the leaf-adjacent level (level 0) to the narrower category group", () => {
    const layout = createLayout(4);
    const headerManager = threeLevelHeaderManager();

    // y=40..60 is level 0; x=125 is column 2 -> "cat-b".
    const result = hitTestHeaderGroup({ x: 125, y: 45 }, createScene(layout), headerManager);

    expect(result).toEqual({ level: 0, group: headerManager.getGroups(0)[1] });
  });

  it("returns null when the point is below the group band (in the leaf row or body)", () => {
    const layout = createLayout(4);
    const headerManager = threeLevelHeaderManager();

    // Group band is y=[0,60) for a 3-level tree; y=60 is the leaf row.
    const result = hitTestHeaderGroup({ x: 10, y: 60 }, createScene(layout), headerManager);

    expect(result).toBeNull();
  });

  it("returns null for a negative y", () => {
    const layout = createLayout(4);
    const headerManager = threeLevelHeaderManager();

    const result = hitTestHeaderGroup({ x: 10, y: -5 }, createScene(layout), headerManager);

    expect(result).toBeNull();
  });

  it("returns null when there are no groups at all (groupDepth 0)", () => {
    const layout = createLayout(4);
    const headerManager = new HeaderManager({ columns: [{ id: "a" }, { id: "b" }, { id: "c" }, { id: "d" }] });

    const result = hitTestHeaderGroup({ x: 10, y: 0 }, createScene(layout), headerManager);

    expect(result).toBeNull();
  });

  it("returns null for a column whose branch doesn't reach this level (blank background)", () => {
    const layout = createLayout(4);
    const headerManager = new HeaderManager({
      columns: [{ id: "a" }, { id: "b" }, { id: "c" }, { id: "d" }],
      groups: [
        { id: "shallow", title: "Shallow", columns: [0, 1] },
        { id: "wrapper", title: "Wrapper", children: [{ id: "deep", title: "Deep", columns: [2, 3] }] }
      ]
    });
    expect(headerManager.groupDepth).toBe(2);

    // Level 1 (y=0..20) has no group over columns 0-1 ("shallow" has nothing above it).
    const result = hitTestHeaderGroup({ x: 10, y: 5 }, createScene(layout), headerManager);

    expect(result).toBeNull();
  });

  it("resolves the correct column when it's inside the frozen band, accounting for scroll on the rest", () => {
    const layout = createLayout(4, { columns: 1 });
    const headerManager = threeLevelHeaderManager();
    const scrolledViewport: ViewportState = { ...VIEWPORT, scrollLeft: 40 };

    // Frozen column 0 stays at its own absolute x regardless of scrollLeft.
    const result = hitTestHeaderGroup({ x: 10, y: 45 }, createScene(layout, scrolledViewport), headerManager);

    expect(result).toEqual({ level: 0, group: headerManager.getGroups(0)[0] }); // "cat-a" covers column 0
  });
});
