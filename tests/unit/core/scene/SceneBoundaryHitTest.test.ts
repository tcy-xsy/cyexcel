import { describe, expect, it, vi } from "vitest";

import { LayoutEngine, SkeletonSnapshot, ViewportSnapshot } from "../../../../src/core/layout";
import { SceneSnapshot } from "../../../../src/core/scene";
import type { ViewportState } from "../../../../src/core/layout";

interface BoundaryHitTestContract {
  hitTestColumnBoundary(x: number, tolerance: number): number | null;
  hitTestRowBoundary(y: number, tolerance: number): number | null;
}

type BoundaryScene = SceneSnapshot & BoundaryHitTestContract;

function createScene(
  layout: LayoutEngine,
  viewport: ViewportState
): BoundaryScene {
  return new SceneSnapshot({
    version: 1,
    skeleton: SkeletonSnapshot.fromLayout(layout, 1),
    viewport: new ViewportSnapshot({ version: 1, ...viewport })
  }) as BoundaryScene;
}

function createStandardScene(
  viewport: ViewportState = { scrollTop: 0, scrollLeft: 0, width: 180, height: 90 }
): BoundaryScene {
  return createScene(new LayoutEngine({
    rowCount: 10,
    columnCount: 5,
    rowHeight: 20,
    columnWidths: [50, 50, 50, 50, 50],
    freeze: { rows: 1, columns: 1 }
  }), viewport);
}

describe("SceneSnapshot boundary hit-test contract", () => {
  it("selects the owning pane before resolving overlapping frozen and partially-scrolled column handles", () => {
    const scene = createStandardScene({
      scrollTop: 0,
      scrollLeft: 48,
      width: 180,
      height: 90
    });

    // Frozen column 0 ends at x=50. Scrolled column 1 ends at x=52. Their
    // tolerance bands overlap, so pane ownership rather than iteration order
    // must decide which logical boundary wins.
    expect(scene.hitTestColumnBoundary(49, 4)).toBe(0);
    expect(scene.hitTestColumnBoundary(50, 4)).toBe(1);
    expect(scene.hitTestColumnBoundary(51, 4)).toBe(1);
  });

  it("selects the owning pane before resolving overlapping frozen and partially-scrolled row handles", () => {
    const scene = createStandardScene({
      scrollTop: 18,
      scrollLeft: 0,
      width: 180,
      height: 90
    });

    // Frozen row 0 ends at y=20. Scrolled row 1 ends at y=22.
    expect(scene.hitTestRowBoundary(19, 4)).toBe(0);
    expect(scene.hitTestRowBoundary(20, 4)).toBe(1);
    expect(scene.hitTestRowBoundary(21, 4)).toBe(1);
  });

  it("clips tolerance to the viewport and never admits a boundary outside its pane", () => {
    const scene = createStandardScene({
      scrollTop: 7,
      scrollLeft: 13,
      width: 180,
      height: 90
    });

    // Column 3 ends at x=187 and row 4 ends at y=93. They remain outside the
    // pane even though an in-viewport pointer can be within a large tolerance.
    expect(scene.hitTestColumnBoundary(179, 10)).toBeNull();
    expect(scene.hitTestRowBoundary(89, 10)).toBeNull();

    expect(scene.hitTestColumnBoundary(-1, 10)).toBeNull();
    expect(scene.hitTestColumnBoundary(180, 10)).toBeNull();
    expect(scene.hitTestRowBoundary(-1, 10)).toBeNull();
    expect(scene.hitTestRowBoundary(90, 10)).toBeNull();
  });

  it("allows an edge boundary to be approached from inside while viewport right and bottom stay open", () => {
    const scene = createStandardScene({
      scrollTop: 10,
      scrollLeft: 20,
      width: 180,
      height: 90
    });

    // Column 3 and row 4 end exactly at the viewport edge.
    expect(scene.hitTestColumnBoundary(179, 2)).toBe(3);
    expect(scene.hitTestRowBoundary(89, 2)).toBe(4);
    expect(scene.hitTestColumnBoundary(180, 2)).toBeNull();
    expect(scene.hitTestRowBoundary(90, 2)).toBeNull();
  });

  it("skips zero-width hidden columns instead of exposing duplicate logical handles", () => {
    const middleHidden = new LayoutEngine({
      rowCount: 3,
      columnCount: 3,
      rowHeight: 20,
      columnWidths: [50, 40, 50]
    });
    middleHidden.hideColumn(1);
    const middleScene = createScene(middleHidden, {
      scrollTop: 0,
      scrollLeft: 0,
      width: 140,
      height: 60
    });

    expect(middleScene.hitTestColumnBoundary(50, 0)).toBe(0);
    expect(middleScene.hitTestColumnBoundary(51, 1)).toBe(0);

    const leadingHidden = new LayoutEngine({
      rowCount: 3,
      columnCount: 3,
      rowHeight: 20,
      columnWidths: [50, 40, 50]
    });
    leadingHidden.hideColumn(0);
    const leadingScene = createScene(leadingHidden, {
      scrollTop: 0,
      scrollLeft: 0,
      width: 90,
      height: 60
    });

    expect(leadingScene.hitTestColumnBoundary(0, 0)).toBeNull();
    expect(leadingScene.hitTestColumnBoundary(40, 0)).toBe(1);
  });

  it("derives both boundary axes without projecting a fake cell through a body helper", () => {
    const scene = createStandardScene({
      scrollTop: 7,
      scrollLeft: 13,
      width: 180,
      height: 90
    });
    const paneRect = vi.spyOn(SceneSnapshot.prototype, "getCellRectInPane");

    try {
      expect(scene.hitTestColumnBoundary(87, 0)).toBe(1);
      expect(scene.hitTestRowBoundary(33, 0)).toBe(1);
      expect(paneRect).not.toHaveBeenCalled();
    } finally {
      paneRect.mockRestore();
    }
  });

  it("keeps row-boundary work bounded to pane-visible geometry for a million-row sheet", () => {
    const rowCount = 1_000_000;
    const row = rowCount - 2;
    const scene = createScene(new LayoutEngine({
      rowCount,
      columnCount: 1,
      rowHeight: 20,
      columnWidths: [50],
      freeze: { rows: 1 }
    }), {
      scrollTop: row * 20 - 20,
      scrollLeft: 0,
      width: 50,
      height: 60
    });
    const visibleRows = scene.getPanes("body")
      .filter((pane) => pane.scroll.y === "scroll")
      .reduce((count, pane) => count + pane.rows.count, 0);
    const paneRect = vi.spyOn(SceneSnapshot.prototype, "getCellRectInPane");

    expect(visibleRows).toBeLessThan(16);
    try {
      expect(scene.hitTestRowBoundary(40, 0)).toBe(row);
      expect(paneRect).not.toHaveBeenCalled();
    } finally {
      paneRect.mockRestore();
    }
  });
});
