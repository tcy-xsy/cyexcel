import { describe, expect, it, vi } from "vitest";

import { ColumnAxisSnapshot, SkeletonSnapshot, ViewportSnapshot } from "../../../../src/core/layout";
import { SceneSnapshot } from "../../../../src/core/scene";

function createScene(): SceneSnapshot {
  const skeleton = new SkeletonSnapshot({
    modelVersion: 1,
    layoutVersion: 1,
    rowHeights: new Float64Array(10).fill(20),
    columnAxis: new ColumnAxisSnapshot({
      columnWidths: [50, 50, 50, 50, 50],
      frozenColumnCount: 1
    }),
    frozenRowCount: 1
  });
  return new SceneSnapshot({
    version: 1,
    skeleton,
    viewport: new ViewportSnapshot({
      version: 1,
      scrollTop: 7,
      scrollLeft: 13,
      width: 180,
      height: 90
    }),
    surfaceMetrics: { columnHeaderHeight: 30, rowHeaderWidth: 40 }
  });
}

describe("SceneSnapshot complete surface hit geometry", () => {
  it("returns one immutable body hit with canonical pane, full rect, clip, and cell-local point", () => {
    const scene = createScene();

    const hit = scene.hitTestCellGeometry({ x: 50, y: 20 });

    expect(hit).toMatchObject({
      surface: "body",
      cell: { row: 1, column: 1 },
      pane: { id: "body", surface: "body" },
      rect: { x: 37, y: 13, width: 50, height: 20 },
      clipRect: { x: 50, y: 20, width: 130, height: 70 },
      localPoint: { x: 13, y: 7 }
    });
    expect(Object.isFrozen(hit)).toBe(true);
    expect(Object.isFrozen(hit?.cell)).toBe(true);
    expect(Object.isFrozen(hit?.rect)).toBe(true);
    expect(Object.isFrozen(hit?.localPoint)).toBe(true);
  });

  it("keeps frozen ownership at body pane splits and uses half-open viewport bounds", () => {
    const scene = createScene();

    expect(scene.hitTestCellGeometry({ x: 49, y: 19 })?.pane.id).toBe("frozen-corner");
    expect(scene.hitTestCellGeometry({ x: 50, y: 19 })?.pane.id).toBe("frozen-rows");
    expect(scene.hitTestCellGeometry({ x: 49, y: 20 })?.pane.id).toBe("frozen-columns");
    expect(scene.hitTestCellGeometry({ x: 50, y: 20 })?.pane.id).toBe("body");
    expect(scene.hitTestCellGeometry({ x: 180, y: 89 })).toBeNull();
    expect(scene.hitTestCellGeometry({ x: 179, y: 90 })).toBeNull();
  });

  it("resolves column-header hits through canonical header panes during partial scroll", () => {
    const scene = createScene();
    const bodyRect = vi.spyOn(SceneSnapshot.prototype, "getCellRectInPane");

    const frozen = scene.hitTestColumnHeader({ x: 49, y: 29 });
    const scroll = scene.hitTestColumnHeader({ x: 50, y: 0 });

    expect(frozen).toMatchObject({
      surface: "column-header",
      band: { id: "leaf", role: "leaf", level: null },
      column: 0,
      pane: { id: "frozen", surface: "column-header" },
      rect: { x: 0, y: 0, width: 50, height: 30 },
      clipRect: { x: 0, y: 0, width: 50, height: 30 },
      localPoint: { x: 49, y: 29 }
    });
    expect(scroll).toMatchObject({
      surface: "column-header",
      band: { id: "leaf", role: "leaf", level: null },
      column: 1,
      pane: { id: "scroll", surface: "column-header" },
      rect: { x: 37, y: 0, width: 50, height: 30 },
      clipRect: { x: 50, y: 0, width: 130, height: 30 },
      localPoint: { x: 13, y: 0 }
    });
    expect(scroll?.band).toBe(scene.getColumnHeaderLeafBand());
    expect(scene.resolveColumnHeader(49)).toBe(0);
    expect(scene.resolveColumnHeader(50)).toBe(1);
    expect(scene.hitTestColumnHeader({ x: 180, y: 0 })).toBeNull();
    expect(scene.hitTestColumnHeader({ x: 179, y: 30 })).toBeNull();
    expect(bodyRect).not.toHaveBeenCalled();
    bodyRect.mockRestore();
  });

  it("returns typed group and leaf hits at half-open band boundaries", () => {
    const base = createScene();
    const scene = new SceneSnapshot({
      version: 2,
      skeleton: new SkeletonSnapshot({
        modelVersion: base.modelVersion,
        layoutVersion: base.layoutVersion,
        rowHeights: new Float64Array(10).fill(20),
        columnAxis: new ColumnAxisSnapshot({
          columnWidths: [50, 50, 50, 50, 50],
          frozenColumnCount: 1
        }),
        frozenRowCount: 1
      }),
      viewport: new ViewportSnapshot({
        version: 2,
        scrollTop: 7,
        scrollLeft: 13,
        width: 180,
        height: 90
      }),
      surfaceMetrics: {
        columnHeaderHeight: 50,
        columnHeaderBands: [
          { role: "group", level: 0, y: 0, height: 20 },
          { role: "leaf", y: 20, height: 30 }
        ],
        rowHeaderWidth: 40
      }
    });

    expect(scene.hitTestColumnHeader({ x: 25, y: 0 })).toMatchObject({
      band: { id: "group:0", role: "group", level: 0 },
      column: 0,
      rect: { x: 0, y: 0, width: 50, height: 20 },
      localPoint: { x: 25, y: 0 }
    });
    expect(scene.hitTestColumnHeader({ x: 25, y: 19 })).toMatchObject({
      band: { id: "group:0", role: "group", level: 0 },
      column: 0,
      rect: { x: 0, y: 0, width: 50, height: 20 },
      localPoint: { x: 25, y: 19 }
    });
    expect(scene.hitTestColumnHeader({ x: 25, y: 20 })).toMatchObject({
      band: { id: "leaf", role: "leaf", level: null },
      column: 0,
      rect: { x: 0, y: 20, width: 50, height: 30 },
      localPoint: { x: 25, y: 0 }
    });
    expect(scene.hitTestColumnHeader({ x: 25, y: 49 })?.band.role).toBe("leaf");
    expect(scene.hitTestColumnHeader({ x: 25, y: 50 })).toBeNull();
  });

  it.each([
    { y: -1, height: 20 },
    { y: 0, height: -1 },
    { y: Number.NaN, height: 20 },
    { y: 0, height: Number.POSITIVE_INFINITY },
    { y: 21, height: 30 }
  ])("rejects an invalid column leaf band $y/$height", ({ y, height }) => {
    expect(() => new SceneSnapshot({
      version: 1,
      skeleton: new SkeletonSnapshot({
        modelVersion: 1,
        layoutVersion: 1,
        rowHeights: [20],
        columnAxis: new ColumnAxisSnapshot({ columnWidths: [50], frozenColumnCount: 0 }),
        frozenRowCount: 0
      }),
      viewport: new ViewportSnapshot({
        version: 1,
        scrollTop: 0,
        scrollLeft: 0,
        width: 50,
        height: 20
      }),
      surfaceMetrics: {
        columnHeaderHeight: 50,
        columnHeaderBands: [{ role: "leaf", y, height }],
        rowHeaderWidth: 40
      }
    })).toThrow(RangeError);
  });

  it("resolves row-header hits through canonical header panes during partial scroll", () => {
    const scene = createScene();
    const bodyRect = vi.spyOn(SceneSnapshot.prototype, "getCellRectInPane");

    const frozen = scene.hitTestRowHeader({ x: 39, y: 19 });
    const scroll = scene.hitTestRowHeader({ x: 0, y: 20 });

    expect(frozen).toMatchObject({
      surface: "row-header",
      row: 0,
      pane: { id: "frozen", surface: "row-header" },
      rect: { x: 0, y: 0, width: 40, height: 20 },
      clipRect: { x: 0, y: 0, width: 40, height: 20 },
      localPoint: { x: 39, y: 19 }
    });
    expect(scroll).toMatchObject({
      surface: "row-header",
      row: 1,
      pane: { id: "scroll", surface: "row-header" },
      rect: { x: 0, y: 13, width: 40, height: 20 },
      clipRect: { x: 0, y: 20, width: 40, height: 70 },
      localPoint: { x: 0, y: 7 }
    });
    expect(scene.resolveRowHeader(19)).toBe(0);
    expect(scene.resolveRowHeader(20)).toBe(1);
    expect(scene.hitTestRowHeader({ x: 40, y: 20 })).toBeNull();
    expect(scene.hitTestRowHeader({ x: 39, y: 90 })).toBeNull();
    expect(bodyRect).not.toHaveBeenCalled();
    bodyRect.mockRestore();
  });
});
