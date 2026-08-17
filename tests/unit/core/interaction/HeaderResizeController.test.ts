import { describe, expect, it } from "vitest";

import {
  HeaderResizeController,
  hitTestHeaderResizeHandle
} from "../../../../src/core/interaction/HeaderResizeController";
import {
  LayoutEngine,
  SkeletonSnapshot,
  ViewportSnapshot,
  type ViewportState
} from "../../../../src/core/layout";
import { SceneSnapshot, surfacePoint } from "../../../../src/core/scene";

const STRUCTURE_VERSION = 13;
const VIEWPORT: ViewportState = {
  scrollTop: 0,
  scrollLeft: 0,
  width: 300,
  height: 200
};

function createLayout(
  overrides: Partial<ConstructorParameters<typeof LayoutEngine>[0]> = {}
): LayoutEngine {
  return new LayoutEngine({
    rowCount: 20,
    columnCount: 6,
    rowHeight: 20,
    columnWidths: [50, 60, 70, 80, 90, 100],
    ...overrides
  });
}

function createScene(
  layout: LayoutEngine = createLayout(),
  viewport: ViewportState = VIEWPORT
): SceneSnapshot {
  return new SceneSnapshot({
    version: 1,
    skeleton: SkeletonSnapshot.fromLayout(layout, 1),
    viewport: new ViewportSnapshot({ version: 1, ...viewport }),
    surfaceMetrics: {
      columnHeaderHeight: 60,
      columnHeaderBands: [
        { role: "group", level: 0, y: 0, height: 20 },
        { role: "leaf", y: 20, height: 40 }
      ],
      rowHeaderWidth: 40
    }
  });
}

function columnPoint(x: number, y = 40) {
  return surfacePoint("column-header", x, y);
}

function rowPoint(x: number, y: number) {
  return surfacePoint("row-header", x, y);
}

describe("hitTestHeaderResizeHandle", () => {
  it("finds a column boundary only in the leaf-header band", () => {
    const scene = createScene();

    expect(hitTestHeaderResizeHandle(columnPoint(50), scene)).toEqual({
      surface: "column-header",
      axis: "column",
      column: 0
    });
    expect(hitTestHeaderResizeHandle(columnPoint(50, 10), scene)).toBeNull();
  });

  it("uses canonical frozen and scrolled column-header geometry", () => {
    const layout = createLayout({ freeze: { columns: 1 } });
    const scene = createScene(layout, { ...VIEWPORT, scrollLeft: 60 });

    expect(hitTestHeaderResizeHandle(columnPoint(49), scene)).toMatchObject({ column: 0 });
    expect(hitTestHeaderResizeHandle(columnPoint(50), scene)).toMatchObject({ column: 1 });
  });

  it("requires both a row boundary and the trailing-edge row grip", () => {
    const scene = createScene();

    expect(hitTestHeaderResizeHandle(rowPoint(35, 40), scene)).toEqual({
      surface: "row-header",
      axis: "row",
      row: 1
    });
    expect(hitTestHeaderResizeHandle(rowPoint(10, 40), scene)).toBeNull();
    expect(hitTestHeaderResizeHandle(rowPoint(35, 50), scene)).toBeNull();
  });

  it("uses canonical frozen and scrolled row-header geometry", () => {
    const layout = createLayout({ freeze: { rows: 1 } });
    const scene = createScene(layout, { ...VIEWPORT, scrollTop: 40 });

    expect(hitTestHeaderResizeHandle(rowPoint(35, 19), scene)).toMatchObject({ row: 0 });
    expect(hitTestHeaderResizeHandle(rowPoint(35, 20), scene)).toMatchObject({ row: 2 });
  });
});

describe("HeaderResizeController", () => {
  it("does not arm when pointerdown misses a handle", () => {
    const controller = new HeaderResizeController();

    expect(controller.handlePointerDown(columnPoint(25), createScene(), STRUCTURE_VERSION)).toBe(false);
    expect(controller.dragging).toBe(false);
  });

  it("returns column previews and an immutable commit intent without mutating layout", () => {
    const layout = createLayout();
    const controller = new HeaderResizeController();
    const scene = createScene(layout);
    expect(controller.handlePointerDown(columnPoint(50), scene, STRUCTURE_VERSION)).toBe(true);

    const preview = controller.handlePointerMove(
      columnPoint(72),
      createScene(layout),
      STRUCTURE_VERSION
    );

    expect(preview).toEqual({
      axis: "column",
      column: 0,
      width: 72,
      structureVersion: STRUCTURE_VERSION
    });
    expect(Object.isFrozen(preview)).toBe(true);
    expect(layout.getColumnWidth(0)).toBe(50);
    expect(controller.previewPublished).toBe(true);

    const intent = controller.handlePointerUp(STRUCTURE_VERSION);
    expect(intent).toEqual(preview);
    expect(Object.isFrozen(intent)).toBe(true);
    expect(controller.dragging).toBe(false);
    expect(controller.preview).toBeNull();
  });

  it("starts column resize in a zero-row sheet without synthesizing a body cell", () => {
    const layout = createLayout({ rowCount: 0 });
    const scene = createScene(layout);
    const controller = new HeaderResizeController();

    expect(controller.handlePointerDown(columnPoint(50), scene, STRUCTURE_VERSION)).toBe(true);
    expect(controller.handlePointerMove(columnPoint(65), scene, STRUCTURE_VERSION)).toMatchObject({
      axis: "column",
      column: 0,
      width: 65
    });
  });

  it("returns row previews and clamps them to the default minimum", () => {
    const controller = new HeaderResizeController();
    const scene = createScene();
    controller.handlePointerDown(rowPoint(35, 20), scene, STRUCTURE_VERSION);

    const preview = controller.handlePointerMove(
      rowPoint(35, -100),
      createScene(),
      STRUCTURE_VERSION
    );

    expect(preview).toEqual({
      axis: "row",
      row: 0,
      height: 16,
      structureVersion: STRUCTURE_VERSION
    });
    expect(controller.handlePointerUp(STRUCTURE_VERSION)).toEqual(preview);
  });

  it("applies per-column limits captured at pointerdown", () => {
    const controller = new HeaderResizeController({
      getColumnLimits: (column) => column === 0
        ? { minSize: 30, maxSize: 65 }
        : undefined
    });
    const scene = createScene();
    controller.handlePointerDown(columnPoint(50), scene, STRUCTURE_VERSION);

    expect(controller.handlePointerMove(columnPoint(200), scene, STRUCTURE_VERSION)).toMatchObject({
      width: 65
    });
    expect(controller.handlePointerMove(columnPoint(-200), scene, STRUCTURE_VERSION)).toMatchObject({
      width: 30
    });
  });

  it("returns no intent for a click or a drag that ends at the starting size", () => {
    const first = new HeaderResizeController();
    const scene = createScene();
    first.handlePointerDown(columnPoint(50), scene, STRUCTURE_VERSION);
    expect(first.handlePointerUp(STRUCTURE_VERSION)).toBeNull();

    const second = new HeaderResizeController();
    second.handlePointerDown(columnPoint(50), scene, STRUCTURE_VERSION);
    second.handlePointerMove(columnPoint(70), scene, STRUCTURE_VERSION);
    second.handlePointerMove(columnPoint(50), scene, STRUCTURE_VERSION);
    expect(second.previewPublished).toBe(true);
    const restoreAuthoritativeScene = second.previewPublished;
    expect(second.handlePointerUp(STRUCTURE_VERSION)).toBeNull();
    expect(restoreAuthoritativeScene).toBe(true);
    expect(second.previewPublished).toBe(false);
  });

  it("cancels permanently when structure changes during move", () => {
    const controller = new HeaderResizeController();
    const scene = createScene();
    controller.handlePointerDown(columnPoint(50), scene, STRUCTURE_VERSION);

    expect(controller.handlePointerMove(
      columnPoint(70),
      scene,
      STRUCTURE_VERSION + 1
    )).toBeNull();
    expect(controller.dragging).toBe(false);
    expect(controller.handlePointerUp(STRUCTURE_VERSION)).toBeNull();
  });

  it("cancels instead of committing when structure changes before pointerup", () => {
    const controller = new HeaderResizeController();
    const scene = createScene();
    controller.handlePointerDown(rowPoint(35, 20), scene, STRUCTURE_VERSION);
    controller.handlePointerMove(rowPoint(35, 35), scene, STRUCTURE_VERSION);

    expect(controller.handlePointerUp(STRUCTURE_VERSION + 1)).toBeNull();
    expect(controller.dragging).toBe(false);
    expect(controller.preview).toBeNull();
  });

  it("cancels on a surface switch and supports explicit cancellation", () => {
    const scene = createScene();
    const switched = new HeaderResizeController();
    switched.handlePointerDown(columnPoint(50), scene, STRUCTURE_VERSION);
    expect(switched.handlePointerMove(
      rowPoint(35, 60),
      scene,
      STRUCTURE_VERSION
    )).toBeNull();
    expect(switched.dragging).toBe(false);

    const cancelled = new HeaderResizeController();
    cancelled.handlePointerDown(rowPoint(35, 20), scene, STRUCTURE_VERSION);
    cancelled.handlePointerMove(rowPoint(35, 40), scene, STRUCTURE_VERSION);
    cancelled.cancel();
    expect(cancelled.dragging).toBe(false);
    expect(cancelled.previewPublished).toBe(false);
    expect(cancelled.handlePointerUp(STRUCTURE_VERSION)).toBeNull();
  });

  it("rejects invalid resize limits before arming", () => {
    const controller = new HeaderResizeController({
      getColumnLimits: () => ({ minSize: 80, maxSize: 40 })
    });

    expect(() => controller.handlePointerDown(
      columnPoint(50),
      createScene(),
      STRUCTURE_VERSION
    )).toThrow(/maxSize cannot be smaller/);
    expect(controller.dragging).toBe(false);
  });
});
