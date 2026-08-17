import { describe, expect, it } from "vitest";

import {
  HeaderSelectionController,
  type HeaderSelectionPoint
} from "../../../../src/core/interaction/HeaderSelectionController";
import { SelectionModel } from "../../../../src/core/interaction/SelectionModel";
import {
  LayoutEngine,
  SkeletonSnapshot,
  ViewportSnapshot,
  type ViewportState
} from "../../../../src/core/layout";
import { SceneSnapshot, surfacePoint } from "../../../../src/core/scene";

const STRUCTURE_VERSION = 11;
const VIEWPORT: ViewportState = {
  scrollTop: 0,
  scrollLeft: 0,
  width: 500,
  height: 400
};

function createLayout(
  overrides: Partial<ConstructorParameters<typeof LayoutEngine>[0]> = {}
): LayoutEngine {
  return new LayoutEngine({
    rowCount: 20,
    columnCount: 10,
    rowHeight: 20,
    columnWidths: Array.from({ length: 10 }, () => 50),
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
    surfaceMetrics: { columnHeaderHeight: 40, rowHeaderWidth: 30 }
  });
}

function createFixture(scene = createScene()) {
  return {
    controller: new HeaderSelectionController(),
    selection: new SelectionModel({ rowCount: 20, columnCount: 10 }),
    scene
  };
}

function columnPoint(x: number, y = 20) {
  return surfacePoint("column-header", x, y);
}

function rowPoint(y: number, x = 15) {
  return surfacePoint("row-header", x, y);
}

describe("HeaderSelectionController column-header", () => {
  it("selects a whole column and returns the applied action", () => {
    const { controller, selection, scene } = createFixture();

    const result = controller.down(columnPoint(125), scene, selection, STRUCTURE_VERSION);

    expect(result?.surface).toBe("column-header");
    expect(result?.snapshot.kind).toBe("column");
    expect(selection.snapshot.range).toEqual({
      startRow: 0,
      startColumn: 2,
      endRow: 19,
      endColumn: 2
    });
    expect(controller.dragging).toBe(true);
  });

  it("keeps the existing anchor for shift selection", () => {
    const { controller, selection, scene } = createFixture();
    selection.selectColumn(2);

    controller.down(
      columnPoint(275),
      scene,
      selection,
      STRUCTURE_VERSION,
      { shift: true }
    );

    expect(selection.snapshot.kind).toBe("column");
    expect(selection.snapshot.anchor.column).toBe(2);
    expect(selection.snapshot.range).toEqual({
      startRow: 0,
      startColumn: 2,
      endRow: 19,
      endColumn: 5
    });
  });

  it("does not shrink the active range on a right-click inside it", () => {
    const { controller, selection, scene } = createFixture();
    selection.selectColumn(2);
    selection.extendTo(0, 5);
    const before = selection.snapshot.range;

    const result = controller.down(
      columnPoint(225),
      scene,
      selection,
      STRUCTURE_VERSION,
      { button: 2 }
    );

    expect(result).toBeNull();
    expect(selection.snapshot.range).toEqual(before);
    expect(controller.dragging).toBe(false);
  });

  it("selects a new column on a right-click outside the active range", () => {
    const { controller, selection, scene } = createFixture();
    selection.selectColumn(2);

    controller.down(
      columnPoint(375),
      scene,
      selection,
      STRUCTURE_VERSION,
      { button: 2 }
    );

    expect(selection.snapshot.range.startColumn).toBe(7);
    expect(selection.snapshot.range.endColumn).toBe(7);
    expect(controller.dragging).toBe(false);
  });

  it("waits for the axis drag threshold before extending", () => {
    const { controller, selection, scene } = createFixture();
    controller.down(columnPoint(75), scene, selection, STRUCTURE_VERSION);

    expect(controller.move(columnPoint(78), scene, selection, STRUCTURE_VERSION)).toBeNull();
    expect(selection.snapshot.range.endColumn).toBe(1);

    const result = controller.move(columnPoint(180), scene, selection, STRUCTURE_VERSION);
    expect(result?.surface).toBe("column-header");
    expect(selection.snapshot.range.endColumn).toBe(3);
  });

  it("uses canonical frozen and scrolled header geometry on down", () => {
    const layout = createLayout({ freeze: { columns: 1 } });
    const scene = createScene(layout, { ...VIEWPORT, scrollLeft: 100 });
    const { controller, selection } = createFixture(scene);

    controller.down(columnPoint(10), scene, selection, STRUCTURE_VERSION);
    expect(selection.snapshot.active.column).toBe(0);
    controller.up(STRUCTURE_VERSION);

    controller.down(columnPoint(60), scene, selection, STRUCTURE_VERSION);
    expect(selection.snapshot.active.column).toBe(3);
  });

  it("declines group bands and selects only inside the configured leaf band", () => {
    const layout = createLayout();
    const scene = new SceneSnapshot({
      version: 1,
      skeleton: SkeletonSnapshot.fromLayout(layout, 1),
      viewport: new ViewportSnapshot({ version: 1, ...VIEWPORT }),
      surfaceMetrics: {
        columnHeaderHeight: 60,
        columnHeaderBands: [
          { role: "group", level: 0, y: 0, height: 20 },
          { role: "leaf", y: 20, height: 40 }
        ],
        rowHeaderWidth: 30
      }
    });
    const { controller, selection } = createFixture(scene);

    expect(controller.down(
      columnPoint(125, 19),
      scene,
      selection,
      STRUCTURE_VERSION
    )).toBeNull();
    expect(controller.dragging).toBe(false);

    expect(controller.down(
      columnPoint(125, 20),
      scene,
      selection,
      STRUCTURE_VERSION
    )?.snapshot.active.column).toBe(2);
    expect(controller.dragging).toBe(true);
  });
});

describe("HeaderSelectionController row-header and corner", () => {
  it("selects and shift-extends whole rows", () => {
    const { controller, selection, scene } = createFixture();
    controller.down(rowPoint(50), scene, selection, STRUCTURE_VERSION);
    controller.up(STRUCTURE_VERSION);

    const result = controller.down(
      rowPoint(110),
      scene,
      selection,
      STRUCTURE_VERSION,
      { shift: true }
    );

    expect(result?.surface).toBe("row-header");
    expect(selection.snapshot.kind).toBe("row");
    expect(selection.snapshot.range).toEqual({
      startRow: 2,
      startColumn: 0,
      endRow: 5,
      endColumn: 9
    });
  });

  it("preserves a multi-row range on right-click and extends after vertical threshold", () => {
    const { controller, selection, scene } = createFixture();
    selection.selectRow(2);
    selection.extendTo(5, 0);

    expect(controller.down(
      rowPoint(70),
      scene,
      selection,
      STRUCTURE_VERSION,
      { button: 2 }
    )).toBeNull();
    expect(selection.snapshot.range.startRow).toBe(2);
    expect(selection.snapshot.range.endRow).toBe(5);

    controller.down(rowPoint(50), scene, selection, STRUCTURE_VERSION);
    expect(controller.move(rowPoint(53), scene, selection, STRUCTURE_VERSION)).toBeNull();
    expect(selection.snapshot.range.endRow).toBe(2);
    controller.move(rowPoint(130), scene, selection, STRUCTURE_VERSION);
    expect(selection.snapshot.range.endRow).toBe(6);
  });

  it("selects all from a point inside the corner and ignores points outside it", () => {
    const { controller, selection, scene } = createFixture();

    expect(controller.down(
      surfacePoint("corner", 30, 10),
      scene,
      selection,
      STRUCTURE_VERSION
    )).toBeNull();
    expect(selection.snapshot.kind).toBe("cell");

    const result = controller.down(
      surfacePoint("corner", 29, 39),
      scene,
      selection,
      STRUCTURE_VERSION
    );
    expect(result?.surface).toBe("corner");
    expect(selection.snapshot.kind).toBe("all");
    expect(controller.dragging).toBe(false);
  });

  it("updates a row selection on secondary click without arming a drag", () => {
    const { controller, selection, scene } = createFixture();
    selection.selectRow(2);

    const result = controller.down(
      rowPoint(150),
      scene,
      selection,
      STRUCTURE_VERSION,
      { button: 2 }
    );

    expect(result?.surface).toBe("row-header");
    expect(selection.snapshot.range).toEqual({
      startRow: 7,
      startColumn: 0,
      endRow: 7,
      endColumn: 9
    });
    expect(controller.dragging).toBe(false);
  });
});

describe("HeaderSelectionController gesture lifetime", () => {
  it("cancels permanently when structure changes before move", () => {
    const { controller, selection, scene } = createFixture();
    controller.down(columnPoint(25), scene, selection, STRUCTURE_VERSION);

    expect(controller.move(
      columnPoint(225),
      scene,
      selection,
      STRUCTURE_VERSION + 1
    )).toBeNull();
    expect(controller.dragging).toBe(false);
    expect(controller.move(
      columnPoint(225),
      scene,
      selection,
      STRUCTURE_VERSION
    )).toBeNull();
    expect(selection.snapshot.range.endColumn).toBe(0);
  });

  it("a mismatched up or explicit cancel prevents later moves", () => {
    const first = createFixture();
    first.controller.down(rowPoint(10), first.scene, first.selection, STRUCTURE_VERSION);
    first.controller.up(STRUCTURE_VERSION + 1);
    expect(first.controller.move(
      rowPoint(150),
      first.scene,
      first.selection,
      STRUCTURE_VERSION
    )).toBeNull();
    expect(first.selection.snapshot.range.endRow).toBe(0);

    const second = createFixture();
    second.controller.down(rowPoint(10), second.scene, second.selection, STRUCTURE_VERSION);
    second.controller.cancel();
    expect(second.controller.move(
      rowPoint(150),
      second.scene,
      second.selection,
      STRUCTURE_VERSION
    )).toBeNull();
    expect(second.selection.snapshot.range.endRow).toBe(0);
  });

  it("rejects unsupported buttons, non-primary pointers, misses, and surface switches", () => {
    const { controller, selection, scene } = createFixture();
    expect(controller.down(
      columnPoint(25),
      scene,
      selection,
      STRUCTURE_VERSION,
      { button: 1 }
    )).toBeNull();
    expect(controller.down(
      rowPoint(10),
      scene,
      selection,
      STRUCTURE_VERSION,
      { isPrimary: false }
    )).toBeNull();
    expect(controller.down(
      columnPoint(25, 40),
      scene,
      selection,
      STRUCTURE_VERSION
    )).toBeNull();

    controller.down(columnPoint(25), scene, selection, STRUCTURE_VERSION);
    expect(controller.move(
      rowPoint(100) as HeaderSelectionPoint,
      scene,
      selection,
      STRUCTURE_VERSION
    )).toBeNull();
    expect(controller.dragging).toBe(false);
  });
});
