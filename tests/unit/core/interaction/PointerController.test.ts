import { describe, expect, it } from "vitest";

import { PointerController, SelectionModel } from "../../../../src/core/interaction";
import type { SelectionSnapshot } from "../../../../src/core/interaction";
import { LayoutEngine, SkeletonSnapshot, ViewportSnapshot } from "../../../../src/core/layout";
import type { ViewportState } from "../../../../src/core/layout";
import { SceneSnapshot } from "../../../../src/core/scene";
import type { CellRange } from "../../../../src/core/types";

const STRUCTURE_VERSION = 7;
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
  layout: LayoutEngine,
  viewport: ViewportState = VIEWPORT
): SceneSnapshot {
  return new SceneSnapshot({
    version: 1,
    skeleton: SkeletonSnapshot.fromLayout(layout, 1),
    viewport: new ViewportSnapshot({ version: 1, ...viewport })
  });
}

function createFixture(options: {
  readonly layout?: LayoutEngine;
  readonly onSelectionChange?: (snapshot: SelectionSnapshot) => void;
  readonly viewport?: ViewportState;
} = {}) {
  const layout = options.layout ?? createLayout();
  const selection = new SelectionModel({ rowCount: 20, columnCount: 10 });
  const controller = new PointerController({
    selection,
    ...(options.onSelectionChange !== undefined
      ? { onSelectionChange: options.onSelectionChange }
      : {})
  });
  return {
    controller,
    selection,
    scene: createScene(layout, options.viewport)
  };
}

describe("PointerController body clicks", () => {
  it("selects a cell on a plain click", () => {
    const { controller, selection, scene } = createFixture();

    controller.down({ x: 55, y: 25 }, scene, STRUCTURE_VERSION);

    expect(selection.snapshot.kind).toBe("cell");
    expect(selection.snapshot.active).toEqual({ row: 1, column: 1 });
  });

  it("extends the active range on shift+click", () => {
    const { controller, selection, scene } = createFixture();
    controller.down({ x: 10, y: 10 }, scene, STRUCTURE_VERSION);
    controller.up(STRUCTURE_VERSION);

    controller.down(
      { x: 260, y: 110 },
      scene,
      STRUCTURE_VERSION,
      { shift: true }
    );

    expect(selection.snapshot.anchor).toEqual({ row: 0, column: 0 });
    expect(selection.snapshot.active).toEqual({ row: 5, column: 5 });
    expect(selection.snapshot.range).toEqual({
      startRow: 0,
      startColumn: 0,
      endRow: 5,
      endColumn: 5
    });
  });

  it("adds a disconnected cell range on ctrl/cmd+click", () => {
    const { controller, selection, scene } = createFixture();
    controller.down({ x: 55, y: 25 }, scene, STRUCTURE_VERSION);
    controller.up(STRUCTURE_VERSION);

    controller.down(
      { x: 155, y: 65 },
      scene,
      STRUCTURE_VERSION,
      { shift: false, ctrlOrMeta: true }
    );

    expect(selection.snapshot.ranges).toEqual([
      { startRow: 1, startColumn: 1, endRow: 1, endColumn: 1 },
      { startRow: 3, startColumn: 3, endRow: 3, endColumn: 3 }
    ]);
  });

  it("a plain click after multi-select replaces all ranges", () => {
    const { controller, selection, scene } = createFixture();
    controller.down(
      { x: 55, y: 25 },
      scene,
      STRUCTURE_VERSION,
      { shift: false, ctrlOrMeta: true }
    );
    controller.up(STRUCTURE_VERSION);

    controller.down({ x: 155, y: 65 }, scene, STRUCTURE_VERSION);

    expect(selection.snapshot.ranges).toHaveLength(1);
    expect(selection.snapshot.active).toEqual({ row: 3, column: 3 });
  });
});

describe("PointerController right-click", () => {
  it("preserves a multi-cell selection when the hit is inside it", () => {
    const { controller, selection, scene } = createFixture();
    controller.down(
      { x: 55, y: 25 },
      scene,
      STRUCTURE_VERSION,
      { shift: false, button: 0 }
    );
    controller.move({ x: 155, y: 65 }, scene, STRUCTURE_VERSION);
    controller.up(STRUCTURE_VERSION);

    controller.down(
      { x: 120, y: 50 },
      scene,
      STRUCTURE_VERSION,
      { shift: false, button: 2 }
    );

    expect(selection.snapshot.range).toEqual({
      startRow: 1,
      startColumn: 1,
      endRow: 3,
      endColumn: 3
    });
  });

  it("replaces selection outside the active range and ignores ctrl/cmd", () => {
    const { controller, selection, scene } = createFixture();
    controller.down({ x: 55, y: 25 }, scene, STRUCTURE_VERSION);
    controller.move({ x: 155, y: 65 }, scene, STRUCTURE_VERSION);
    controller.up(STRUCTURE_VERSION);

    controller.down(
      { x: 10, y: 5 },
      scene,
      STRUCTURE_VERSION,
      { shift: false, button: 2, ctrlOrMeta: true }
    );

    expect(selection.snapshot.ranges).toHaveLength(1);
    expect(selection.snapshot.range).toEqual({
      startRow: 0,
      startColumn: 0,
      endRow: 0,
      endColumn: 0
    });
  });
});

describe("PointerController drag state", () => {
  it("normalizes forward and reverse drags to the same range", () => {
    const forward = createFixture();
    forward.controller.down(
      { x: 10, y: 10 },
      forward.scene,
      STRUCTURE_VERSION
    );
    forward.controller.move(
      { x: 260, y: 110 },
      forward.scene,
      STRUCTURE_VERSION
    );

    const backward = createFixture();
    backward.controller.down(
      { x: 260, y: 110 },
      backward.scene,
      STRUCTURE_VERSION
    );
    backward.controller.move(
      { x: 10, y: 10 },
      backward.scene,
      STRUCTURE_VERSION
    );

    expect(forward.selection.snapshot.range).toEqual(
      backward.selection.snapshot.range
    );
    expect(forward.selection.snapshot.range).toEqual({
      startRow: 0,
      startColumn: 0,
      endRow: 5,
      endColumn: 5
    });
  });

  it("ignores click jitter until movement reaches the drag threshold", () => {
    const { controller, selection, scene } = createFixture();
    controller.down({ x: 10, y: 19 }, scene, STRUCTURE_VERSION);

    controller.move({ x: 10, y: 21 }, scene, STRUCTURE_VERSION);
    expect(selection.snapshot.active).toEqual({ row: 0, column: 0 });

    controller.move({ x: 10, y: 50 }, scene, STRUCTURE_VERSION);
    expect(selection.snapshot.active).toEqual({ row: 2, column: 0 });

    controller.move({ x: 10, y: 15 }, scene, STRUCTURE_VERSION);
    expect(selection.snapshot.active).toEqual({ row: 0, column: 0 });
  });

  it("stops extending after up or cancel", () => {
    const afterUp = createFixture();
    afterUp.controller.down(
      { x: 10, y: 10 },
      afterUp.scene,
      STRUCTURE_VERSION
    );
    afterUp.controller.up(STRUCTURE_VERSION);
    afterUp.controller.move(
      { x: 260, y: 110 },
      afterUp.scene,
      STRUCTURE_VERSION
    );
    expect(afterUp.selection.snapshot.active).toEqual({ row: 0, column: 0 });

    const afterCancel = createFixture();
    afterCancel.controller.down(
      { x: 10, y: 10 },
      afterCancel.scene,
      STRUCTURE_VERSION
    );
    afterCancel.controller.cancel();
    afterCancel.controller.move(
      { x: 260, y: 110 },
      afterCancel.scene,
      STRUCTURE_VERSION
    );
    expect(afterCancel.selection.snapshot.active).toEqual({ row: 0, column: 0 });
  });
});

describe("PointerController Scene resolution", () => {
  it("preserves frozen and scrollable body semantics", () => {
    const layout = createLayout({ freeze: { rows: 2, columns: 1 } });
    const fixture = createFixture({
      layout,
      viewport: {
        scrollTop: 100,
        scrollLeft: 100,
        width: 500,
        height: 400
      }
    });

    fixture.controller.down(
      { x: 10, y: 10 },
      fixture.scene,
      STRUCTURE_VERSION
    );
    expect(fixture.selection.snapshot.active).toEqual({ row: 0, column: 0 });
    fixture.controller.up(STRUCTURE_VERSION);

    fixture.controller.down(
      { x: 60, y: 50 },
      fixture.scene,
      STRUCTURE_VERSION
    );
    expect(fixture.selection.snapshot.active).toEqual({ row: 7, column: 3 });
  });

  it("uses the current Scene on move when viewport state changes", () => {
    const layout = createLayout();
    const { controller, selection, scene } = createFixture({ layout });
    const scrolled = createScene(layout, {
      scrollTop: 100,
      scrollLeft: 100,
      width: 500,
      height: 400
    });
    controller.down({ x: 10, y: 10 }, scene, STRUCTURE_VERSION);

    controller.move({ x: 60, y: 50 }, scrolled, STRUCTURE_VERSION);

    expect(selection.snapshot.active).toEqual({ row: 7, column: 3 });
  });
});

describe("PointerController structure generation", () => {
  it("cancels a drag before resolving a move from another structure", () => {
    const notifications: SelectionSnapshot[] = [];
    const { controller, selection, scene } = createFixture({
      onSelectionChange: (snapshot) => notifications.push(snapshot)
    });
    controller.down({ x: 10, y: 10 }, scene, STRUCTURE_VERSION);

    controller.move({ x: 260, y: 110 }, scene, STRUCTURE_VERSION + 1);
    controller.move({ x: 260, y: 110 }, scene, STRUCTURE_VERSION);

    expect(selection.snapshot.active).toEqual({ row: 0, column: 0 });
    expect(notifications).toHaveLength(1);
  });

  it("a mismatched up cancels the gesture permanently", () => {
    const { controller, selection, scene } = createFixture();
    controller.down({ x: 10, y: 10 }, scene, STRUCTURE_VERSION);

    controller.up(STRUCTURE_VERSION + 1);
    controller.move({ x: 260, y: 110 }, scene, STRUCTURE_VERSION);

    expect(selection.snapshot.active).toEqual({ row: 0, column: 0 });
  });
});

describe("PointerController merge-anchor normalization", () => {
  // A 3-row-by-3-column merge whose anchor is not the origin.
  const MERGE: CellRange = { startRow: 3, endRow: 5, startColumn: 2, endColumn: 4 };
  const findMergeAt = (row: number, column: number): CellRange | undefined =>
    row >= MERGE.startRow && row <= MERGE.endRow && column >= MERGE.startColumn && column <= MERGE.endColumn
      ? MERGE
      : undefined;

  function createMergeAwareFixture() {
    const layout = createLayout();
    const selection = new SelectionModel({ rowCount: 20, columnCount: 10 });
    const controller = new PointerController({ selection, findMergeAt });
    return { controller, selection, scene: createScene(layout) };
  }

  it("clicking a non-anchor cell of the merge selects the merge's anchor", () => {
    const { controller, selection, scene } = createMergeAwareFixture();

    // row 4, column 3 (x=170,y=90) is inside the merge but not its anchor.
    controller.down({ x: 170, y: 90 }, scene, STRUCTURE_VERSION);

    expect(selection.snapshot.active).toEqual({ row: 3, column: 2 });
    expect(selection.snapshot.range).toEqual({
      startRow: 3,
      startColumn: 2,
      endRow: 3,
      endColumn: 2
    });
  });

  it("dragging into a merge extends the range to the merge's anchor, not the physical cell", () => {
    const { controller, selection, scene } = createMergeAwareFixture();

    controller.down({ x: 10, y: 10 }, scene, STRUCTURE_VERSION);
    controller.move({ x: 170, y: 90 }, scene, STRUCTURE_VERSION);

    expect(selection.snapshot.range).toEqual({
      startRow: 0,
      startColumn: 0,
      endRow: 3,
      endColumn: 2
    });
  });

  it("a click outside any merge is unaffected", () => {
    const { controller, selection, scene } = createMergeAwareFixture();

    controller.down({ x: 55, y: 25 }, scene, STRUCTURE_VERSION);

    expect(selection.snapshot.active).toEqual({ row: 1, column: 1 });
  });
});

describe("PointerController notifications", () => {
  it("notifies on down and an effective drag move only", () => {
    const snapshots: SelectionSnapshot[] = [];
    const { controller, scene } = createFixture({
      onSelectionChange: (snapshot) => snapshots.push(snapshot)
    });

    controller.move({ x: 60, y: 30 }, scene, STRUCTURE_VERSION);
    controller.down({ x: 10, y: 19 }, scene, STRUCTURE_VERSION);
    controller.move({ x: 10, y: 21 }, scene, STRUCTURE_VERSION);
    controller.move({ x: 60, y: 30 }, scene, STRUCTURE_VERSION);

    expect(snapshots).toHaveLength(2);
  });
});
