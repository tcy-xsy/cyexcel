import { describe, expect, it } from "vitest";

import { LayoutEngine, PointerController, SelectionModel } from "../../../../src/core";
import type { SelectionSnapshot, ViewportState } from "../../../../src/core";

function createLayout(overrides: Partial<ConstructorParameters<typeof LayoutEngine>[0]> = {}) {
  return new LayoutEngine({
    rowCount: 20,
    columnCount: 10,
    rowHeight: 20,
    columnWidths: Array.from({ length: 10 }, () => 50),
    ...overrides
  });
}

const VIEWPORT: ViewportState = { scrollTop: 0, scrollLeft: 0, width: 500, height: 400 };

describe("PointerController content clicks", () => {
  it("selects a single cell on a plain click in the content area", () => {
    const layout = createLayout();
    const selection = new SelectionModel({ rowCount: 20, columnCount: 10 });
    const controller = new PointerController({ layout, selection });
    controller.setViewport(VIEWPORT);

    controller.handlePointerDown({ x: 55, y: 25 }); // column 1 (50-100), row 1 (20-40)

    expect(selection.snapshot.kind).toBe("cell");
    expect(selection.snapshot.active).toEqual({ row: 1, column: 1 });
  });
});

describe("PointerController header clicks", () => {
  const headerLayout = { rowHeaderWidth: 40, columnHeaderHeight: 20 };

  it("selects the whole row when clicking the row header", () => {
    const layout = createLayout();
    const selection = new SelectionModel({ rowCount: 20, columnCount: 10 });
    const controller = new PointerController({ layout, selection, headerLayout });
    controller.setViewport(VIEWPORT);

    // y=60 lands on row (60-20)/20 = 2 once the header offset is removed.
    controller.handlePointerDown({ x: 10, y: 60 });

    expect(selection.snapshot.kind).toBe("row");
    expect(selection.snapshot.range).toEqual({
      startRow: 2,
      startColumn: 0,
      endRow: 2,
      endColumn: 9
    });
  });

  it("selects the whole column when clicking the column header", () => {
    const layout = createLayout();
    const selection = new SelectionModel({ rowCount: 20, columnCount: 10 });
    const controller = new PointerController({ layout, selection, headerLayout });
    controller.setViewport(VIEWPORT);

    // x=140 lands on column (140-40)/50 = 2 once the header offset is removed.
    controller.handlePointerDown({ x: 140, y: 10 });

    expect(selection.snapshot.kind).toBe("column");
    expect(selection.snapshot.range).toEqual({
      startRow: 0,
      startColumn: 2,
      endRow: 19,
      endColumn: 2
    });
  });

  it("selects the whole table when clicking the corner", () => {
    const layout = createLayout();
    const selection = new SelectionModel({ rowCount: 20, columnCount: 10 });
    const controller = new PointerController({ layout, selection, headerLayout });
    controller.setViewport(VIEWPORT);

    controller.handlePointerDown({ x: 10, y: 10 });

    expect(selection.snapshot.kind).toBe("all");
    expect(selection.snapshot.range).toEqual({
      startRow: 0,
      startColumn: 0,
      endRow: 19,
      endColumn: 9
    });
  });
});

describe("PointerController shift+click", () => {
  it("extends the existing selection instead of starting a new one", () => {
    const layout = createLayout();
    const selection = new SelectionModel({ rowCount: 20, columnCount: 10 });
    const controller = new PointerController({ layout, selection });
    controller.setViewport(VIEWPORT);

    controller.handlePointerDown({ x: 10, y: 10 }); // selects (0, 0)
    controller.handlePointerUp();
    controller.handlePointerDown({ x: 260, y: 110 }, { shift: true }); // extend to (5, 5)

    expect(selection.snapshot.anchor).toEqual({ row: 0, column: 0 });
    expect(selection.snapshot.active).toEqual({ row: 5, column: 5 });
    expect(selection.snapshot.range).toEqual({
      startRow: 0,
      startColumn: 0,
      endRow: 5,
      endColumn: 5
    });
  });
});

describe("PointerController drag", () => {
  it("produces the same normalized range for a forward and a reverse drag", () => {
    const layout = createLayout();

    const forwardSelection = new SelectionModel({ rowCount: 20, columnCount: 10 });
    const forward = new PointerController({ layout, selection: forwardSelection });
    forward.setViewport(VIEWPORT);
    forward.handlePointerDown({ x: 10, y: 10 }); // row 0, column 0
    forward.handlePointerMove({ x: 260, y: 110 }); // row 5, column 5
    forward.handlePointerUp();

    const backwardSelection = new SelectionModel({ rowCount: 20, columnCount: 10 });
    const backward = new PointerController({ layout, selection: backwardSelection });
    backward.setViewport(VIEWPORT);
    backward.handlePointerDown({ x: 260, y: 110 });
    backward.handlePointerMove({ x: 10, y: 10 });
    backward.handlePointerUp();

    expect(forwardSelection.snapshot.range).toEqual(backwardSelection.snapshot.range);
    expect(forwardSelection.snapshot.range).toEqual({
      startRow: 0,
      startColumn: 0,
      endRow: 5,
      endColumn: 5
    });
  });

  it("stops extending the selection after pointer up", () => {
    const layout = createLayout();
    const selection = new SelectionModel({ rowCount: 20, columnCount: 10 });
    const controller = new PointerController({ layout, selection });
    controller.setViewport(VIEWPORT);

    controller.handlePointerDown({ x: 10, y: 10 });
    controller.handlePointerUp();
    controller.handlePointerMove({ x: 260, y: 110 });

    expect(selection.snapshot.range).toEqual({
      startRow: 0,
      startColumn: 0,
      endRow: 0,
      endColumn: 0
    });
  });

  it("does not extend the selection for pointer movement within the click jitter threshold, even across a cell boundary", () => {
    const layout = createLayout(); // rowHeight 20, so y=19 is row 0, y=21 is row 1
    const selection = new SelectionModel({ rowCount: 20, columnCount: 10 });
    const controller = new PointerController({ layout, selection });
    controller.setViewport(VIEWPORT);

    controller.handlePointerDown({ x: 10, y: 19 });
    controller.handlePointerMove({ x: 10, y: 21 }); // 2px of real-world click jitter

    expect(selection.snapshot.range).toEqual({
      startRow: 0,
      startColumn: 0,
      endRow: 0,
      endColumn: 0
    });
  });

  it("extends the selection once movement exceeds the jitter threshold, and keeps extending for the rest of the drag", () => {
    const layout = createLayout();
    const selection = new SelectionModel({ rowCount: 20, columnCount: 10 });
    const controller = new PointerController({ layout, selection });
    controller.setViewport(VIEWPORT);

    controller.handlePointerDown({ x: 10, y: 10 }); // row 0
    controller.handlePointerMove({ x: 10, y: 50 }); // well past the threshold -> row 2
    expect(selection.snapshot.active).toEqual({ row: 2, column: 0 });

    controller.handlePointerMove({ x: 10, y: 15 }); // small move back toward the start point
    // Once real dragging has begun, every subsequent move keeps extending,
    // even a small one — the threshold only guards the *first* move off
    // the down point, not the whole gesture.
    expect(selection.snapshot.active).toEqual({ row: 0, column: 0 });
  });
});

describe("PointerController with frozen rows/columns and scroll", () => {
  it("resolves clicks in the frozen band without applying scroll, and scrollable clicks with it", () => {
    const layout = createLayout({ freeze: { rows: 2, columns: 1 } });
    const selection = new SelectionModel({ rowCount: 20, columnCount: 10 });
    const controller = new PointerController({ layout, selection });
    const viewport: ViewportState = { scrollTop: 100, scrollLeft: 100, width: 500, height: 400 };
    controller.setViewport(viewport);

    // Frozen band: x < 50 (1 frozen column * 50px), y < 40 (2 frozen rows * 20px).
    controller.handlePointerDown({ x: 10, y: 10 });
    expect(selection.snapshot.active).toEqual({ row: 0, column: 0 });

    // Scrollable area: content position = scroll + point.
    controller.handlePointerDown({ x: 60, y: 50 });
    // column: (100 scrollLeft + 60) / 50 = 3.2 -> 3; row: (100 scrollTop + 50) / 20 = 7.5 -> 7
    expect(selection.snapshot.active).toEqual({ row: 7, column: 3 });
  });
});

describe("PointerController onSelectionChange", () => {
  it("notifies on both pointer down and pointer move", () => {
    const layout = createLayout();
    const selection = new SelectionModel({ rowCount: 20, columnCount: 10 });
    const snapshots: SelectionSnapshot[] = [];
    const controller = new PointerController({
      layout,
      selection,
      onSelectionChange: (snapshot) => snapshots.push(snapshot)
    });
    controller.setViewport(VIEWPORT);

    controller.handlePointerDown({ x: 10, y: 10 });
    controller.handlePointerMove({ x: 60, y: 30 });

    expect(snapshots).toHaveLength(2);
  });

  it("does not notify on pointer move when not dragging", () => {
    const layout = createLayout();
    const selection = new SelectionModel({ rowCount: 20, columnCount: 10 });
    const snapshots: SelectionSnapshot[] = [];
    const controller = new PointerController({
      layout,
      selection,
      onSelectionChange: (snapshot) => snapshots.push(snapshot)
    });
    controller.setViewport(VIEWPORT);

    controller.handlePointerMove({ x: 60, y: 30 });

    expect(snapshots).toHaveLength(0);
  });
});
