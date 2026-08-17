import { describe, expect, it } from "vitest";

import { KeyboardController, SelectionModel } from "../../../../src/core/interaction";
import type { SelectionSnapshot } from "../../../../src/core/interaction";

function createController(rowCount = 20, columnCount = 10) {
  const selection = new SelectionModel({ rowCount, columnCount });
  const snapshots: SelectionSnapshot[] = [];
  const controller = new KeyboardController({
    selection,
    onSelectionChange: (snapshot) => snapshots.push(snapshot)
  });
  selection.selectCell(5, 5);
  return { selection, controller, snapshots };
}

describe("KeyboardController arrow keys", () => {
  it("moves the active cell by one in each direction", () => {
    const { selection, controller } = createController();

    controller.handleKeyDown("ArrowUp");
    expect(selection.snapshot.active).toEqual({ row: 4, column: 5 });

    controller.handleKeyDown("ArrowDown");
    expect(selection.snapshot.active).toEqual({ row: 5, column: 5 });

    controller.handleKeyDown("ArrowLeft");
    expect(selection.snapshot.active).toEqual({ row: 5, column: 4 });

    controller.handleKeyDown("ArrowRight");
    expect(selection.snapshot.active).toEqual({ row: 5, column: 5 });
  });

  it("extends the range instead of moving the anchor when Shift is held", () => {
    const { selection, controller } = createController();

    controller.handleKeyDown("ArrowDown", { shift: true });

    expect(selection.snapshot.anchor).toEqual({ row: 5, column: 5 });
    expect(selection.snapshot.active).toEqual({ row: 6, column: 5 });
    expect(selection.snapshot.range).toEqual({
      startRow: 5,
      startColumn: 5,
      endRow: 6,
      endColumn: 5
    });
  });

  it("clamps at the table edges", () => {
    const { selection, controller } = createController();
    selection.selectCell(0, 0);

    controller.handleKeyDown("ArrowUp");
    controller.handleKeyDown("ArrowLeft");

    expect(selection.snapshot.active).toEqual({ row: 0, column: 0 });
  });
});

describe("KeyboardController Tab/Enter", () => {
  it("Tab moves right and collapses even if the selection was extended", () => {
    const { selection, controller } = createController();
    selection.selectCell(5, 5);
    selection.extendTo(8, 8);

    controller.handleKeyDown("Tab");

    // extendTo(8, 8) left active at (8, 8); Tab moves one column right from
    // there and collapses the range (anchor follows active).
    expect(selection.snapshot.kind).toBe("cell");
    expect(selection.snapshot.anchor).toEqual(selection.snapshot.active);
    expect(selection.snapshot.active).toEqual({ row: 8, column: 9 });
  });

  it("Shift+Tab moves left", () => {
    const { selection, controller } = createController();

    controller.handleKeyDown("Tab", { shift: true });

    expect(selection.snapshot.active).toEqual({ row: 5, column: 4 });
  });

  it("Enter moves down, Shift+Enter moves up", () => {
    const { selection, controller } = createController();

    controller.handleKeyDown("Enter");
    expect(selection.snapshot.active).toEqual({ row: 6, column: 5 });

    controller.handleKeyDown("Enter", { shift: true });
    expect(selection.snapshot.active).toEqual({ row: 5, column: 5 });
  });
});

describe("KeyboardController Home/End", () => {
  it("Home jumps to the first column of the active row", () => {
    const { selection, controller } = createController();

    controller.handleKeyDown("Home");

    expect(selection.snapshot.active).toEqual({ row: 5, column: 0 });
  });

  it("End jumps to the last column of the active row", () => {
    const { selection, controller } = createController(20, 10);

    controller.handleKeyDown("End");

    expect(selection.snapshot.active).toEqual({ row: 5, column: 9 });
  });

  it("Shift+Home/End extend instead of collapsing", () => {
    const { selection, controller } = createController();

    controller.handleKeyDown("End", { shift: true });

    expect(selection.snapshot.anchor).toEqual({ row: 5, column: 5 });
    expect(selection.snapshot.active).toEqual({ row: 5, column: 9 });
  });
});

describe("KeyboardController unhandled keys", () => {
  it("returns false and does not change selection for an unmapped key", () => {
    const { selection, controller } = createController();
    const before = selection.snapshot;

    const handled = controller.handleKeyDown("a");

    expect(handled).toBe(false);
    expect(selection.snapshot).toEqual(before);
  });
});

describe("KeyboardController onSelectionChange", () => {
  it("notifies once per handled key", () => {
    const { controller, snapshots } = createController();

    controller.handleKeyDown("ArrowRight");
    controller.handleKeyDown("a"); // unhandled, no notification
    controller.handleKeyDown("Tab");

    expect(snapshots).toHaveLength(2);
  });
});
