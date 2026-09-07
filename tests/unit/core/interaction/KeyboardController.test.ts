import { describe, expect, it } from "vitest";

import { KeyboardController, SelectionModel } from "../../../../src/core/interaction";
import type { SelectionSnapshot } from "../../../../src/core/interaction";
import type { CellRange } from "../../../../src/core/types";

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

describe("KeyboardController arrow keys over a merge", () => {
  // A 3-row-by-3-column merge whose anchor is not the origin.
  const MERGE: CellRange = { startRow: 3, endRow: 5, startColumn: 2, endColumn: 4 };
  const findMergeAt = (row: number, column: number): CellRange | undefined =>
    row >= MERGE.startRow && row <= MERGE.endRow && column >= MERGE.startColumn && column <= MERGE.endColumn
      ? MERGE
      : undefined;

  function createMergeAwareController(rowCount = 20, columnCount = 10) {
    const selection = new SelectionModel({ rowCount, columnCount });
    const controller = new KeyboardController({ selection, findMergeAt });
    selection.selectCell(3, 2); // the merge's anchor
    return { selection, controller };
  }

  it("ArrowRight steps from the merge's far (right) edge, past the whole merge", () => {
    const { selection, controller } = createMergeAwareController();

    controller.handleKeyDown("ArrowRight");

    // Stepping from the anchor (column 2 + 1 = 3) would still land inside
    // the merge and re-normalize back to (3, 2) — a no-op. Stepping from
    // the far edge (column 4 + 1 = 5) actually leaves the merge.
    expect(selection.snapshot.active).toEqual({ row: 3, column: 5 });
  });

  it("ArrowDown steps from the merge's far (bottom) edge", () => {
    const { selection, controller } = createMergeAwareController();

    controller.handleKeyDown("ArrowDown");

    expect(selection.snapshot.active).toEqual({ row: 6, column: 2 });
  });

  it("ArrowLeft steps from the merge's near (left) edge, landing just outside it", () => {
    const { selection, controller } = createMergeAwareController();

    controller.handleKeyDown("ArrowLeft");

    expect(selection.snapshot.active).toEqual({ row: 3, column: 1 });
  });

  it("ArrowUp steps from the merge's near (top) edge, landing just outside it", () => {
    const { selection, controller } = createMergeAwareController();

    controller.handleKeyDown("ArrowUp");

    expect(selection.snapshot.active).toEqual({ row: 2, column: 2 });
  });

  it("Shift+ArrowDown extends the range past the merge instead of moving the anchor", () => {
    const { selection, controller } = createMergeAwareController();

    controller.handleKeyDown("ArrowDown", { shift: true });

    expect(selection.snapshot.anchor).toEqual({ row: 3, column: 2 });
    expect(selection.snapshot.active).toEqual({ row: 6, column: 2 });
  });

  it("stepping toward a merge lands on the merge's anchor, not the physically adjacent cell", () => {
    const selection = new SelectionModel({ rowCount: 20, columnCount: 10 });
    const controller = new KeyboardController({ selection, findMergeAt });
    selection.selectCell(3, 1); // just left of the merge

    controller.handleKeyDown("ArrowRight");

    expect(selection.snapshot.active).toEqual({ row: 3, column: 2 });
  });

  it("clamps at the table edge even when the merge touches it", () => {
    const selection = new SelectionModel({ rowCount: 20, columnCount: 5 });
    const rightEdgeMerge: CellRange = { startRow: 3, endRow: 5, startColumn: 3, endColumn: 4 };
    const controller = new KeyboardController({
      selection,
      findMergeAt: (row, column) =>
        row >= rightEdgeMerge.startRow &&
        row <= rightEdgeMerge.endRow &&
        column >= rightEdgeMerge.startColumn &&
        column <= rightEdgeMerge.endColumn
          ? rightEdgeMerge
          : undefined
    });
    selection.selectCell(3, 3);

    controller.handleKeyDown("ArrowRight");

    expect(selection.snapshot.active).toEqual({ row: 3, column: 3 });
  });
});

describe("KeyboardController selection-kind collapse regression (final-review Finding 2)", () => {
  // Reachable with zero merges anywhere on the sheet: HeaderSelectionController
  // calls `selectAll()`/`selectColumn()`/`selectRow()` on the same SelectionModel
  // instance KeyboardController holds. The old code called
  // `selection.moveActive(dr, dc, true)` (== `moveActiveTo(..., extend=true)`),
  // which always collapses `kind` to "cell". Using `extendTo()` instead is a
  // complete no-op when `kind === "all"` (and preserves "row"/"column" kind
  // otherwise) — so corner-click (select-all) then Shift+ArrowDown silently did
  // nothing, a real regression with no merges involved.

  it("select-all then Shift+ArrowUp moves the active cell and collapses kind to \"cell\"", () => {
    const { selection, controller } = createController(20, 10);
    selection.selectAll();
    expect(selection.snapshot.kind).toBe("all");
    expect(selection.snapshot.active).toEqual({ row: 19, column: 9 });

    controller.handleKeyDown("ArrowUp", { shift: true });

    // Bug: this stayed { kind: "all", active: { row: 19, column: 9 } } — a
    // total no-op — because `extendTo()` bails out immediately for kind "all".
    expect(selection.snapshot.kind).toBe("cell");
    expect(selection.snapshot.active).toEqual({ row: 18, column: 9 });
    expect(selection.snapshot.anchor).toEqual({ row: 0, column: 0 });
  });

  it("select-column then Shift+ArrowRight moves the active cell and collapses kind to \"cell\"", () => {
    const { selection, controller } = createController(20, 10);
    selection.selectColumn(2);
    expect(selection.snapshot.kind).toBe("column");
    expect(selection.snapshot.active).toEqual({ row: 0, column: 2 });

    controller.handleKeyDown("ArrowRight", { shift: true });

    // Bug: `extendTo()` preserves kind "column" instead of collapsing to
    // "cell", so the selection stayed a full-height column instead of
    // becoming a normal extending cell range.
    expect(selection.snapshot.kind).toBe("cell");
    expect(selection.snapshot.active).toEqual({ row: 0, column: 3 });
  });

  it("select-row then Shift+ArrowDown (no Shift) moving and collapsing still works without regressing the non-extend path", () => {
    const { selection, controller } = createController(20, 10);
    selection.selectRow(4);
    expect(selection.snapshot.kind).toBe("row");

    controller.handleKeyDown("ArrowDown");

    expect(selection.snapshot.kind).toBe("cell");
    expect(selection.snapshot.active).toEqual({ row: 5, column: 0 });
    expect(selection.snapshot.anchor).toEqual({ row: 5, column: 0 });
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
