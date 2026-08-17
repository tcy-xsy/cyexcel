import { describe, expect, it } from "vitest";

import { CyExcelError } from "../../../../src/core/errors";
import { SelectionModel } from "../../../../src/core/interaction";

function createModel(rowCount = 20, columnCount = 10): SelectionModel {
  return new SelectionModel({ rowCount, columnCount });
}

describe("SelectionModel.selectCell", () => {
  it("sets active and anchor to the same cell with kind cell", () => {
    const model = createModel();

    const snapshot = model.selectCell(3, 4);

    expect(snapshot.kind).toBe("cell");
    expect(snapshot.active).toEqual({ row: 3, column: 4 });
    expect(snapshot.anchor).toEqual({ row: 3, column: 4 });
    expect(snapshot.range).toEqual({ startRow: 3, startColumn: 4, endRow: 3, endColumn: 4 });
  });
});

describe("SelectionModel drag normalization", () => {
  it("produces the same normalized range regardless of drag direction", () => {
    const forward = createModel();
    forward.selectCell(2, 2);
    const forwardSnapshot = forward.extendTo(5, 6);

    const backward = createModel();
    backward.selectCell(5, 6);
    const backwardSnapshot = backward.extendTo(2, 2);

    const expectedRange = { startRow: 2, startColumn: 2, endRow: 5, endColumn: 6 };
    expect(forwardSnapshot.range).toEqual(expectedRange);
    expect(backwardSnapshot.range).toEqual(expectedRange);
  });
});

describe("SelectionModel.selectRow/selectColumn/selectAll", () => {
  it("selectRow spans the full row width", () => {
    const model = createModel(20, 10);

    const snapshot = model.selectRow(4);

    expect(snapshot.kind).toBe("row");
    expect(snapshot.range).toEqual({ startRow: 4, startColumn: 0, endRow: 4, endColumn: 9 });
  });

  it("extending a row selection to another row spans both rows, full width", () => {
    const model = createModel(20, 10);
    model.selectRow(4);

    const snapshot = model.extendTo(7, 0);

    expect(snapshot.range).toEqual({ startRow: 4, startColumn: 0, endRow: 7, endColumn: 9 });
  });

  it("selectColumn spans the full column height", () => {
    const model = createModel(20, 10);

    const snapshot = model.selectColumn(3);

    expect(snapshot.kind).toBe("column");
    expect(snapshot.range).toEqual({ startRow: 0, startColumn: 3, endRow: 19, endColumn: 3 });
  });

  it("selectAll spans the whole table", () => {
    const model = createModel(20, 10);

    const snapshot = model.selectAll();

    expect(snapshot.kind).toBe("all");
    expect(snapshot.range).toEqual({ startRow: 0, startColumn: 0, endRow: 19, endColumn: 9 });
  });

  it("extendTo is a no-op once kind is all", () => {
    const model = createModel(20, 10);
    model.selectAll();

    const snapshot = model.extendTo(1, 1);

    expect(snapshot.kind).toBe("all");
    expect(snapshot.range).toEqual({ startRow: 0, startColumn: 0, endRow: 19, endColumn: 9 });
  });
});

describe("SelectionModel snapshot shape", () => {
  it("only exposes numeric row/column indices, never pixel coordinates", () => {
    const model = createModel();

    const snapshot = model.selectCell(1, 1);

    expect(typeof snapshot.active.row).toBe("number");
    expect(typeof snapshot.active.column).toBe("number");
    expect(snapshot).not.toHaveProperty("x");
    expect(snapshot).not.toHaveProperty("y");
  });

  it("returns a frozen snapshot", () => {
    const model = createModel();

    const snapshot = model.selectCell(0, 0);

    expect(Object.isFrozen(snapshot)).toBe(true);
  });
});

describe("SelectionModel has no readonly/data-source concept", () => {
  it("selects and reads freely, constructed from plain bounds with no DataSource involved", () => {
    // SelectionModel is constructed from plain { rowCount, columnCount } —
    // it never imports DataSource/TableStore, so there is nothing to
    // configure as "readonly" here. That's exactly why selection (and copy,
    // later) stays available in a readonly sheet: there is no readonly
    // check to bypass in the first place.
    const model = new SelectionModel({ rowCount: 5, columnCount: 5 });

    expect(model.selectCell(2, 2).kind).toBe("cell");
    expect(model.selectAll().kind).toBe("all");
  });
});

describe("SelectionModel bounds", () => {
  it("clamps out-of-range and negative coordinates instead of throwing", () => {
    const model = createModel(5, 5);

    expect(model.selectCell(-3, -3).active).toEqual({ row: 0, column: 0 });
    expect(model.selectCell(100, 100).active).toEqual({ row: 4, column: 4 });
  });

  it("clamps to 0 for an empty table without throwing", () => {
    const model = createModel(0, 0);

    expect(model.selectCell(5, 5).active).toEqual({ row: 0, column: 0 });
  });

  it.each([
    ["negative rowCount", { rowCount: -1, columnCount: 5 }],
    ["non-integer columnCount", { rowCount: 5, columnCount: 1.5 }]
  ])("rejects %s", (_label, bounds) => {
    expect(() => new SelectionModel(bounds)).toThrow(CyExcelError);
  });
});

describe("SelectionModel.setBounds", () => {
  it("without a rowChange, just clamps the current selection into the new bounds", () => {
    const model = createModel(10, 5);
    model.selectCell(8, 2);

    model.setBounds({ rowCount: 5, columnCount: 5 });

    expect(model.snapshot.active).toEqual({ row: 4, column: 2 });
    expect(model.bounds).toEqual({ rowCount: 5, columnCount: 5 });
  });

  it("with an insert rowChange, follows the selected row to its new position", () => {
    const model = createModel(10, 5);
    model.selectCell(6, 2);

    model.setBounds(
      { rowCount: 12, columnCount: 5 },
      { operation: "insert", start: 2, count: 2 }
    );

    expect(model.snapshot.active.row).toBe(8); // row 6 -> 8 after inserting 2 rows above it
  });

  it("with a delete rowChange, a selection outside the deleted band follows its row", () => {
    const model = createModel(10, 5);
    model.selectCell(8, 2);

    model.setBounds(
      { rowCount: 7, columnCount: 5 },
      { operation: "delete", start: 2, count: 3 }
    );

    expect(model.snapshot.active.row).toBe(5); // row 8 -> 5 after deleting 3 rows above it
  });

  it("with a delete rowChange, a selection that was inside the deleted band lands on the deletion point (clamped into the new bounds)", () => {
    const model = createModel(10, 5);
    model.selectCell(3, 2); // inside [2, 5)

    model.setBounds(
      { rowCount: 7, columnCount: 5 },
      { operation: "delete", start: 2, count: 3 }
    );

    expect(model.snapshot.active.row).toBe(2);
  });

  it("with a move rowChange, a selection on the moved row follows it to the new position", () => {
    const model = createModel(6, 5);
    model.selectCell(1, 2); // one of the two rows being moved (rows 1-2)

    model.setBounds(
      { rowCount: 6, columnCount: 5 },
      { operation: "move", start: 1, count: 2, target: 5 }
    );

    // [0,1,2,3,4,5] -> moveRows(1,2,5) -> row 1 ends up at index 3.
    expect(model.snapshot.active.row).toBe(3);
  });

  it("also remaps the anchor, not just the active cell, so a range selection stays intact", () => {
    const model = createModel(10, 5);
    model.selectCell(4, 1);
    model.extendTo(6, 3); // anchor stays at (4,1), active moves to (6,3)

    model.setBounds(
      { rowCount: 12, columnCount: 5 },
      { operation: "insert", start: 0, count: 2 }
    );

    expect(model.snapshot.anchor.row).toBe(6);
    expect(model.snapshot.active.row).toBe(8);
  });

  it("rejects invalid bounds the same way the constructor does", () => {
    const model = createModel(5, 5);

    expect(() => model.setBounds({ rowCount: -1, columnCount: 5 })).toThrow(CyExcelError);
  });
});

describe("SelectionModel.setBounds — columnChange", () => {
  it("with an insert columnChange, follows the selected column to its new position", () => {
    const model = createModel(5, 10);
    model.selectCell(2, 6);

    model.setBounds(
      { rowCount: 5, columnCount: 12 },
      undefined,
      { operation: "insert", start: 2, count: 2 }
    );

    expect(model.snapshot.active.column).toBe(8); // column 6 -> 8 after inserting 2 columns before it
  });

  it("with a delete columnChange, a selection outside the deleted band follows its column", () => {
    const model = createModel(5, 10);
    model.selectCell(2, 8);

    model.setBounds(
      { rowCount: 5, columnCount: 7 },
      undefined,
      { operation: "delete", start: 2, count: 3 }
    );

    expect(model.snapshot.active.column).toBe(5); // column 8 -> 5 after deleting 3 columns before it
  });

  it("with a move columnChange, a selection on the moved column follows it to the new position", () => {
    const model = createModel(5, 6);
    model.selectCell(2, 1); // one of the two columns being moved (columns 1-2)

    model.setBounds(
      { rowCount: 5, columnCount: 6 },
      undefined,
      { operation: "move", start: 1, count: 2, target: 5 }
    );

    expect(model.snapshot.active.column).toBe(3);
  });

  it("also remaps the anchor, not just the active cell", () => {
    const model = createModel(5, 10);
    model.selectCell(1, 4);
    model.extendTo(3, 6); // anchor stays at (1,4), active moves to (3,6)

    model.setBounds(
      { rowCount: 5, columnCount: 12 },
      undefined,
      { operation: "insert", start: 0, count: 2 }
    );

    expect(model.snapshot.anchor.column).toBe(6);
    expect(model.snapshot.active.column).toBe(8);
  });

  it("remaps every frozen region in #otherRanges, not just the active one", () => {
    const model = createModel(5, 10);
    model.selectCell(1, 1);
    model.addRange(1, 6);

    model.setBounds(
      { rowCount: 5, columnCount: 12 },
      undefined,
      { operation: "insert", start: 0, count: 2 }
    );

    expect(model.snapshot.ranges).toEqual([
      { startRow: 1, startColumn: 3, endRow: 1, endColumn: 3 },
      { startRow: 1, startColumn: 8, endRow: 1, endColumn: 8 }
    ]);
  });

  it("rowChange and columnChange can both apply in the same setBounds call", () => {
    const model = createModel(10, 10);
    model.selectCell(5, 5);

    model.setBounds(
      { rowCount: 12, columnCount: 12 },
      { operation: "insert", start: 0, count: 2 },
      { operation: "insert", start: 0, count: 2 }
    );

    expect(model.snapshot.active).toEqual({ row: 7, column: 7 });
  });
});

describe("SelectionModel multi-region selection", () => {
  it("range always equals the last entry in ranges", () => {
    const model = createModel();

    const snapshot = model.selectCell(2, 2);

    expect(snapshot.ranges).toEqual([snapshot.range]);
  });

  it("addRange freezes the current active range and starts a new one", () => {
    const model = createModel();
    model.selectCell(1, 1);

    const snapshot = model.addRange(3, 3);

    expect(snapshot.ranges).toEqual([
      { startRow: 1, startColumn: 1, endRow: 1, endColumn: 1 },
      { startRow: 3, startColumn: 3, endRow: 3, endColumn: 3 }
    ]);
    expect(snapshot.range).toEqual({ startRow: 3, startColumn: 3, endRow: 3, endColumn: 3 });
    expect(snapshot.kind).toBe("cell");
  });

  it("extendTo after addRange only extends the new active range, not the frozen one", () => {
    const model = createModel();
    model.selectCell(1, 1);
    model.addRange(3, 3);

    const snapshot = model.extendTo(5, 5);

    expect(snapshot.ranges).toEqual([
      { startRow: 1, startColumn: 1, endRow: 1, endColumn: 1 },
      { startRow: 3, startColumn: 3, endRow: 5, endColumn: 5 }
    ]);
  });

  it("addRow freezes the active range and starts a new full-width row selection", () => {
    const model = createModel(20, 10);
    model.selectCell(1, 1);

    const snapshot = model.addRow(4);

    expect(snapshot.kind).toBe("row");
    expect(snapshot.ranges).toEqual([
      { startRow: 1, startColumn: 1, endRow: 1, endColumn: 1 },
      { startRow: 4, startColumn: 0, endRow: 4, endColumn: 9 }
    ]);
  });

  it("addColumn freezes the active range and starts a new full-height column selection", () => {
    const model = createModel(20, 10);
    model.selectCell(1, 1);

    const snapshot = model.addColumn(3);

    expect(snapshot.kind).toBe("column");
    expect(snapshot.ranges).toEqual([
      { startRow: 1, startColumn: 1, endRow: 1, endColumn: 1 },
      { startRow: 0, startColumn: 3, endRow: 19, endColumn: 3 }
    ]);
  });

  it.each([
    ["selectCell", (model: SelectionModel) => model.selectCell(0, 0)],
    ["selectRow", (model: SelectionModel) => model.selectRow(0)],
    ["selectColumn", (model: SelectionModel) => model.selectColumn(0)],
    ["selectAll", (model: SelectionModel) => model.selectAll()]
  ])("%s clears any previously frozen regions", (_label, act) => {
    const model = createModel();
    model.selectCell(1, 1);
    model.addRange(3, 3);
    expect(model.snapshot.ranges.length).toBe(2);

    const snapshot = act(model);

    expect(snapshot.ranges.length).toBe(1);
  });

  it("moveActiveTo without extend clears frozen regions (plain arrow-key navigation collapses multi-select)", () => {
    const model = createModel();
    model.selectCell(1, 1);
    model.addRange(3, 3);

    const snapshot = model.moveActiveTo(5, 5);

    expect(snapshot.ranges.length).toBe(1);
  });

  it("moveActiveTo with extend keeps frozen regions intact and only extends the active one", () => {
    const model = createModel();
    model.selectCell(1, 1);
    model.addRange(3, 3);

    const snapshot = model.moveActiveTo(5, 5, true);

    expect(snapshot.ranges.length).toBe(2);
    expect(snapshot.range).toEqual({ startRow: 3, startColumn: 3, endRow: 5, endColumn: 5 });
  });

  it("setBounds with a rowChange remaps every frozen region, not just the active one", () => {
    const model = createModel(10, 5);
    model.selectCell(1, 1);
    model.addRange(6, 2);

    model.setBounds(
      { rowCount: 12, columnCount: 5 },
      { operation: "insert", start: 0, count: 2 }
    );

    expect(model.snapshot.ranges).toEqual([
      { startRow: 3, startColumn: 1, endRow: 3, endColumn: 1 },
      { startRow: 8, startColumn: 2, endRow: 8, endColumn: 2 }
    ]);
  });
});
