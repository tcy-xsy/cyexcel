import { describe, expect, it } from "vitest";

import { CyExcelError, SelectionModel } from "../../../../src/core";

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
