import { describe, expect, it } from "vitest";

import { SheetModel } from "../../../../src/core/sheet";

describe("SheetModel", () => {
  it("starts empty with the coordinate space coming from blank counts alone", () => {
    const model = SheetModel.createEmpty({ blankRows: 1000, blankColumns: 26 });
    expect(model.rowCount()).toBe(1000);
    expect(model.columnCount()).toBe(26);
    expect(model.occupiedRowCount()).toBe(0);
  });

  it("grows the coordinate space as data is entered", () => {
    // spec §3.3 后果 ①：网格随数据增长
    const model = SheetModel.createEmpty({ blankRows: 50, blankColumns: 10 });
    expect(model.rowCount()).toBe(50);
    model.setCell(19, 0, "x");
    expect(model.rowCount()).toBe(70); // 20 行数据 + 50 空白
    expect(model.occupiedRowCount()).toBe(20);
  });

  it("shrinks the coordinate space when the last occupied row is deleted", () => {
    const model = SheetModel.createEmpty({ blankRows: 50, blankColumns: 10 });
    model.setCell(5, 0, "a");
    model.setCell(19, 0, "b");
    expect(model.rowCount()).toBe(70);
    model.deleteCell(19, 0);
    expect(model.rowCount()).toBe(56); // 6 行数据 + 50 空白
  });

  it("getCell returns the bare value, not the record", () => {
    // ADR 0004 §3.6：消费方都吃裸值，接线时才不用改
    const model = SheetModel.createEmpty();
    model.setRecord(1, 1, { v: 42, t: "money", s: "s7" });
    expect(model.getCell(1, 1)).toBe(42);
    expect(model.getRecord(1, 1)).toEqual({ v: 42, t: "money", s: "s7" });
  });

  it("getCell returns undefined for an unoccupied cell", () => {
    const model = SheetModel.createEmpty();
    expect(model.getCell(3, 3)).toBeUndefined();
  });

  it("setCell preserves the other fields of an existing record", () => {
    const model = SheetModel.createEmpty();
    model.setRecord(0, 0, { v: "old", t: "money", s: "s1" });
    model.setCell(0, 0, "new");
    expect(model.getRecord(0, 0)).toEqual({ v: "new", t: "money", s: "s1" });
  });

  it("setCell on an unoccupied cell creates a record with only v", () => {
    const model = SheetModel.createEmpty();
    model.setCell(0, 0, "fresh");
    expect(model.getRecord(0, 0)).toEqual({ v: "fresh" });
  });

  it("round-trips through a snapshot", () => {
    const model = SheetModel.createEmpty({ blankRows: 7, blankColumns: 3 });
    model.setRecord(2, 1, { v: 5, t: "number" });
    const restored = SheetModel.fromSnapshot(model.toSnapshot());
    expect(restored.getRecord(2, 1)).toEqual({ v: 5, t: "number" });
    expect(restored.rowCount()).toBe(10); // 3 行数据 + 7 空白
    expect(restored.columnCount()).toBe(5); // 2 列数据 + 3 空白
  });

  it("round-trips through JSON", () => {
    const model = SheetModel.createEmpty();
    model.setCell(0, 0, "a");
    const json = JSON.parse(JSON.stringify(model.toSnapshot())) as ReturnType<
      SheetModel["toSnapshot"]
    >;
    expect(SheetModel.fromSnapshot(json).getCell(0, 0)).toBe("a");
  });

  it("rejects a negative blank count at construction", () => {
    expect(() => SheetModel.createEmpty({ blankRows: -1 })).toThrow();
  });

  it("round-trips freeze through toSnapshot and fromSnapshot", () => {
    const model = SheetModel.createEmpty({ freeze: { rows: 2, columns: 3 } });
    const restored = SheetModel.fromSnapshot(model.toSnapshot());
    const snapshot = restored.toSnapshot();
    expect(snapshot.freeze).toEqual({ rows: 2, columns: 3 });
  });

  it("does not add spurious freeze: undefined to snapshot when not set", () => {
    const model = SheetModel.createEmpty();
    const snapshot = model.toSnapshot();
    expect("freeze" in snapshot).toBe(false);
  });
});
