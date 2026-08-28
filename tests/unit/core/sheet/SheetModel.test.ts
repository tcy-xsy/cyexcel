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
    // "only v" 要用 toStrictEqual 才测得到：toEqual 看不见 `{ v, t: undefined }`
    expect(model.getRecord(0, 0)).toStrictEqual({ v: "fresh" });
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

  describe("pinCoordinateSpace", () => {
    it("absorbs a shrinking occupied range into the blank counts", () => {
      const model = SheetModel.createEmpty({ blankRows: 0, blankColumns: 0 });
      model.setCell(2, 1, "a");
      expect(model.rowCount()).toBe(3);

      model.deleteCell(2, 1);
      model.pinCoordinateSpace(3, 2);

      expect(model.rowCount()).toBe(3);
      expect(model.columnCount()).toBe(2);
      expect(model.occupiedCount()).toBe(0);
    });

    it("survives a snapshot round trip — the absorbed blanks are persisted", () => {
      const model = SheetModel.createEmpty({ blankRows: 0, blankColumns: 0 });
      model.setCell(2, 1, "a");
      model.deleteCell(2, 1);
      model.pinCoordinateSpace(3, 2);

      const restored = SheetModel.fromSnapshot(model.toSnapshot());

      expect(restored.rowCount()).toBe(3);
      expect(restored.columnCount()).toBe(2);
    });

    /**
     * 钉到比占用范围还小的轴上是没有意义的请求——空白数会是负的。正常路径上
     * 到不了（越界的值写在 `TableStore` 的 prepare 期就被拒了），所以这里直接
     * 钉住"抛，而不是夹到 0"：夹的话轴会悄悄放大成占用范围，正是这个方法要防的事。
     */
    it("refuses to pin an axis below its occupied extent", () => {
      const model = SheetModel.createEmpty({ blankRows: 0, blankColumns: 0 });
      model.setCell(4, 0, "a");

      expect(() => model.pinCoordinateSpace(2, 1)).toThrow(/Cannot pin row count to 2/);
      expect(model.rowCount()).toBe(5);
    });
  });

  /**
   * 插/删改坐标空间，改的量是 `count`，与插在哪里无关。
   *
   * 从前它只在**占用范围里**成立：轴是派生量，插在占用范围之外（尾部空白区里）
   * 没有格子要搬，`extent` 与 `blank` 都不动，轴就原地不动。默认的 1000 行空白让
   * 这个洞平时只表现为"末尾插入不生效"，把空白调到 0 才看得见。
   */
  describe("coordinate space under structural changes", () => {
    it("grows and shrinks by count when the change lands inside the occupied range", () => {
      const model = SheetModel.createEmpty({ blankRows: 0, blankColumns: 0 });
      model.setCell(2, 2, "a");
      expect([model.rowCount(), model.columnCount()]).toEqual([3, 3]);

      model.insertRows(1, 2);
      model.insertColumns(0, 1);
      expect([model.rowCount(), model.columnCount()]).toEqual([5, 4]);

      model.deleteRows(0, 1);
      model.deleteColumns(0, 1);
      expect([model.rowCount(), model.columnCount()]).toEqual([4, 3]);
    });

    it("grows and shrinks by count when the change lands in the blank tail", () => {
      const model = SheetModel.createEmpty({ blankRows: 0, blankColumns: 0 });
      model.setCell(1, 1, "a");
      expect([model.rowCount(), model.columnCount()]).toEqual([2, 2]);

      // 追加在末尾：一个格子都不用搬，从前这两句什么也不会改变
      model.insertRows(2, 1);
      model.insertColumns(2, 1);
      expect([model.rowCount(), model.columnCount()]).toEqual([3, 3]);
      expect(model.getCell(1, 1)).toBe("a");

      model.deleteRows(2, 1);
      model.deleteColumns(2, 1);
      expect([model.rowCount(), model.columnCount()]).toEqual([2, 2]);
      expect(model.getCell(1, 1)).toBe("a");
    });

    it("keeps the blank tail intact when the change lands inside the occupied range", () => {
      const model = SheetModel.createEmpty({ blankRows: 5, blankColumns: 5 });
      model.setCell(1, 1, "a");
      expect([model.rowCount(), model.columnCount()]).toEqual([7, 7]);

      model.insertRows(0, 3);
      expect(model.rowCount()).toBe(10);
      // 占用范围自己涨了 3，空白不该跟着一起涨
      expect(model.toSnapshot().blankRows).toBe(5);
    });

    it("round-trips through a snapshot after appending past the occupied range", () => {
      const model = SheetModel.createEmpty({ blankRows: 0, blankColumns: 0 });
      model.setCell(0, 0, "a");
      model.insertRows(1, 4);
      model.insertColumns(1, 2);

      const restored = SheetModel.fromSnapshot(model.toSnapshot());
      expect([restored.rowCount(), restored.columnCount()]).toEqual([5, 3]);
    });

    /**
     * 移动与重排把仅有的那一行/列占用挪到更前面，占用范围自己就缩了；轴长是
     * 派生量，不钉的话会跟着凭空变短——移动改变了行数。
     */
    it("leaves the axis alone for move and reorder", () => {
      const model = SheetModel.createEmpty({ blankRows: 2, blankColumns: 2 });
      model.setCell(1, 1, "a");
      expect([model.rowCount(), model.columnCount()]).toEqual([4, 4]);

      model.moveRows(1, 1, 0);
      model.moveColumns(1, 1, 0);
      expect([model.rowCount(), model.columnCount()]).toEqual([4, 4]);
      expect(model.getCell(0, 0)).toBe("a");

      model.reorderRows(Int32Array.from([3, 1, 2, 0]));
      expect([model.rowCount(), model.columnCount()]).toEqual([4, 4]);
      expect(model.getCell(3, 0)).toBe("a");
    });
  });
});
