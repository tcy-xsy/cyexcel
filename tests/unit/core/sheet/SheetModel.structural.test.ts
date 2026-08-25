import { describe, expect, it } from "vitest";

import type { CyExcelError } from "../../../../src/core/errors";
import { SheetModel } from "../../../../src/core/sheet";
import { CyExcelErrorCode } from "../../../../src/core/errors";

describe("SheetModel row operations", () => {
  it("insertRows opens blank rows and pushes data down", () => {
    const model = SheetModel.createEmpty({ blankRows: 10 });
    model.setCell(0, 0, "a");
    model.setCell(1, 0, "b");
    model.insertRows(1, 2);
    expect(model.getCell(0, 0)).toBe("a");
    expect(model.getCell(1, 0)).toBeUndefined();
    expect(model.getCell(2, 0)).toBeUndefined();
    expect(model.getCell(3, 0)).toBe("b");
  });

  it("insertRows carries row metadata with the row it belongs to", () => {
    const model = SheetModel.createEmpty();
    model.setCell(1, 0, "x");
    model.setRowMeta(1, { h: 40, hp: "manual" });
    model.insertRows(0, 1);
    expect(model.getRowMeta(1)).toBeUndefined();
    expect(model.getRowMeta(2)).toEqual({ h: 40, hp: "manual" });
  });

  it("deleteRows removes the range and closes the gap", () => {
    const model = SheetModel.createEmpty();
    model.setCell(0, 0, "a");
    model.setCell(1, 0, "gone");
    model.setCell(2, 0, "c");
    model.deleteRows(1, 1);
    expect(model.getCell(0, 0)).toBe("a");
    expect(model.getCell(1, 0)).toBe("c");
    expect(model.occupiedRowCount()).toBe(2);
  });

  it("deleteRows drops the metadata of deleted rows", () => {
    const model = SheetModel.createEmpty();
    model.setCell(0, 0, "a");
    model.setRowMeta(0, { h: 40 });
    model.setRowMeta(1, { h: 60 });
    model.deleteRows(0, 1);
    expect(model.getRowMeta(0)).toEqual({ h: 60 });
  });

  it("moveRows relocates a block", () => {
    const model = SheetModel.createEmpty();
    model.setCell(0, 0, "a");
    model.setCell(1, 0, "b");
    model.setCell(2, 0, "c");
    // target 是"落在原第几行之前"（同 moveColumns、同仓库既有 moveRows）：
    // 把行 0 移到原行 2 之前 => b 上移到 0，a 落在 b 之后，c 不动。
    model.moveRows(0, 1, 2);
    expect(model.getCell(0, 0)).toBe("b");
    expect(model.getCell(1, 0)).toBe("a");
    expect(model.getCell(2, 0)).toBe("c");
  });

  it("reorderRows still works through the facade", () => {
    const model = SheetModel.createEmpty();
    model.setCell(0, 0, "a");
    model.setCell(1, 0, "b");
    model.reorderRows(Int32Array.from([1, 0]));
    expect(model.getCell(0, 0)).toBe("b");
    expect(model.getCell(1, 0)).toBe("a");
  });

  it("the coordinate space follows an insert", () => {
    const model = SheetModel.createEmpty({ blankRows: 50 });
    model.setCell(0, 0, "a");
    expect(model.rowCount()).toBe(51);
    model.insertRows(0, 5);
    expect(model.rowCount()).toBe(56);
  });

  it("reorderRows refuses when order is too short for metadata extent", () => {
    const model = SheetModel.createEmpty();
    model.setCell(0, 0, "a");
    model.setRowMeta(3, { h: 50 });
    expect(() => {
      model.reorderRows(Int32Array.from([0]));
    }).toThrow();
  });

  it("reorderRows carries metadata to new index when order covers extent", () => {
    const model = SheetModel.createEmpty();
    model.setCell(0, 0, "a");
    model.setRowMeta(1, { h: 60 });
    model.reorderRows(Int32Array.from([1, 0]));
    expect(model.getRowMeta(0)).toEqual({ h: 60 });
    expect(model.getRowMeta(1)).toBeUndefined();
  });
});

describe("SheetModel column operations", () => {
  it("insertColumns shifts cells and column metadata together", () => {
    const model = SheetModel.createEmpty();
    model.setCell(0, 1, "b");
    model.setColumnMeta(1, { w: 200 });
    model.insertColumns(1, 1);
    expect(model.getCell(0, 2)).toBe("b");
    expect(model.getColumnMeta(1)).toBeUndefined();
    expect(model.getColumnMeta(2)).toEqual({ w: 200 });
  });

  it("deleteColumns removes cells and their column metadata", () => {
    const model = SheetModel.createEmpty();
    model.setCell(0, 0, "a");
    model.setCell(0, 1, "gone");
    model.setColumnMeta(1, { w: 200 });
    model.deleteColumns(1, 1);
    expect(model.getCell(0, 1)).toBeUndefined();
    expect(model.getColumnMeta(1)).toBeUndefined();
    expect(model.occupiedCount()).toBe(1);
  });

  it("moveColumns relocates cells", () => {
    const model = SheetModel.createEmpty();
    model.setCell(0, 0, "a");
    model.setCell(0, 1, "b");
    model.setCell(0, 2, "c");
    // 把列 0 移到原列 2 之前（仓库既有 target 口径）=> b, a, c
    // 注意 moveColumns(0, 1, 1) 是**空操作**——把一个块移到紧跟其后那一列
    // 之前，等于没动。用它当用例什么都测不到。
    model.moveColumns(0, 1, 2);
    expect(model.getCell(0, 0)).toBe("b");
    expect(model.getCell(0, 1)).toBe("a");
    expect(model.getCell(0, 2)).toBe("c");
  });
});

describe("SheetModel merge guard", () => {
  function modelWithMerge(): SheetModel {
    const base = SheetModel.createEmpty();
    base.setCell(0, 0, "a");
    const snapshot = base.toSnapshot();
    return SheetModel.fromSnapshot({
      ...snapshot,
      mergeData: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }]
    });
  }

  it("exposes merges through an accessor", () => {
    expect(modelWithMerge().merges()).toEqual([
      { startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }
    ]);
  });

  it("refuses every structural operation while merges exist", () => {
    expect.hasAssertions();
    const operations: ReadonlyArray<readonly [string, (m: SheetModel) => void]> = [
      ["insertRows", (m) => m.insertRows(0, 1)],
      ["deleteRows", (m) => m.deleteRows(0, 1)],
      ["moveRows", (m) => m.moveRows(0, 1, 1)],
      ["reorderRows", (m) => m.reorderRows(Int32Array.from([0]))],
      ["insertColumns", (m) => m.insertColumns(0, 1)],
      ["deleteColumns", (m) => m.deleteColumns(0, 1)],
      ["moveColumns", (m) => m.moveColumns(0, 1, 1)]
    ];
    for (const [name, run] of operations) {
      let threw = false;
      try {
        run(modelWithMerge());
      } catch (error) {
        threw = true;
        expect((error as CyExcelError).code).toBe(CyExcelErrorCode.MutationNotSupported);
      }
      expect(threw, `${name} should refuse while merges exist`).toBe(true);
    }
  });

  it("allows structural operations when there are no merges", () => {
    const model = SheetModel.createEmpty();
    model.setCell(0, 0, "a");
    expect(() => model.insertRows(0, 1)).not.toThrow();
  });
});

describe("SheetModel counts and metadata accessors", () => {
  it("occupiedCount counts cells, occupiedColumnCount counts the column extent", () => {
    const model = SheetModel.createEmpty();
    model.setCell(0, 0, "a");
    model.setCell(3, 7, "b");
    expect(model.occupiedCount()).toBe(2);
    expect(model.occupiedColumnCount()).toBe(8);
    expect(model.occupiedRowCount()).toBe(4);
  });

  it("both counts are zero on an empty model", () => {
    const model = SheetModel.createEmpty();
    expect(model.occupiedCount()).toBe(0);
    expect(model.occupiedColumnCount()).toBe(0);
  });

  it("row and column metadata survive a snapshot round trip", () => {
    const model = SheetModel.createEmpty();
    model.setRowMeta(2, { h: 44, hp: "auto" });
    model.setColumnMeta(3, { w: 180, sortable: false });
    const restored = SheetModel.fromSnapshot(model.toSnapshot());
    expect(restored.getRowMeta(2)).toEqual({ h: 44, hp: "auto" });
    expect(restored.getColumnMeta(3)).toEqual({ w: 180, sortable: false });
  });
});
