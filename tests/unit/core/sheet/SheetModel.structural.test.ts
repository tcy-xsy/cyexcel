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

  it("deleteRows past the occupied extent keeps the rows it never covered", () => {
    // 坐标空间是 maxOccupied + 1 + blankRows（spec §3.3），所以 start/count
    // 落在占用范围之外是合法的日常操作，不能连带把范围内的行丢掉。
    const model = SheetModel.createEmpty();
    for (let row = 0; row < 5; row += 1) {
      model.setCell(row, 0, `r${row}`);
    }
    model.deleteRows(3, 10);
    expect([0, 1, 2].map((row) => model.getCell(row, 0))).toEqual(["r0", "r1", "r2"]);
    expect(model.occupiedRowCount()).toBe(3);
  });

  it("deleteRows entirely inside the blank area is a no-op", () => {
    const model = SheetModel.createEmpty();
    for (let row = 0; row < 5; row += 1) {
      model.setCell(row, 0, `r${row}`);
    }
    model.deleteRows(10, 2);
    expect([0, 1, 2, 3, 4].map((row) => model.getCell(row, 0))).toEqual([
      "r0",
      "r1",
      "r2",
      "r3",
      "r4"
    ]);
  });

  it("moveRows into the blank area lands on the target index, gap intact", () => {
    const model = SheetModel.createEmpty();
    for (let row = 0; row < 5; row += 1) {
      model.setCell(row, 0, `r${row}`);
    }
    model.moveRows(0, 2, 8);
    expect(model.getCell(6, 0)).toBe("r0");
    expect(model.getCell(7, 0)).toBe("r1");
    expect([0, 1, 2].map((row) => model.getCell(row, 0))).toEqual(["r2", "r3", "r4"]);
  });

  it("moveRows of a block that reaches into the blank area keeps the gap", () => {
    const model = SheetModel.createEmpty();
    for (let row = 0; row < 5; row += 1) {
      model.setCell(row, 0, `r${row}`);
    }
    model.moveRows(3, 10, 0);
    expect(model.getCell(0, 0)).toBe("r3");
    expect(model.getCell(1, 0)).toBe("r4");
    expect(model.getCell(10, 0)).toBe("r0");
    expect(model.getCell(12, 0)).toBe("r2");
  });

  it("insertRows at a point in the blank area leaves the occupied rows alone", () => {
    const model = SheetModel.createEmpty();
    model.setCell(0, 0, "a");
    model.setCell(1, 0, "b");
    model.insertRows(5, 1);
    expect(model.getCell(0, 0)).toBe("a");
    expect(model.getCell(1, 0)).toBe("b");
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
  function modelWithMerge(merge: { startRow: number; startColumn: number; endRow: number; endColumn: number }): SheetModel {
    const base = SheetModel.createEmpty();
    base.setCell(merge.startRow, merge.startColumn, "a");
    const snapshot = base.toSnapshot();
    return SheetModel.fromSnapshot({ ...snapshot, mergeData: [merge] });
  }

  it("exposes merges through an accessor", () => {
    const model = modelWithMerge({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 });
    expect(model.merges()).toEqual([{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }]);
  });

  it("setMerges replaces the merge list wholesale", () => {
    const model = modelWithMerge({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 });
    model.setMerges([{ startRow: 2, startColumn: 2, endRow: 3, endColumn: 3 }]);
    expect(model.merges()).toEqual([{ startRow: 2, startColumn: 2, endRow: 3, endColumn: 3 }]);
  });

  it("insertRows shifts a merge that starts at or after the insertion point", () => {
    // merge spans row 2 only; insert(1, 2) 落在合并区起点之前 => 整体下移 2 行
    const model = modelWithMerge({ startRow: 2, startColumn: 0, endRow: 2, endColumn: 0 });
    model.insertRows(1, 2);
    expect(model.merges()).toEqual([{ startRow: 4, startColumn: 0, endRow: 4, endColumn: 0 }]);
  });

  it("deleteRows that fully covers a merge removes it", () => {
    // merge 只占 row 0；delete(0, 1) 全覆盖 => 合并区消失
    const model = modelWithMerge({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 });
    model.deleteRows(0, 1);
    expect(model.merges()).toEqual([]);
  });

  it("moveRows carries a merge that sits entirely inside the moved block", () => {
    // merge 只占 row 0，block = row [0,0]，target=1 是恒等移动 => 合并区不变
    const model = modelWithMerge({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 });
    model.moveRows(0, 1, 1);
    expect(model.merges()).toEqual([{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }]);
  });

  it("moveRows rejects a merge that straddles the move boundary", () => {
    // merge 跨 rows [0,1]，block 只有 row 0：merge 一部分在块内一部分在块外
    const model = modelWithMerge({ startRow: 0, startColumn: 0, endRow: 1, endColumn: 0 });
    let error: unknown;
    try {
      model.moveRows(0, 1, 2);
    } catch (e) {
      error = e;
    }
    expect((error as CyExcelError).code).toBe(CyExcelErrorCode.MergeSpansMoveBoundary);
    // 拒绝时合并区保持原样，没有被部分改写
    expect(model.merges()).toEqual([{ startRow: 0, startColumn: 0, endRow: 1, endColumn: 0 }]);
  });

  it("insertColumns shifts a merge that starts at or after the insertion point", () => {
    const model = modelWithMerge({ startRow: 0, startColumn: 2, endRow: 0, endColumn: 2 });
    model.insertColumns(1, 2);
    expect(model.merges()).toEqual([{ startRow: 0, startColumn: 4, endRow: 0, endColumn: 4 }]);
  });

  it("deleteColumns that fully covers a merge removes it", () => {
    const model = modelWithMerge({ startRow: 0, startColumn: 0, endRow: 1, endColumn: 0 });
    model.deleteColumns(0, 1);
    expect(model.merges()).toEqual([]);
  });

  it("moveColumns carries a merge that sits entirely inside the moved block", () => {
    const model = modelWithMerge({ startRow: 0, startColumn: 0, endRow: 1, endColumn: 0 });
    model.moveColumns(0, 1, 1);
    expect(model.merges()).toEqual([{ startRow: 0, startColumn: 0, endRow: 1, endColumn: 0 }]);
  });

  it("moveColumns rejects a merge that straddles the move boundary", () => {
    // merge 跨 columns [0,1]，block 只有 column 0
    const model = modelWithMerge({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 });
    let error: unknown;
    try {
      model.moveColumns(0, 1, 2);
    } catch (e) {
      error = e;
    }
    expect((error as CyExcelError).code).toBe(CyExcelErrorCode.MergeSpansMoveBoundary);
    expect(model.merges()).toEqual([{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }]);
  });

  it("reorderRows still refuses unconditionally while merges exist", () => {
    // reorderRows 接受任意排列，没有良定义的"合并区怎么跟着走"，所以它继续
    // 无条件拒绝——这是与另外六个方法唯一不同的地方，全局约束 3。
    const model = modelWithMerge({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 });
    let error: unknown;
    try {
      model.reorderRows(Int32Array.from([0]));
    } catch (e) {
      error = e;
    }
    expect((error as CyExcelError).code).toBe(CyExcelErrorCode.MutationNotSupported);
  });

  it("allows structural operations when there are no merges", () => {
    const model = SheetModel.createEmpty();
    model.setCell(0, 0, "a");
    expect(() => model.insertRows(0, 1)).not.toThrow();
  });
});

describe("SheetModel structural argument validation", () => {
  function populated(): SheetModel {
    const model = SheetModel.createEmpty();
    model.setCell(0, 0, "a");
    model.setCell(1, 1, "b");
    return model;
  }

  it("rejects every illegal argument shape with a RangeError", () => {
    expect.hasAssertions();
    const cases: ReadonlyArray<readonly [string, (m: SheetModel) => void]> = [
      // start = -2、count = 1：`newRow - count` 会算出 -1，而 -1 就是
      // ROW_PLAN_HOLE——哨兵撞车，把一行真数据变成一个洞
      ["insertRows negative start", (m) => m.insertRows(-2, 1)],
      ["insertRows zero count", (m) => m.insertRows(0, 0)],
      ["deleteRows negative start", (m) => m.deleteRows(-1, 1)],
      ["deleteRows zero count", (m) => m.deleteRows(0, 0)],
      ["moveRows negative target", (m) => m.moveRows(0, 1, -1)],
      ["moveRows target inside the block", (m) => m.moveRows(0, 2, 1)],
      ["insertColumns negative start", (m) => m.insertColumns(-1, 1)],
      ["deleteColumns zero count", (m) => m.deleteColumns(0, 0)],
      // 落点在块内会算出 insertAt = -1，产出负的列键
      ["moveColumns target inside the block", (m) => m.moveColumns(0, 2, 1)]
    ];
    for (const [name, run] of cases) {
      expect(() => run(populated()), name).toThrow(RangeError);
    }
  });

  it("rejects a delete that runs past the coordinate space", () => {
    const model = SheetModel.createEmpty({ blankRows: 2 });
    model.setCell(0, 0, "a"); // rowCount = 1 + 2 = 3
    expect(() => model.deleteRows(2, 5)).toThrow(RangeError);
  });

  it("allows operations that only reach into the blank area", () => {
    // 空白区里的行是坐标空间的一部分，删它们是合法的（Excel 也允许）——
    // 用占用范围当轴长会把这类操作误挡掉。
    const model = SheetModel.createEmpty();
    model.setCell(0, 0, "a");
    expect(() => model.deleteRows(10, 2)).not.toThrow();
    expect(model.getCell(0, 0)).toBe("a");
  });

  it("validates arguments before the merge guard", () => {
    // 参数非法与有没有合并区无关，先报 MutationNotSupported 会误导调用方
    const base = SheetModel.createEmpty();
    base.setCell(0, 0, "a");
    const model = SheetModel.fromSnapshot({
      ...base.toSnapshot(),
      mergeData: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }]
    });
    expect(() => model.insertRows(-2, 1)).toThrow(RangeError);
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
