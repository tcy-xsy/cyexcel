import { describe, expect, it } from "vitest";

import { TableStore } from "../../../../src/core/data";
import { CyExcelError, CyExcelErrorCode } from "../../../../src/core/errors";
import { rowsToWorkbookSnapshot } from "../../../../src/core/sheet/import";
import { SheetModel } from "../../../../src/core/sheet/SheetModel";
import { WorkbookModel } from "../../../../src/core/sheet/WorkbookModel";

function createStore() {
  const sheet = WorkbookModel.fromSnapshot(
    rowsToWorkbookSnapshot(
      [
        { name: "A", price: 1, tag: "x" },
        { name: "B", price: 2, tag: "y" },
        { name: "C", price: 3, tag: "z" }
      ],
      ["name", "price", "tag"]
    )
  ).sheet();
  return { sheet, store: new TableStore(sheet) };
}

describe("TableStore.setCell", () => {
  it("emits exactly one cell-value patch and bumps the version by one", () => {
    const { store } = createStore();

    const result = store.setCell(1, 1, 99);

    expect(result.patches).toEqual([{ type: "cell-value", row: 1, column: 1 }]);
    expect(result.changeSet).toEqual({
      modelVersion: 1,
      changes: [{ type: "cell-value", row: 1, column: 1 }]
    });
    expect(result.version).toBe(1);
    expect(store.version).toBe(1);
    expect(store.getCell(1, 1)).toBe(99);
  });

  it("increments the version by exactly one per successful commit", () => {
    const { store } = createStore();

    store.setCell(0, 0, "a");
    store.setCell(0, 0, "b");
    store.setCell(0, 0, "c");

    expect(store.version).toBe(3);
  });
});

describe("TableStore.setRange", () => {
  it("emits exactly one range-value patch for the whole range", () => {
    const { store } = createStore();

    const result = store.setRange(
      { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 },
      [
        ["A2", 10],
        ["B2", 20]
      ]
    );

    expect(result.patches).toEqual([
      {
        type: "range-value",
        range: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 }
      }
    ]);
    expect(result.version).toBe(1);
    expect(store.getCell(0, 0)).toBe("A2");
    expect(store.getCell(0, 1)).toBe(10);
    expect(store.getCell(1, 0)).toBe("B2");
    expect(store.getCell(1, 1)).toBe(20);
  });
});

describe("TableStore.updateCell", () => {
  it("reads the current value, writes the updater's result, and emits one patch", () => {
    // 不再有"就地改一个复杂值，store 看不见"那条测试了：`CellValue` 收紧成
    // string | number | boolean 之后，格子里根本没有可以就地改的东西，那件事由
    // 类型保证而不是由这条断言保证。
    const { store } = createStore();
    expect(store.version).toBe(0);

    const result = store.updateCell(0, 2, (value) => `${String(value)}y`);

    expect(result.patches).toEqual([{ type: "cell-value", row: 0, column: 2 }]);
    expect(store.version).toBe(1);
    expect(store.getCell(0, 2)).toBe("xy");
  });
});

describe("TableStore failure paths do not bump the version", () => {
  it.each([
    ["negative row", () => {
      const { store } = createStore();
      return () => store.setCell(-1, 0, 1);
    }],
    ["non-integer column", () => {
      const { store } = createStore();
      return () => store.setCell(0, 1.5, 1);
    }],
    ["column out of range", () => {
      const { store } = createStore();
      return () => store.setCell(0, 99, 1);
    }]
  ])("%s throws a CyExcelError and leaves version at 0", (_label, setup) => {
    const attempt = setup();

    expect(attempt).toThrow(CyExcelError);
  });

  /**
   * 稀疏矩阵会痛快地把第 5 行建出来，所以「越界写」不再由存储自己拒绝——
   * 挡它的是 `TableStore`。这两条钉的就是那道闸：值写永远不改坐标空间，
   * 否则表在渲染层不知情的情况下变长/变宽（只发了一条 `cell-value`）。
   */
  it("a write past the last row is refused, and the row axis does not grow", () => {
    const { sheet, store } = createStore();

    expect(() => store.setCell(5, 0, "B")).toThrow(CyExcelError);
    expect(store.version).toBe(0);
    expect(sheet.rowCount()).toBe(3);
  });

  it("a write past the last column is refused with column-not-found", () => {
    const { sheet, store } = createStore();

    try {
      store.setCell(0, 3, "B");
      expect.unreachable("expected setCell to throw");
    } catch (error) {
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.ColumnNotFound);
    }
    expect(sheet.columnCount()).toBe(3);
  });

  it("setRange refuses a range that runs past the last row", () => {
    const { sheet, store } = createStore();

    expect(() =>
      store.setRange({ startRow: 2, startColumn: 0, endRow: 3, endColumn: 0 }, [["x"], ["y"]])
    ).toThrow(CyExcelError);
    expect(store.version).toBe(0);
    expect(sheet.rowCount()).toBe(3);
  });
});

/**
 * 越界写有闸挡着（上面那三条），但轴还能被**合法**的值写挪动：派生轴是
 * `maxOccupied + 1 + blank`，清空尾行会让它变短，填一行空白行会让它变长。两种都只
 * 发一条值补丁，渲染层收不到"表变长了"的消息。`SheetModel.pinCoordinateSpace` 把
 * 差额记进 `blank`，这一组钉的是那条不变量的两个方向。
 */
describe("TableStore — value writes never move the axes", () => {
  it("clearing the last row's cells keeps the row axis at three", () => {
    const { sheet, store } = createStore();

    store.setRange(
      { startRow: 2, startColumn: 0, endRow: 2, endColumn: 2 },
      [[undefined, undefined, undefined]]
    );

    expect(store.getCell(2, 0)).toBeUndefined();
    // 占用范围确实缩了——轴没跟着缩，是因为差额进了 blankRows
    expect(sheet.occupiedRowCount()).toBe(2);
    expect(sheet.rowCount()).toBe(3);
    expect(store.getRowCount()).toBe(3);
  });

  it("clearing every cell keeps the axes instead of collapsing the sheet", () => {
    const { sheet, store } = createStore();

    store.setRange(
      { startRow: 0, startColumn: 0, endRow: 2, endColumn: 2 },
      [
        [undefined, undefined, undefined],
        [undefined, undefined, undefined],
        [undefined, undefined, undefined]
      ]
    );

    expect(sheet.occupiedCount()).toBe(0);
    expect(sheet.rowCount()).toBe(3);
    expect(sheet.columnCount()).toBe(3);
  });

  it("filling a trailing blank row does not grow the row axis", () => {
    // 第二行整行为空——导入器把它折进 blankRows，坐标空间仍是 2 行
    const sheet = WorkbookModel.fromSnapshot(
      rowsToWorkbookSnapshot([{ a: 1 }, { a: undefined }], ["a"])
    ).sheet();
    const store = new TableStore(sheet);
    expect(store.getRowCount()).toBe(2);

    store.setCell(1, 0, "filled");

    expect(store.getCell(1, 0)).toBe("filled");
    expect(store.getRowCount()).toBe(2);
  });
});

describe("TableStore does not scan unrelated rows", () => {
  it("setCell touches a constant number of cells regardless of table size", () => {
    // 一百万行是**空白**行：稀疏矩阵一个格子都不建，坐标空间由
    // `blankRows` 直接给出。旧口径靠计数 DataSource 的 getCell/setCell 次数，
    // 这里换成更强的一条：写完之后矩阵里仍然只有那一个格子——既证明没有扫描，
    // 也证明没有为了写一格而把整表物化出来。
    const sheet = SheetModel.createEmpty({ blankRows: 1_000_000, blankColumns: 50 });
    const store = new TableStore(sheet);

    const result = store.setCell(10, 2, 99);

    expect(result.patches).toEqual([{ type: "cell-value", row: 10, column: 2 }]);
    expect(sheet.occupiedCount()).toBe(1);
    expect(sheet.getCell(10, 2)).toBe(99);
    // 值写不改坐标空间：填掉第 10 行这个空白行之后，`blankRows` 相应少 11，轴还是
    // 一百万（`SheetModel.pinCoordinateSpace`）。派生轴的原始算术会给出
    // 10 + 1 + 1_000_000（spec §3.3），那正是这条不变量要挡住的"凭空长一行"。
    expect(store.getRowCount()).toBe(1_000_000);
  });
});

describe("TableStore — structural row operations", () => {
  it("insertRows emits a row-structure Patch and bumps the version", () => {
    const { store } = createStore();

    const result = store.insertRows(1, [["X", 10, ""]]);

    expect(result.patches).toEqual([{ type: "row-structure", operation: "insert", start: 1, count: 1 }]);
    expect(result.version).toBe(1);
    expect(store.getCell(1, 0)).toBe("X");
    expect(store.getCell(2, 0)).toBe("B");
  });

  it("deleteRows emits a row-structure Patch and bumps the version", () => {
    const { store } = createStore();

    const result = store.deleteRows(0, 1);

    expect(result.patches).toEqual([{ type: "row-structure", operation: "delete", start: 0, count: 1 }]);
    expect(result.version).toBe(1);
    expect(store.getCell(0, 0)).toBe("B");
  });

  it("moveRows emits a row-structure Patch (including target) and bumps the version", () => {
    const { store } = createStore();

    const result = store.moveRows(0, 1, 3);

    expect(result.patches).toEqual([
      { type: "row-structure", operation: "move", start: 0, count: 1, target: 3 }
    ]);
    expect(result.version).toBe(1);
  });

  // An asymmetric permutation ([2, 0, 1], not self-inverse) is used so a
  // swapped apply/revert would show up as a wrong row order, not
  // accidentally pass.
  it("reorderRows emits a row-reorder Patch carrying the order and bumps the version", () => {
    const { store } = createStore();

    const result = store.reorderRows(Int32Array.from([2, 0, 1]));

    expect(result.patches).toEqual([{ type: "row-reorder", order: Int32Array.from([2, 0, 1]) }]);
    expect(result.version).toBe(1);
    expect(store.getCell(0, 0)).toBe("C");
    expect(store.getCell(1, 0)).toBe("A");
    expect(store.getCell(2, 0)).toBe("B");
  });

  it.each([
    ["zero-count insert", (store: TableStore) => store.insertRows(1, [])],
    ["delete past the end", (store: TableStore) => store.deleteRows(2, 2)],
    ["move target inside the block", (store: TableStore) => store.moveRows(0, 2, 1)],
    // A wrong-length order against the 3-row createStore() fixture — this is
    // the only TableStore-layer coverage of the assertValidRowOrder() call
    // in createReorderRowsMutation. Calling createReorderRowsMutation()
    // directly (not reorderRows()) matters here: reorderRows() would still
    // pass even if TableStore's own check were deleted, because
    // `SheetModel.reorderRows()` independently re-validates `order` when the
    // mutation is applied and would mask the gap. Calling
    // createReorderRowsMutation() in isolation, before #execute/prepare()
    // ever runs, is what actually exercises the call site in this diff.
    [
      "reorder with a malformed (wrong-length) order",
      (store: TableStore) => store.createReorderRowsMutation(Int32Array.from([2, 0]))
    ]
  ])("rejects %s before mutating data or version", (_label, run) => {
    const { store } = createStore();

    expect(() => run(store)).toThrow(RangeError);

    expect(store.version).toBe(0);
    expect([store.getCell(0, 0), store.getCell(1, 0), store.getCell(2, 0)]).toEqual(["A", "B", "C"]);
  });
});

describe("TableStore — structural column operations", () => {
  it("insertColumns emits a column-structure Patch and bumps the version", () => {
    const { store } = createStore();

    const result = store.insertColumns(1, 1);

    expect(result.patches).toEqual([{ type: "column-structure", operation: "insert", start: 1, count: 1 }]);
    expect(result.version).toBe(1);
    expect(store.getColumnCount()).toBe(4);
    expect(store.getCell(0, 0)).toBe("A"); // name unaffected
    expect(store.getCell(0, 2)).toBe(1); // price shifted right
  });

  it("deleteColumns emits a column-structure Patch and bumps the version", () => {
    const { store } = createStore();

    const result = store.deleteColumns(0, 1);

    expect(result.patches).toEqual([{ type: "column-structure", operation: "delete", start: 0, count: 1 }]);
    expect(result.version).toBe(1);
    expect(store.getColumnCount()).toBe(2);
    expect(store.getCell(0, 0)).toBe(1); // price shifted into name's old slot
  });

  it("moveColumns emits a column-structure Patch (including target) and bumps the version", () => {
    const { store } = createStore();

    const result = store.moveColumns(0, 1, 3);

    expect(result.patches).toEqual([
      { type: "column-structure", operation: "move", start: 0, count: 1, target: 3 }
    ]);
    expect(result.version).toBe(1);
  });

  it.each([
    ["zero-count insert", (store: TableStore) => store.insertColumns(1, 0)],
    ["delete past the end", (store: TableStore) => store.deleteColumns(2, 2)],
    ["move target inside the block", (store: TableStore) => store.moveColumns(0, 2, 1)]
  ])("rejects %s before mutating columns or version", (_label, run) => {
    const { store } = createStore();

    expect(() => run(store)).toThrow(RangeError);

    expect(store.version).toBe(0);
    expect(store.getColumnCount()).toBe(3);
    expect(store.getCell(0, 0)).toBe("A");
    expect(store.getCell(0, 1)).toBe(1);
    expect(store.getCell(0, 2)).toBe("x");
  });
});

/**
 * 归属校验：一个 `TableStore` 一次只服务一张表。旧口径这里问的是「这个
 * DataSource 声明了 beginTransaction 吗」——数据层不再可替换之后那个问题不
 * 存在了，但「拿着 A 表的 mutation 去 B 表的事务里施加」这个错误还在，而且
 * 更危险：它会安静地写错表。
 */
describe("TableStore mutation ownership", () => {
  it("refuses a mutation built against another sheet", () => {
    const { store } = createStore();
    const other = createStore();
    const foreign = other.store.createSetCellMutation(0, 0, "Z");

    const session = store.beginMutationSession("t1");
    try {
      session.stage(foreign);
      expect.unreachable("expected stage to throw");
    } catch (error) {
      expect(error).toBeInstanceOf(CyExcelError);
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.MutationConflict);
    }

    expect(store.version).toBe(0);
    expect(other.store.getCell(0, 0)).toBe("A");
  });
});

describe("TableStore.occupiedCells", () => {
  it("delegates to the sheet's occupied-cell walk", () => {
    const { store } = createStore();

    const seen = [...store.occupiedCells()].sort((left, right) => left.row - right.row || left.column - right.column);
    expect(seen).toEqual([
      { row: 0, column: 0, value: "A" },
      { row: 0, column: 1, value: 1 },
      { row: 0, column: 2, value: "x" },
      { row: 1, column: 0, value: "B" },
      { row: 1, column: 1, value: 2 },
      { row: 1, column: 2, value: "y" },
      { row: 2, column: 0, value: "C" },
      { row: 2, column: 1, value: 3 },
      { row: 2, column: 2, value: "z" }
    ]);
  });
});
