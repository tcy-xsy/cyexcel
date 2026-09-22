import { describe, expect, it } from "vitest";
import {
  matrixToWorkbookSnapshot,
  rowValuesToCells,
  rowsToWorkbookSnapshot,
  valueToRecord
} from "../../../../src/core/sheet/import";
import { WorkbookModel } from "../../../../src/core/sheet/WorkbookModel";

const FIELDS = ["sku", "name", "price"] as const;

function sheetOf(snapshot: ReturnType<typeof rowsToWorkbookSnapshot>) {
  return WorkbookModel.fromSnapshot(snapshot).sheet();
}

describe("valueToRecord — 只有 `undefined` 表示没有这个格子", () => {
  it("普通值包成一条只带 `v` 的记录", () => {
    expect(valueToRecord("x")).toStrictEqual({ v: "x" });
    expect(valueToRecord(0)).toStrictEqual({ v: 0 });
    expect(valueToRecord(false)).toStrictEqual({ v: false });
    expect(valueToRecord("")).toStrictEqual({ v: "" });
  });

  it("`undefined` 与 `null` 都不产生记录", () => {
    // `v` 收紧成 string | number | boolean（对齐 xlsx/Univer）之后 `null` 没地方
    // 放了，而"空格子 = 没有记录"本来就是 xlsx 的表示法。导入与写入两条路仍然同
    // 口径，只是那个共同答案从"存 { v: null }"变成了"不存"。
    expect(valueToRecord(undefined)).toBeUndefined();
    expect(valueToRecord(null)).toBeUndefined();
  });
});

describe("rowValuesToCells — 按列号定位的稀疏一行", () => {
  it("下标就是列号，空洞不占键", () => {
    expect(rowValuesToCells(["a", undefined, "c"])).toStrictEqual({ 0: { v: "a" }, 2: { v: "c" } });
  });

  it("全空的一行是一个空对象，不是三条空记录", () => {
    // 稀疏性靠这一条：存三条 `{ v: undefined }` 会让这一行进入占用范围。
    expect(rowValuesToCells([undefined, undefined, undefined])).toStrictEqual({});
  });

  it("`null` 不占格子，和 `undefined` 同一个答案", () => {
    // `CellValue` 里没有 `null`（对齐 xlsx/Univer），"空格子"只有"没有记录"一种表示。
    expect(rowValuesToCells([null])).toStrictEqual({});
  });
});

describe("rowsToWorkbookSnapshot — 字段按位置落到列上", () => {
  it("第 i 个字段就是第 i 列", () => {
    const sheet = sheetOf(
      rowsToWorkbookSnapshot([{ sku: "s1", name: "n1", price: 10 }], [...FIELDS])
    );
    expect(sheet.getCell(0, 0)).toBe("s1");
    expect(sheet.getCell(0, 1)).toBe("n1");
    expect(sheet.getCell(0, 2)).toBe(10);
  });

  it("`fields` 里没有的字段不进表", () => {
    // 反空洞：如果实现是"把行对象的每个键都铺开"，这一条会红。
    const sheet = sheetOf(rowsToWorkbookSnapshot([{ sku: "s1", secret: "x" }], ["sku"]));
    expect(sheet.columnCount()).toBe(1);
    expect(sheet.getCell(0, 0)).toBe("s1");
  });

  it("行里缺的字段不产生格子，但它右边的字段仍落在自己那一列", () => {
    const sheet = sheetOf(rowsToWorkbookSnapshot([{ price: 10 }], [...FIELDS]));
    expect(sheet.getRecord(0, 0)).toBeUndefined();
    expect(sheet.getRecord(0, 1)).toBeUndefined();
    expect(sheet.getCell(0, 2)).toBe(10);
  });
});

describe("rowsToWorkbookSnapshot — 坐标空间恰好是 rows.length × fields.length", () => {
  it("常规数据：行列数与输入一致", () => {
    const sheet = sheetOf(
      rowsToWorkbookSnapshot(
        [
          { sku: "s1", name: "n1", price: 1 },
          { sku: "s2", name: "n2", price: 2 }
        ],
        [...FIELDS]
      )
    );
    expect(sheet.rowCount()).toBe(2);
    expect(sheet.columnCount()).toBe(3);
  });

  it("末尾整行为空时行数不缩水", () => {
    // 这是稀疏模型下最容易掉的一格：全空的尾行不占矩阵位置，直接
    // `createEmpty({ blankRows: 0 })` 会得到一张 1 行的表，导入 3 行少 2 行。
    const sheet = sheetOf(
      rowsToWorkbookSnapshot(
        [{ sku: "s1" }, { sku: undefined }, {}],
        [...FIELDS]
      )
    );
    expect(sheet.rowCount()).toBe(3);
    expect(sheet.columnCount()).toBe(3);
  });

  it("一行都没有时是一张空表，不是一张负数行的表", () => {
    const sheet = sheetOf(rowsToWorkbookSnapshot([], [...FIELDS]));
    expect(sheet.rowCount()).toBe(0);
    expect(sheet.columnCount()).toBe(3);
  });

  it("`blankRows` / `blankColumns` 是**额外**的空白，加在 rows.length 之上", () => {
    const sheet = sheetOf(
      rowsToWorkbookSnapshot([{ sku: "s1" }, {}], [...FIELDS], { blankRows: 10, blankColumns: 2 })
    );
    expect(sheet.rowCount()).toBe(12);
    expect(sheet.columnCount()).toBe(5);
  });

  it("其余 sheet 选项原样透传", () => {
    const snapshot = rowsToWorkbookSnapshot([{ sku: "s1" }], ["sku"], {
      id: "s",
      name: "导入的表",
      defaultRowHeight: 33
    });
    expect(snapshot.sheetOrder).toStrictEqual(["s"]);
    expect(snapshot.sheets["s"]?.name).toBe("导入的表");
    expect(snapshot.sheets["s"]?.defaultRowHeight).toBe(33);
  });
});

describe("rowsToWorkbookSnapshot — 产出的是一份合法快照", () => {
  it("能被 `WorkbookModel.fromSnapshot` 吃下去，且 styles / resources 是空的", () => {
    // `fromSnapshot` 头一件事就是 `assertValidWorkbookSnapshot`，所以这一条同时
    // 钉住了 version / sheetOrder / sheets 三者自洽。
    const snapshot = rowsToWorkbookSnapshot([{ sku: "s1" }], ["sku"]);
    expect(() => WorkbookModel.fromSnapshot(snapshot)).not.toThrow();
    expect(snapshot.styles).toStrictEqual({});
    expect(snapshot.resources).toStrictEqual([]);
  });

  it("快照与调用方手里的行对象再无联系（写回消失，设计稿 §3.3）", () => {
    const rows = [{ sku: "s1" }];
    const sheet = sheetOf(rowsToWorkbookSnapshot(rows, ["sku"]));
    sheet.setCell(0, 0, "改过了");
    expect(rows[0]?.sku).toBe("s1");
  });
});

describe("matrixToWorkbookSnapshot — Excel 式二维数组", () => {
  it("row/column 下标直接落到单元格坐标", () => {
    const sheet = sheetOf(matrixToWorkbookSnapshot([
      ["A1", "B1"],
      ["A2", "B2"]
    ]));

    expect(sheet.getCell(0, 0)).toBe("A1");
    expect(sheet.getCell(0, 1)).toBe("B1");
    expect(sheet.getCell(1, 0)).toBe("A2");
    expect(sheet.getCell(1, 1)).toBe("B2");
  });

  it("保留空行、空列和锯齿数组的位置", () => {
    const sheet = sheetOf(matrixToWorkbookSnapshot([
      ["A1", undefined, "C1"],
      [],
      [null, "B3"]
    ]));

    expect(sheet.rowCount()).toBe(3);
    expect(sheet.columnCount()).toBe(3);
    expect(sheet.getRecord(1, 0)).toBeUndefined();
    expect(sheet.getCell(2, 0)).toBeUndefined();
    expect(sheet.getCell(2, 1)).toBe("B3");
  });

  it("`0` / `\"\"` / `false` 是值，`null` / `undefined` 是空格", () => {
    const sheet = sheetOf(matrixToWorkbookSnapshot([[0, "", false, null, undefined]]));

    expect(sheet.getCell(0, 0)).toBe(0);
    expect(sheet.getCell(0, 1)).toBe("");
    expect(sheet.getCell(0, 2)).toBe(false);
    expect(sheet.getCell(0, 3)).toBeUndefined();
    expect(sheet.getCell(0, 4)).toBeUndefined();
  });

  it("额外 blank 保持工作区大小，sheet 选项原样透传", () => {
    const snapshot = matrixToWorkbookSnapshot([["A1"]], {
      blankRows: 9,
      blankColumns: 2,
      id: "s",
      name: "Matrix"
    });
    const sheet = sheetOf(snapshot);

    expect(sheet.rowCount()).toBe(10);
    expect(sheet.columnCount()).toBe(3);
    expect(snapshot.sheets["s"]?.name).toBe("Matrix");
  });
});
