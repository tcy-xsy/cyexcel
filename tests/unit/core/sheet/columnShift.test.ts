import { describe, expect, it } from "vitest";

import {
  CellMatrix,
  SheetModel,
  buildDeleteColumnsShift,
  buildInsertColumnsShift,
  buildMoveColumnsShift
} from "../../../../src/core/sheet";
import type { ColumnShift } from "../../../../src/core/sheet";

/** 把一段连续列号喂给移位函数，`undefined` 表示该列被删掉。 */
function shiftRange(shift: ColumnShift, upTo: number): ReadonlyArray<number | undefined> {
  return Array.from({ length: upTo }, (_unused, column) => shift(column));
}

describe("buildInsertColumnsShift", () => {
  it("shifts columns at or after the insertion point", () => {
    expect(shiftRange(buildInsertColumnsShift(2, 3), 5)).toEqual([0, 1, 5, 6, 7]);
  });
});

describe("buildDeleteColumnsShift", () => {
  it("drops the deleted band and closes the gap", () => {
    expect(shiftRange(buildDeleteColumnsShift(1, 2), 5)).toEqual([
      0,
      undefined,
      undefined,
      1,
      2
    ]);
  });

  it("is identity when the whole band sits past every occupied column", () => {
    // 列轴对"占用范围外丢数据"天然免疫：这是逐个占用列过的纯算术，不以 extent
    // 为界，所以行轴那一类收窄在列轴上根本不需要。
    expect(shiftRange(buildDeleteColumnsShift(10, 2), 5)).toEqual([0, 1, 2, 3, 4]);
  });
});

describe("buildMoveColumnsShift", () => {
  it("carries the block to the target and shifts the rest around it", () => {
    // target 与行轴同一口径：落在**原**下标第几个之前。
    // A B C D E，把 A、B 移到原第 4 列（E）之前 => C D A B E
    expect(shiftRange(buildMoveColumnsShift(0, 2, 4), 5)).toEqual([2, 3, 0, 1, 4]);
  });

  it("moves a block upward", () => {
    // A B C D E，把 D、E 移到原第 1 列（B）之前 => A D E B C
    expect(shiftRange(buildMoveColumnsShift(3, 2, 1), 5)).toEqual([0, 3, 4, 1, 2]);
  });

  it("is identity when the target equals the start", () => {
    expect(shiftRange(buildMoveColumnsShift(1, 2, 1), 4)).toEqual([0, 1, 2, 3]);
  });
});

describe("the column shift primitives are the single source of the arithmetic", () => {
  interface Case {
    readonly name: string;
    readonly shift: ColumnShift;
    readonly onMatrix: (matrix: CellMatrix) => void;
    readonly onModel: (model: SheetModel) => void;
  }

  // 这段算术过去在 CellMatrix（重键格子）与 SheetModel（重键 columnData）里
  // 各写了一遍，三个操作都是。这条用例把两边都钉在同一份原语上——spec §7.2
  // 里按列索引的插件资源要用的也是它，不能再手抄第三遍。
  const cases: readonly Case[] = [
    {
      name: "insert",
      shift: buildInsertColumnsShift(1, 2),
      onMatrix: (matrix) => matrix.insertColumns(1, 2),
      onModel: (model) => model.insertColumns(1, 2)
    },
    {
      name: "delete",
      shift: buildDeleteColumnsShift(1, 2),
      onMatrix: (matrix) => matrix.deleteColumns(1, 2),
      onModel: (model) => model.deleteColumns(1, 2)
    },
    {
      name: "move",
      shift: buildMoveColumnsShift(0, 2, 4),
      onMatrix: (matrix) => matrix.moveColumns(0, 2, 4),
      onModel: (model) => model.moveColumns(0, 2, 4)
    }
  ];

  it.each(cases)("CellMatrix and SheetModel agree with $name column for column", (testCase) => {
    const matrix = new CellMatrix();
    const model = SheetModel.createEmpty();
    for (let column = 0; column < 5; column += 1) {
      matrix.set(0, column, { v: `c${column}` });
      model.setCell(0, column, `c${column}`);
      model.setColumnMeta(column, { w: column });
    }
    testCase.onMatrix(matrix);
    testCase.onModel(model);
    for (let column = 0; column < 5; column += 1) {
      const target = testCase.shift(column);
      if (target === undefined) {
        continue;
      }
      expect(matrix.get(0, target), `matrix column ${column}`).toEqual({ v: `c${column}` });
      expect(model.getCell(0, target), `model cell ${column}`).toBe(`c${column}`);
      expect(model.getColumnMeta(target), `model meta ${column}`).toEqual({ w: column });
    }
    // 被删掉的列不许在别处冒出来
    const survivors = [0, 1, 2, 3, 4].filter((column) => testCase.shift(column) !== undefined);
    expect(matrix.occupiedCount()).toBe(survivors.length);
    expect(model.occupiedCount()).toBe(survivors.length);
  });
});
