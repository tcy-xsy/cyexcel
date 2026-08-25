import { describe, expect, it } from "vitest";

import {
  ROW_PLAN_HOLE,
  assertValidRowPlan,
  buildDeleteRowsPlan,
  buildInsertRowsPlan,
  buildMoveRowsPlan
} from "../../../../src/core/sheet";
import { CyExcelError, CyExcelErrorCode } from "../../../../src/core/errors";

const H = ROW_PLAN_HOLE;

describe("buildInsertRowsPlan", () => {
  it("opens holes at the insertion point and shifts the rest down", () => {
    // 占用 0..3，在 1 处插 2 行 => 新的 1、2 是空行，旧的 1..3 后移
    expect([...buildInsertRowsPlan(4, 1, 2)]).toEqual([0, H, H, 1, 2, 3]);
  });

  it("appends at the end", () => {
    expect([...buildInsertRowsPlan(2, 2, 1)]).toEqual([0, 1, H]);
  });

  it("inserts into an empty matrix", () => {
    expect([...buildInsertRowsPlan(0, 0, 3)]).toEqual([H, H, H]);
  });

  it("inserts at the very front", () => {
    expect([...buildInsertRowsPlan(2, 0, 1)]).toEqual([H, 0, 1]);
  });

  it("reaches an insertion point that sits in the blank area", () => {
    // 坐标空间比占用范围大（spec §3.3），插入点落在空白区是日常操作：
    // 排列必须一直铺到落点，否则它根本没描述那一段。
    expect([...buildInsertRowsPlan(2, 5, 1)]).toEqual([0, 1, 2, 3, 4, H]);
  });
});

describe("buildDeleteRowsPlan", () => {
  it("omits the deleted rows so they are dropped", () => {
    // 占用 0..4，删 1..2 => 剩下 0,3,4
    expect([...buildDeleteRowsPlan(5, 1, 2)]).toEqual([0, 3, 4]);
  });

  it("deletes a trailing range", () => {
    expect([...buildDeleteRowsPlan(3, 1, 2)]).toEqual([0]);
  });

  it("deletes everything", () => {
    expect([...buildDeleteRowsPlan(3, 0, 3)]).toEqual([]);
  });

  it("only shrinks by the part of the range that lands inside the occupied extent", () => {
    // 删除范围可以整段或部分落在占用范围之外——那是坐标空间里合法的空白区
    // （spec §3.3）。按整个 count 收缩排列会把范围之外的行连带丢掉。
    const cases: ReadonlyArray<readonly [number, number, number, readonly number[]]> = [
      [5, 3, 10, [0, 1, 2]],
      [5, 10, 2, [0, 1, 2, 3, 4]],
      [5, 0, 5, []],
      [5, 2, 2, [0, 1, 4]],
      [5, 4, 1, [0, 1, 2, 3]],
      [0, 0, 3, []]
    ];
    for (const [extent, start, count, expected] of cases) {
      expect([...buildDeleteRowsPlan(extent, start, count)], `E=${extent} ${start}+${count}`).toEqual(
        [...expected]
      );
    }
  });
});

describe("buildMoveRowsPlan", () => {
  it("moves a block downward", () => {
    // target 是"落在原第几行之前"，与仓库既有 moveRows 同一口径
    // （ArrayDataSource.test.ts:207 `moveRows(1,2,4)` 注释 "move B,C to just
    // before E"，结果 A,D,B,C,E）。占用 0..4，把 0..1 移到原第 3 行之前 =>
    // 2 上移到头，0、1 紧随其后，3、4 不动。
    expect([...buildMoveRowsPlan(5, 0, 2, 3)]).toEqual([2, 0, 1, 3, 4]);
  });

  it("moves a block upward", () => {
    expect([...buildMoveRowsPlan(5, 3, 2, 1)]).toEqual([0, 3, 4, 1, 2]);
  });

  it("is identity when the target equals the start", () => {
    expect([...buildMoveRowsPlan(4, 1, 2, 1)]).toEqual([0, 1, 2, 3]);
  });

  it("keeps the plan the same length as the extent", () => {
    expect(buildMoveRowsPlan(6, 1, 2, 4).length).toBe(6);
  });

  it("keeps the blank gap when the block or the landing point reaches past the occupied extent", () => {
    // 只铺到 occupiedExtent 会把空白间隙压掉，被移动的数据落在错误的下标上。
    const cases: ReadonlyArray<
      readonly [number, number, number, number, readonly number[]]
    > = [
      [5, 0, 2, 8, [2, 3, 4, 5, 6, 7, 0, 1]],
      [5, 3, 10, 0, [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 0, 1, 2]],
      [5, 1, 2, 4, [0, 3, 1, 2, 4]]
    ];
    for (const [extent, start, count, target, expected] of cases) {
      expect(
        [...buildMoveRowsPlan(extent, start, count, target)],
        `E=${extent} ${start}+${count}->${target}`
      ).toEqual([...expected]);
    }
  });
});

describe("assertValidRowPlan", () => {
  it("accepts a plan with holes and dropped rows", () => {
    expect(() => assertValidRowPlan(Int32Array.from([2, H, 0]), 3)).not.toThrow();
  });

  it("accepts an empty plan", () => {
    expect(() => assertValidRowPlan(Int32Array.from([]), 0)).not.toThrow();
  });

  it("rejects a source row named twice", () => {
    // 同一行被搬到两个位置，会凭空复制数据
    expect(() => assertValidRowPlan(Int32Array.from([0, 0]), 1)).toThrow(CyExcelError);
  });

  it("rejects a source row outside the occupied extent", () => {
    expect(() => assertValidRowPlan(Int32Array.from([0, 5]), 2)).toThrow(CyExcelError);
  });

  it("rejects a negative entry other than the hole marker", () => {
    expect(() => assertValidRowPlan(Int32Array.from([0, -2]), 2)).toThrow(CyExcelError);
  });

  it("reports InvalidRange for every rejection", () => {
    expect.hasAssertions();
    const bad: ReadonlyArray<readonly [Int32Array, number]> = [
      [Int32Array.from([0, 0]), 1],
      [Int32Array.from([0, 5]), 2],
      [Int32Array.from([0, -2]), 2]
    ];
    for (const [plan, extent] of bad) {
      let threw = false;
      try {
        assertValidRowPlan(plan, extent);
      } catch (error) {
        threw = true;
        expect((error as CyExcelError).code).toBe(CyExcelErrorCode.InvalidRange);
      }
      expect(threw).toBe(true);
    }
  });
});
