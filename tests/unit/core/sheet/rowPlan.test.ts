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
