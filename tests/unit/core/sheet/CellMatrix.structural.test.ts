import { describe, expect, it } from "vitest";

import { CellMatrix, ROW_PLAN_HOLE } from "../../../../src/core/sheet";
import { CyExcelError } from "../../../../src/core/errors";

const H = ROW_PLAN_HOLE;

function occupiedKeys(matrix: CellMatrix): string[] {
  const keys: string[] = [];
  for (const [row, bucket] of matrix.rows()) {
    for (const column of Object.keys(bucket)) {
      keys.push(`${row}:${column}`);
    }
  }
  return keys.sort();
}

describe("CellMatrix.applyRowPlan", () => {
  it("leaves a hole row genuinely unoccupied", () => {
    const matrix = new CellMatrix();
    matrix.set(0, 0, { v: "a" });
    matrix.applyRowPlan(Int32Array.from([H, 0]));
    expect(matrix.get(0, 0)).toBeUndefined();
    expect(matrix.get(1, 0)).toEqual({ v: "a" });
    expect(matrix.occupiedCount()).toBe(1);
  });

  it("drops rows the plan does not name", () => {
    const matrix = new CellMatrix();
    matrix.set(0, 0, { v: "keep" });
    matrix.set(1, 0, { v: "drop" });
    matrix.applyRowPlan(Int32Array.from([0]));
    expect(matrix.get(0, 0)).toEqual({ v: "keep" });
    expect(matrix.occupiedCount()).toBe(1);
    expect(matrix.maxOccupiedRow()).toBe(0);
  });

  it("recomputes occupiedCount when rows are dropped", () => {
    const matrix = new CellMatrix();
    matrix.set(0, 0, { v: "a" });
    matrix.set(0, 1, { v: "b" });
    matrix.set(1, 0, { v: "c" });
    matrix.applyRowPlan(Int32Array.from([1]));
    expect(matrix.occupiedCount()).toBe(1);
  });

  it("recomputes maxOccupiedColumn when the widest row is dropped", () => {
    const matrix = new CellMatrix();
    matrix.set(0, 2, { v: "narrow" });
    matrix.set(1, 9, { v: "wide" });
    matrix.applyRowPlan(Int32Array.from([0]));
    expect(matrix.maxOccupiedColumn()).toBe(2);
  });

  it("reports undefined maxima when the plan drops everything", () => {
    const matrix = new CellMatrix();
    matrix.set(3, 4, { v: "a" });
    matrix.applyRowPlan(Int32Array.from([]));
    expect(matrix.occupiedCount()).toBe(0);
    expect(matrix.maxOccupiedRow()).toBeUndefined();
    expect(matrix.maxOccupiedColumn()).toBeUndefined();
  });

  it("rejects a plan that names a source row twice", () => {
    const matrix = new CellMatrix();
    matrix.set(0, 0, { v: "a" });
    expect(() => matrix.applyRowPlan(Int32Array.from([0, 0]))).toThrow(CyExcelError);
  });
});

describe("CellMatrix.insertColumns", () => {
  it("shifts columns at and after the insertion point", () => {
    const matrix = new CellMatrix();
    matrix.set(0, 0, { v: "a" });
    matrix.set(0, 1, { v: "b" });
    matrix.insertColumns(1, 2);
    expect(matrix.get(0, 0)).toEqual({ v: "a" });
    expect(matrix.get(0, 1)).toBeUndefined();
    expect(matrix.get(0, 2)).toBeUndefined();
    expect(matrix.get(0, 3)).toEqual({ v: "b" });
    expect(matrix.maxOccupiedColumn()).toBe(3);
  });

  it("touches every occupied row, not just the first", () => {
    const matrix = new CellMatrix();
    matrix.set(0, 0, { v: "a" });
    matrix.set(5, 0, { v: "b" });
    matrix.insertColumns(0, 1);
    expect(occupiedKeys(matrix)).toEqual(["0:1", "5:1"]);
  });

  it("is a no-op on an empty matrix", () => {
    const matrix = new CellMatrix();
    matrix.insertColumns(0, 3);
    expect(matrix.occupiedCount()).toBe(0);
    expect(matrix.maxOccupiedColumn()).toBeUndefined();
  });
});

describe("CellMatrix.deleteColumns", () => {
  it("removes the range and shifts what follows back", () => {
    const matrix = new CellMatrix();
    matrix.set(0, 0, { v: "a" });
    matrix.set(0, 1, { v: "gone" });
    matrix.set(0, 2, { v: "c" });
    matrix.deleteColumns(1, 1);
    expect(matrix.get(0, 0)).toEqual({ v: "a" });
    expect(matrix.get(0, 1)).toEqual({ v: "c" });
    expect(matrix.occupiedCount()).toBe(2);
    expect(matrix.maxOccupiedColumn()).toBe(1);
  });

  it("prunes a row whose every cell was deleted", () => {
    const matrix = new CellMatrix();
    matrix.set(0, 0, { v: "keep" });
    matrix.set(4, 5, { v: "gone" });
    matrix.deleteColumns(5, 1);
    expect(occupiedKeys(matrix)).toEqual(["0:0"]);
    expect(matrix.maxOccupiedRow()).toBe(0);
  });

  it("empties the matrix when every column goes", () => {
    const matrix = new CellMatrix();
    matrix.set(0, 0, { v: "a" });
    matrix.set(1, 1, { v: "b" });
    matrix.deleteColumns(0, 2);
    expect(matrix.occupiedCount()).toBe(0);
    expect(matrix.maxOccupiedRow()).toBeUndefined();
    expect(matrix.maxOccupiedColumn()).toBeUndefined();
  });
});

describe("CellMatrix.moveColumns", () => {
  it("moves a column block to the right", () => {
    const matrix = new CellMatrix();
    matrix.set(0, 0, { v: "a" });
    matrix.set(0, 1, { v: "b" });
    matrix.set(0, 2, { v: "c" });
    // target 是"落在原第几列之前"（同仓库既有 moveRows 口径）：把列 0 移到
    // 原列 2 之前 => b 上移到 0，a 落在 b 之后，c 不动。
    matrix.moveColumns(0, 1, 2);
    expect(matrix.get(0, 0)).toEqual({ v: "b" });
    expect(matrix.get(0, 1)).toEqual({ v: "a" });
    expect(matrix.get(0, 2)).toEqual({ v: "c" });
  });

  it("moves a column block to the left across every row", () => {
    const matrix = new CellMatrix();
    matrix.set(0, 2, { v: "x" });
    matrix.set(7, 2, { v: "y" });
    matrix.moveColumns(2, 1, 0);
    expect(occupiedKeys(matrix)).toEqual(["0:0", "7:0"]);
  });
});
