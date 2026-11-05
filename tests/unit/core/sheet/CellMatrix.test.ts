import { describe, expect, it, vi } from "vitest";

import { CellMatrix } from "../../../../src/core/sheet";
import { CellValueType } from "../../../../src/core/sheet/types";

describe("CellMatrix", () => {
  it("tracks formula counts independently across copy-on-write snapshots", () => {
    const matrix = new CellMatrix();
    matrix.set(0, 0, { v: undefined, f: '=1' });
    matrix.set(2, 2, { v: undefined, f: '=2' });
    const copy = matrix.clone();
    matrix.set(0, 0, { v: 1 });
    matrix.delete(2, 2);
    expect(matrix.formulaCount()).toBe(0);
    expect(copy.formulaCount()).toBe(2);
    copy.deleteColumns(0, 1);
    expect(copy.formulaCount()).toBe(1);
    copy.applyRowPlan(new Int32Array());
    expect(copy.formulaCount()).toBe(0);
  });
  it("clones without enumerating rows and detaches writes across multiple snapshots", () => {
    const matrix = new CellMatrix();
    for (const row of [0, 1023, 1024, 100_000, 1_048_576]) matrix.set(row, 0, { v: row });
    const keys = vi.spyOn(Object, "keys");
    let first: CellMatrix;
    let second: CellMatrix;
    try {
      first = matrix.clone();
      second = first.clone();
      expect(keys).not.toHaveBeenCalled();
    } finally {
      keys.mockRestore();
    }
    matrix.set(1024, 0, { v: "live" });
    first.delete(100_000, 0);
    second.set(1_048_576, 0, { v: "second" });
    expect(first.get(1024, 0)?.v).toBe(1024);
    expect(second.get(100_000, 0)?.v).toBe(100_000);
    expect(matrix.get(1_048_576, 0)?.v).toBe(1_048_576);
    expect([...first.rows()].map(([row]) => row)).toEqual([0, 1023, 1024, 1_048_576]);
  });

  it("clears the rightmost column without scanning all remaining row buckets per cell", () => {
    const matrix = new CellMatrix();
    for (let row = 0; row < 2_000; row += 1) matrix.set(row, 4, { v: row });
    const keys = vi.spyOn(Object, "keys");
    try {
      for (let row = 0; row < 2_000; row += 1) matrix.delete(row, 4);
      expect(matrix.maxOccupiedColumn()).toBeUndefined();
      expect(matrix.maxOccupiedRow()).toBeUndefined();
      expect(keys.mock.calls.length).toBeLessThan(6_000);
    } finally {
      keys.mockRestore();
    }
  });

  it("refreshes dirty boundaries correctly after adding cells and cloning", () => {
    const matrix = new CellMatrix();
    matrix.set(10, 10, { v: 1 });
    matrix.set(2, 2, { v: 2 });
    matrix.delete(10, 10);
    const copy = matrix.clone();
    matrix.set(4, 4, { v: 4 });
    expect(matrix.maxOccupiedRow()).toBe(4);
    expect(matrix.maxOccupiedColumn()).toBe(4);
    expect(copy.maxOccupiedRow()).toBe(2);
    expect(copy.maxOccupiedColumn()).toBe(2);
  });
  it("returns undefined for an unoccupied cell instead of throwing", () => {
    const matrix = new CellMatrix();
    expect(matrix.get(50, 5)).toBeUndefined();
    expect(matrix.occupiedCount()).toBe(0);
  });

  it("stores and reads back a record", () => {
    const matrix = new CellMatrix();
    matrix.set(3, 2, { v: 42, t: CellValueType.NUMBER });
    expect(matrix.get(3, 2)).toEqual({ v: 42, t: CellValueType.NUMBER });
    expect(matrix.occupiedCount()).toBe(1);
  });

  it("stays sparse — a high index costs one occupied cell", () => {
    const matrix = new CellMatrix();
    matrix.set(1_000_000, 49, { v: "far" });
    expect(matrix.occupiedCount()).toBe(1);
    expect(matrix.maxOccupiedRow()).toBe(1_000_000);
    expect(matrix.maxOccupiedColumn()).toBe(49);
  });

  it("reports undefined maxOccupied when empty", () => {
    const matrix = new CellMatrix();
    expect(matrix.maxOccupiedRow()).toBeUndefined();
    expect(matrix.maxOccupiedColumn()).toBeUndefined();
  });

  it("tracks maxOccupiedRow as cells are added", () => {
    const matrix = new CellMatrix();
    matrix.set(5, 0, { v: "a" });
    expect(matrix.maxOccupiedRow()).toBe(5);
    matrix.set(2, 0, { v: "b" });
    expect(matrix.maxOccupiedRow()).toBe(5); // 更小的下标不改变最大值
    matrix.set(9, 0, { v: "c" });
    expect(matrix.maxOccupiedRow()).toBe(9);
  });

  it("walks maxOccupiedRow back to the next-largest on deleting the largest", () => {
    const matrix = new CellMatrix();
    matrix.set(2, 0, { v: "a" });
    matrix.set(7, 0, { v: "b" });
    matrix.delete(7, 0);
    expect(matrix.maxOccupiedRow()).toBe(2);
  });

  it("keeps maxOccupiedRow when deleting a non-largest row", () => {
    const matrix = new CellMatrix();
    matrix.set(2, 0, { v: "a" });
    matrix.set(7, 0, { v: "b" });
    matrix.delete(2, 0);
    expect(matrix.maxOccupiedRow()).toBe(7);
  });

  it("keeps maxOccupiedRow while the largest row still has other cells", () => {
    const matrix = new CellMatrix();
    matrix.set(7, 0, { v: "a" });
    matrix.set(7, 1, { v: "b" });
    matrix.delete(7, 0);
    expect(matrix.maxOccupiedRow()).toBe(7); // 该行还有别的格子
  });

  it("walks maxOccupiedColumn back across other rows", () => {
    const matrix = new CellMatrix();
    matrix.set(0, 3, { v: "a" });
    matrix.set(5, 9, { v: "b" });
    matrix.delete(5, 9);
    // 第 9 列没了，但第 3 列仍被第 0 行占用
    expect(matrix.maxOccupiedColumn()).toBe(3);
  });

  it("keeps maxOccupiedColumn when another row still occupies it", () => {
    const matrix = new CellMatrix();
    matrix.set(0, 9, { v: "a" });
    matrix.set(5, 9, { v: "b" });
    matrix.delete(5, 9);
    expect(matrix.maxOccupiedColumn()).toBe(9);
  });

  it("returns to undefined maxOccupied after everything is deleted", () => {
    const matrix = new CellMatrix();
    matrix.set(3, 4, { v: "a" });
    matrix.delete(3, 4);
    expect(matrix.maxOccupiedRow()).toBeUndefined();
    expect(matrix.maxOccupiedColumn()).toBeUndefined();
    expect(matrix.occupiedCount()).toBe(0);
  });

  it("prunes the row bucket when its last cell is deleted", () => {
    const matrix = new CellMatrix();
    matrix.set(3, 4, { v: "a" });
    matrix.delete(3, 4);
    expect([...matrix.rows()]).toEqual([]);
  });

  it("deleting an unoccupied cell is a no-op", () => {
    const matrix = new CellMatrix();
    matrix.set(1, 1, { v: "a" });
    matrix.delete(9, 9);
    expect(matrix.occupiedCount()).toBe(1);
    expect(matrix.maxOccupiedRow()).toBe(1);
  });

  it("iterates occupied rows with their buckets", () => {
    const matrix = new CellMatrix();
    matrix.set(2, 0, { v: "a" });
    matrix.set(0, 1, { v: "b" });
    const seen = [...matrix.rows()].map(([row]) => row).sort((a, b) => a - b);
    expect(seen).toEqual([0, 2]);
  });

  it("reorderRows moves whole rows with every field attached", () => {
    const matrix = new CellMatrix();
    matrix.set(0, 0, { v: "a0", s: "s1" });
    matrix.set(0, 1, { v: "a1" });
    matrix.set(2, 0, { v: "c0", s: "s2" });
    // order[newRow] = oldRow（仓库既有约定，ADR 0003）
    matrix.reorderRows(Int32Array.from([2, 1, 0]));
    expect(matrix.get(0, 0)).toEqual({ v: "c0", s: "s2" });
    expect(matrix.get(2, 0)).toEqual({ v: "a0", s: "s1" });
    expect(matrix.get(2, 1)).toEqual({ v: "a1" });
    expect(matrix.get(1, 0)).toBeUndefined();
    expect(matrix.occupiedCount()).toBe(3);
  });

  it("recomputes maxOccupiedRow after a reorder", () => {
    const matrix = new CellMatrix();
    matrix.set(0, 0, { v: "a" });
    matrix.reorderRows(Int32Array.from([1, 0]));
    // 原第 0 行搬到了第 1 行
    expect(matrix.maxOccupiedRow()).toBe(1);
    expect(matrix.get(1, 0)).toEqual({ v: "a" });
  });

  it("rejects non-permutation row order", () => {
    const matrix = new CellMatrix();
    matrix.set(0, 0, { v: "a" });
    const duplicateOrder = Int32Array.from([0, 0]);
    expect(() => matrix.reorderRows(duplicateOrder)).toThrow(RangeError);
  });

  it("rejects row order too short to cover largest occupied row", () => {
    const matrix = new CellMatrix();
    matrix.set(5, 0, { v: "a" });
    const tooShortOrder = Int32Array.from([0, 1, 2]);
    expect(() => matrix.reorderRows(tooShortOrder)).toThrow(RangeError);
  });

  it("skips rows deleted mid-iteration and yields only surviving rows", () => {
    const matrix = new CellMatrix();
    matrix.set(0, 0, { v: "a" });
    matrix.set(1, 0, { v: "b" });
    matrix.set(2, 0, { v: "c" });
    const rows: Array<[number, unknown]> = [];
    for (const [row, bucket] of matrix.rows()) {
      rows.push([row, bucket]);
      if (row === 0) {
        // Delete a not-yet-visited row mid-iteration
        matrix.delete(2, 0);
      }
    }
    const rowNumbers = rows.map(([row]) => row).sort((a, b) => a - b);
    expect(rowNumbers).toEqual([0, 1]);
    // Verify all buckets are valid (not undefined)
    for (const [, bucket] of rows) {
      expect(bucket).toBeDefined();
      expect(typeof bucket).toBe("object");
    }
  });
});
