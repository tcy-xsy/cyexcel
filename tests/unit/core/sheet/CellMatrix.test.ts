import { describe, expect, it } from "vitest";

import { CellMatrix } from "../../../../src/core/sheet";

describe("CellMatrix", () => {
  it("returns undefined for an unoccupied cell instead of throwing", () => {
    const matrix = new CellMatrix();
    expect(matrix.get(50, 5)).toBeUndefined();
    expect(matrix.occupiedCount()).toBe(0);
  });

  it("stores and reads back a record", () => {
    const matrix = new CellMatrix();
    matrix.set(3, 2, { v: 42, t: "number" });
    expect(matrix.get(3, 2)).toEqual({ v: 42, t: "number" });
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
});
