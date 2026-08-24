import { describe, expect, it } from "vitest";

import { MapCellMatrix } from "../../spike/adr-0004/MapCellMatrix";
import { ObjectCellMatrix } from "../../spike/adr-0004/ObjectCellMatrix";
import type { CellMatrix } from "../../spike/adr-0004/types";

const CANDIDATES: ReadonlyArray<readonly [string, new (rows: number, columns: number) => CellMatrix]> = [
  ["MapCellMatrix", MapCellMatrix],
  ["ObjectCellMatrix", ObjectCellMatrix]
];

describe.each(CANDIDATES)("%s", (_name, Matrix) => {
  it("returns undefined for an unoccupied cell instead of throwing", () => {
    const matrix = new Matrix(100, 10);
    expect(matrix.get(50, 5)).toBeUndefined();
    expect(matrix.occupiedCount()).toBe(0);
  });

  it("stores and reads back a record", () => {
    const matrix = new Matrix(100, 10);
    matrix.set(3, 2, { v: 42, t: 2 });
    expect(matrix.get(3, 2)).toEqual({ v: 42, t: 2 });
    expect(matrix.occupiedCount()).toBe(1);
  });

  it("stays sparse — a high row index costs one occupied cell, not a million", () => {
    const matrix = new Matrix(1_048_576, 50);
    matrix.set(1_000_000, 49, { v: "far" });
    expect(matrix.get(1_000_000, 49)).toEqual({ v: "far" });
    expect(matrix.occupiedCount()).toBe(1);
  });

  it("delete removes the cell and drops the occupied count", () => {
    const matrix = new Matrix(10, 10);
    matrix.set(1, 1, { v: "x" });
    matrix.delete(1, 1);
    expect(matrix.get(1, 1)).toBeUndefined();
    expect(matrix.occupiedCount()).toBe(0);
  });

  it("reorderRows moves whole rows with every attribute attached", () => {
    const matrix = new Matrix(3, 2);
    matrix.set(0, 0, { v: "a0", s: "style-a" });
    matrix.set(0, 1, { v: "a1" });
    matrix.set(2, 0, { v: "c0", s: "style-c" });
    // order[new] = old  =>  new row 0 takes old row 2, new row 2 takes old row 0
    matrix.reorderRows(Int32Array.from([2, 1, 0]));
    expect(matrix.get(0, 0)).toEqual({ v: "c0", s: "style-c" });
    expect(matrix.get(2, 0)).toEqual({ v: "a0", s: "style-a" });
    expect(matrix.get(2, 1)).toEqual({ v: "a1" });
    expect(matrix.get(1, 0)).toBeUndefined();
    expect(matrix.occupiedCount()).toBe(3);
  });

  it("reorderRows keeps a row that maps to itself", () => {
    const matrix = new Matrix(2, 1);
    matrix.set(0, 0, { v: "keep" });
    matrix.reorderRows(Int32Array.from([0, 1]));
    expect(matrix.get(0, 0)).toEqual({ v: "keep" });
  });
});
