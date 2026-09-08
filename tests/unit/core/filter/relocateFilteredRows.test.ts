import { describe, expect, it } from "vitest";

import {
  relocateFilteredRowsForDelete,
  relocateFilteredRowsForInsert,
  relocateFilteredRowsForMove
} from "../../../../src/core/filter/relocateFilteredRows";

describe("relocateFilteredRowsForInsert", () => {
  it("shifts rows at or after the insertion point, leaves earlier rows untouched", () => {
    const result = relocateFilteredRowsForInsert(new Set([1, 3, 5]), 3, 2);
    expect([...result].sort((a, b) => a - b)).toEqual([1, 5, 7]);
  });

  it("returns the same instance when nothing is filtered", () => {
    const filteredRows = new Set<number>();
    expect(relocateFilteredRowsForInsert(filteredRows, 0, 5)).toBe(filteredRows);
  });

  it("returns the same instance when count is 0", () => {
    const filteredRows = new Set([1, 2]);
    expect(relocateFilteredRowsForInsert(filteredRows, 0, 0)).toBe(filteredRows);
  });

  it("returns the same instance when the insertion point is after every filtered row", () => {
    const filteredRows = new Set([1, 2]);
    expect(relocateFilteredRowsForInsert(filteredRows, 10, 3)).toBe(filteredRows);
  });
});

describe("relocateFilteredRowsForDelete", () => {
  it("drops rows inside the deleted range and shifts rows after it down", () => {
    // filtered: 1 (before, untouched), 3 (inside [3,5), dropped), 6 (after, shifts to 4)
    const result = relocateFilteredRowsForDelete(new Set([1, 3, 6]), 3, 2);
    expect([...result].sort((a, b) => a - b)).toEqual([1, 4]);
  });

  it("returns the same instance when nothing is filtered", () => {
    const filteredRows = new Set<number>();
    expect(relocateFilteredRowsForDelete(filteredRows, 0, 5)).toBe(filteredRows);
  });

  it("returns the same instance when the deleted range is entirely before every filtered row and count is 0", () => {
    const filteredRows = new Set([1, 2]);
    expect(relocateFilteredRowsForDelete(filteredRows, 0, 0)).toBe(filteredRows);
  });
});

describe("relocateFilteredRowsForMove", () => {
  it("moves a filtered row inside the moved block to its new position", () => {
    // move rows [2,5) (length 3) to target 10; row 3 is inside the block, offset 1 -> lands at adjustedTarget+1.
    // target(10) > start(2) so adjustedTarget = 10 - 3 = 7; row 3 -> 7 + (3-2) = 8.
    const result = relocateFilteredRowsForMove(new Set([3]), 2, 3, 10);
    expect([...result]).toEqual([8]);
  });

  it("shifts a filtered row that sits between the vacated gap and the insertion point down by count", () => {
    // move rows [2,5) to target 10. Row 6 was after the moved block (position-after-removal = 6-3=3),
    // adjustedTarget=7, 3 < 7 so it stays at 3 (unchanged) -- it's still before the insertion point.
    const result = relocateFilteredRowsForMove(new Set([6]), 2, 3, 10);
    expect([...result]).toEqual([3]);
  });

  it("leaves a filtered row before the moved block and before the target untouched", () => {
    const result = relocateFilteredRowsForMove(new Set([0]), 2, 3, 10);
    expect([...result]).toEqual([0]);
  });

  it("matches a dense-array oracle across many pseudo-random moves", () => {
    let seed = 0x9a1f;
    const random = (limit: number): number => {
      seed = (seed * 1664525 + 1013904223) >>> 0;
      return limit === 0 ? 0 : seed % limit;
    };
    const rowCount = 30;
    // oracle[position] = the original row identity currently sitting there.
    const oracle = Array.from({ length: rowCount }, (_, index) => index);
    const originalFilteredRows = [2, 5, 11, 19, 27];
    let filteredRows: ReadonlySet<number> = new Set(originalFilteredRows);

    for (let step = 0; step < 40; step += 1) {
      const start = random(rowCount - 1);
      const count = Math.min(random(3) + 1, rowCount - start);
      let target = random(rowCount + 1);
      if (target > start && target < start + count) {
        target = start;
      }

      const moved = oracle.splice(start, count);
      const adjustedTarget = target > start ? target - count : target;
      oracle.splice(adjustedTarget, 0, ...moved);

      filteredRows = relocateFilteredRowsForMove(filteredRows, start, count, target);

      const expectedPositions = originalFilteredRows
        .map((row) => oracle.indexOf(row))
        .sort((a, b) => a - b);
      expect([...filteredRows].sort((a, b) => a - b)).toEqual(expectedPositions);
    }
  });
});
