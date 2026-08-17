import { describe, expect, it } from "vitest";

import { mergePatches } from "../../../../src/core/data";
import type { Patch } from "../../../../src/core/types";

describe("mergePatches", () => {
  it("dedupes structurally identical patches, keeping the first occurrence", () => {
    const patches: Patch[] = [
      { type: "cell-value", row: 1, column: 2 },
      { type: "cell-value", row: 1, column: 2 },
      { type: "cell-value", row: 1, column: 3 }
    ];

    expect(mergePatches(patches)).toEqual([
      { type: "cell-value", row: 1, column: 2 },
      { type: "cell-value", row: 1, column: 3 }
    ]);
  });

  it("keeps patches of different types or coordinates independent", () => {
    const patches: Patch[] = [
      { type: "cell-value", row: 0, column: 0 },
      { type: "cell-style", row: 0, column: 0 },
      {
        type: "range-value",
        range: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 }
      },
      { type: "full-reset" },
      { type: "full-reset" }
    ];

    const merged = mergePatches(patches);

    expect(merged).toHaveLength(4);
    expect(merged).toEqual([
      { type: "cell-value", row: 0, column: 0 },
      { type: "cell-style", row: 0, column: 0 },
      {
        type: "range-value",
        range: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 }
      },
      { type: "full-reset" }
    ]);
  });

  it("dedupes row-structure patches by operation, start, count, and target", () => {
    const patches: Patch[] = [
      { type: "row-structure", operation: "insert", start: 0, count: 1 },
      { type: "row-structure", operation: "insert", start: 0, count: 1 },
      { type: "row-structure", operation: "move", start: 0, count: 1, target: 5 },
      { type: "row-structure", operation: "move", start: 0, count: 1, target: 5 }
    ];

    expect(mergePatches(patches)).toEqual([
      { type: "row-structure", operation: "insert", start: 0, count: 1 },
      { type: "row-structure", operation: "move", start: 0, count: 1, target: 5 }
    ]);
  });

  it("dedupes column-layout patches by column index", () => {
    const patches: Patch[] = [
      { type: "column-layout", column: 2 },
      { type: "column-layout", column: 2 },
      { type: "column-layout", column: 3 }
    ];

    expect(mergePatches(patches)).toEqual([
      { type: "column-layout", column: 2 },
      { type: "column-layout", column: 3 }
    ]);
  });

  it("dedupes row-layout batches by their complete row set", () => {
    const patches: Patch[] = [
      {
        type: "row-layout",
        rows: [1, 4],
        range: { startRow: 1, startColumn: 0, endRow: 9, endColumn: 2 }
      },
      {
        type: "row-layout",
        rows: [1, 4],
        range: { startRow: 1, startColumn: 0, endRow: 9, endColumn: 2 }
      },
      {
        type: "row-layout",
        rows: [1, 5],
        range: { startRow: 1, startColumn: 0, endRow: 9, endColumn: 2 }
      },
      {
        type: "row-layout",
        rows: [4, 1],
        range: { startRow: 1, startColumn: 0, endRow: 9, endColumn: 2 }
      },
      {
        type: "row-layout",
        rows: [1, 4],
        range: { startRow: 1, startColumn: 0, endRow: 10, endColumn: 2 }
      }
    ];

    expect(mergePatches(patches)).toEqual([patches[0], patches[2], patches[3], patches[4]]);
  });
});
