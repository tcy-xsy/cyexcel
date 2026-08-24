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

  it("dedupes row-reorder patches by their order sequence", () => {
    const patches: Patch[] = [
      { type: "row-reorder", order: Int32Array.from([1, 0, 2]) },
      { type: "row-reorder", order: Int32Array.from([1, 0, 2]) },
      { type: "row-reorder", order: Int32Array.from([2, 0, 1]) }
    ];

    expect(mergePatches(patches)).toEqual([
      { type: "row-reorder", order: Int32Array.from([1, 0, 2]) },
      { type: "row-reorder", order: Int32Array.from([2, 0, 1]) }
    ]);
  });

  it("dedupes row-reorder patches without ever stringifying a permutation", () => {
    // The two patches one sort commit produces (TableStore's and the layout
    // participant's) are distinct Int32Arrays with equal content —
    // `createChangeSet` slices each order, so identity comparison would stop
    // them merging. Keying them by content is what this guards against: a
    // 1M-element `join(",")` costs ~190ms and ~6.9M chars *per patch*, on
    // every commit and again on every undo/redo, for a merged array
    // `#publishCommitReceipt` then discards.
    class ObservedOrder extends Int32Array {
      joins = 0;
      override join(separator?: string): string {
        this.joins += 1;
        return super.join(separator);
      }
    }
    const first = ObservedOrder.from([1, 0, 2]) as ObservedOrder;
    const second = ObservedOrder.from([1, 0, 2]) as ObservedOrder;
    const third = ObservedOrder.from([2, 0, 1]) as ObservedOrder;

    const merged = mergePatches([
      { type: "row-reorder", order: first },
      { type: "row-reorder", order: second },
      { type: "row-reorder", order: third }
    ]);

    expect(
      merged.map((patch) => (patch.type === "row-reorder" ? [...patch.order] : patch.type))
    ).toEqual([[1, 0, 2], [2, 0, 1]]);
    expect([first.joins, second.joins, third.joins]).toEqual([0, 0, 0]);
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
