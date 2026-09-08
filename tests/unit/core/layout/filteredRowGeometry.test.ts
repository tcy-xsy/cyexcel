import { describe, expect, it } from "vitest";

import { LayoutEngine } from "../../../../src/core/layout/LayoutEngine";

function createEngine(rowCount = 5, rowHeight = 20): LayoutEngine {
  return new LayoutEngine({
    rowCount,
    columnCount: 3,
    rowHeight,
    columnWidths: [50, 50, 50]
  });
}

describe("LayoutEngine filtered-row geometry: baseline (Task 3)", () => {
  it("starts with an empty filteredRows set and visibleRowMetrics aliased to rowMetrics", () => {
    const layout = createEngine();
    expect(layout.root.filteredRows.size).toBe(0);
    expect(layout.root.visibleRowMetrics).toBe(layout.root.rowMetrics);
  });

  it("isRowVisible is true for every row by default", () => {
    const layout = createEngine();
    for (let row = 0; row < 5; row += 1) {
      expect(layout.isRowVisible(row)).toBe(true);
    }
  });

  it("prepareSetFilteredRows marks the given rows invisible without touching stored row heights", () => {
    const layout = createEngine();
    layout.replaceRoot(layout.prepareSetFilteredRows(new Set([1, 3])));

    expect(layout.isRowVisible(0)).toBe(true);
    expect(layout.isRowVisible(1)).toBe(false);
    expect(layout.isRowVisible(2)).toBe(true);
    expect(layout.isRowVisible(3)).toBe(false);
    expect(layout.isRowVisible(4)).toBe(true);

    // Authoritative row heights are untouched.
    expect(layout.root.rowMetrics.heightAt(1)).toBe(20);
    expect(layout.root.rowMetrics.heightAt(3)).toBe(20);
    expect(layout.getRowHeight(1)).toBe(20);

    // Derived visible geometry treats them as 0-height.
    expect(layout.root.visibleRowMetrics.heightAt(1)).toBe(0);
    expect(layout.root.visibleRowMetrics.totalHeight).toBe(3 * 20);
  });

  it("clearing the filter (empty set) restores visibleRowMetrics to alias rowMetrics again", () => {
    const layout = createEngine();
    layout.replaceRoot(layout.prepareSetFilteredRows(new Set([1])));
    layout.replaceRoot(layout.prepareSetFilteredRows(new Set()));

    expect(layout.root.filteredRows.size).toBe(0);
    expect(layout.root.visibleRowMetrics).toBe(layout.root.rowMetrics);
    expect(layout.isRowVisible(1)).toBe(true);
  });
});

describe("LayoutEngine filtered-row geometry: row-height edits (Task 4)", () => {
  it("resizing a filtered row updates the authoritative height but leaves it at 0 in visibleRowMetrics", () => {
    const layout = createEngine();
    layout.replaceRoot(layout.prepareSetFilteredRows(new Set([1])));
    layout.setRowHeight(1, 99);

    expect(layout.root.rowMetrics.heightAt(1)).toBe(99);
    expect(layout.root.visibleRowMetrics.heightAt(1)).toBe(0);
  });

  it("un-filtering a row that was resized while filtered shows the new height, not the pre-resize one", () => {
    const layout = createEngine();
    layout.replaceRoot(layout.prepareSetFilteredRows(new Set([1])));
    layout.setRowHeight(1, 99);
    layout.replaceRoot(layout.prepareSetFilteredRows(new Set()));

    expect(layout.root.visibleRowMetrics.heightAt(1)).toBe(99);
  });

  it("resizing a visible row updates both trees identically", () => {
    const layout = createEngine();
    layout.replaceRoot(layout.prepareSetFilteredRows(new Set([1])));
    layout.setRowHeight(2, 55);

    expect(layout.root.rowMetrics.heightAt(2)).toBe(55);
    expect(layout.root.visibleRowMetrics.heightAt(2)).toBe(55);
  });

  it("keeps the alias when no filter is active", () => {
    const layout = createEngine();
    layout.setRowHeight(2, 55);
    expect(layout.root.visibleRowMetrics).toBe(layout.root.rowMetrics);
  });
});

describe("LayoutEngine filtered-row geometry: structural ops (Task 5)", () => {
  it("insertRows shifts filtered-row membership and never marks new rows filtered", () => {
    const layout = createEngine(5);
    layout.replaceRoot(layout.prepareSetFilteredRows(new Set([1, 3])));
    layout.insertRows(2, 2);

    // row 1 (< insert point 2) stays; row 3 (>= 2) shifts to 5.
    expect(layout.isRowVisible(1)).toBe(false);
    expect(layout.isRowVisible(5)).toBe(false);
    // Newly inserted rows 2,3 are visible.
    expect(layout.isRowVisible(2)).toBe(true);
    expect(layout.isRowVisible(3)).toBe(true);
    expect(layout.root.rowCount).toBe(7);
    expect(layout.root.visibleRowMetrics.rowCount).toBe(7);
  });

  it("deleteRows drops filtered rows caught in the deleted range and shifts the rest down", () => {
    const layout = createEngine(6);
    layout.replaceRoot(layout.prepareSetFilteredRows(new Set([1, 2, 4])));
    layout.deleteRows(2, 2); // deletes rows [2,4)

    expect(layout.root.rowCount).toBe(4);
    // row 1 stays at 1; row 2 was deleted; row 4 shifts to 2.
    expect(layout.isRowVisible(1)).toBe(false);
    expect(layout.isRowVisible(2)).toBe(false);
    expect([...layout.root.filteredRows].sort()).toEqual([1, 2]);
  });

  it("moveRows relocates a filtered row inside the moved block to its new position", () => {
    const layout = createEngine(6);
    layout.replaceRoot(layout.prepareSetFilteredRows(new Set([1])));
    layout.moveRows(0, 2, 5); // moves rows [0,2) to before index 5

    // row 1 was inside the moved block [0,2); it travels with the block.
    expect(layout.root.filteredRows.has(1)).toBe(false);
    expect(layout.root.visibleRowMetrics.totalHeight).toBe(layout.root.rowMetrics.totalHeight - 20);
  });

  it("reorderRows never touches filteredRows membership (reference-stable) and mirrors the permutation onto visibleRowMetrics", () => {
    const layout = createEngine(4);
    layout.replaceRoot(layout.prepareSetFilteredRows(new Set([1])));
    const before = layout.root.filteredRows;
    // Identity for the filtered row (index 1 unchanged), visible rows 0,2,3 permuted arbitrarily.
    const order = Int32Array.from([2, 1, 3, 0]);
    layout.replaceRoot(layout.prepareReorderRows(order));

    expect(layout.root.filteredRows).toBe(before);
    expect(layout.root.visibleRowMetrics.heightAt(1)).toBe(0);
  });
});

describe("LayoutEngine filtered-row geometry: row count changes", () => {
  it("growing row count never marks the new rows filtered", () => {
    const layout = createEngine(3);
    layout.replaceRoot(layout.prepareSetFilteredRows(new Set([1])));
    layout.setRowCount(6);

    expect(layout.isRowVisible(1)).toBe(false);
    expect(layout.isRowVisible(3)).toBe(true);
    expect(layout.isRowVisible(5)).toBe(true);
    expect(layout.root.visibleRowMetrics.rowCount).toBe(6);
  });

  it("shrinking row count drops out-of-range filtered rows", () => {
    const layout = createEngine(6);
    layout.replaceRoot(layout.prepareSetFilteredRows(new Set([1, 4])));
    layout.setRowCount(3);

    expect([...layout.root.filteredRows]).toEqual([1]);
    expect(layout.root.visibleRowMetrics.rowCount).toBe(3);
  });
});
