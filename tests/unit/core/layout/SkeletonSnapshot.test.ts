import { describe, expect, it } from "vitest";

import { ColumnAxisSnapshot, LayoutEngine, SkeletonSnapshot } from "../../../../src/core/layout";

describe("SkeletonSnapshot", () => {
  function createLayout(): LayoutEngine {
    return new LayoutEngine({
      rowCount: 4,
      columnCount: 3,
      rowHeight: 20,
      rowHeights: new Map([[1, 35]]),
      columnWidths: [50, 60, 70],
      freeze: { rows: 1, columns: 1 }
    });
  }

  it("captures logical geometry from LayoutEngine for one model/layout version", () => {
    const layout = createLayout();
    const skeleton = SkeletonSnapshot.fromLayout(layout, 12);

    expect(skeleton.modelVersion).toBe(12);
    expect(skeleton.layoutVersion).toBe(layout.version);
    expect(skeleton.rowCount).toBe(4);
    expect(skeleton.columnCount).toBe(3);
    expect(skeleton.frozenRowCount).toBe(1);
    expect(skeleton.frozenColumnCount).toBe(1);
    expect([0, 1, 2, 3].map((row) => skeleton.getRowHeight(row))).toEqual([20, 35, 20, 20]);
    expect([0, 1, 2, 3].map((row) => skeleton.getRowOffset(row))).toEqual([0, 20, 55, 75]);
    expect([0, 1, 2].map((column) => skeleton.getColumnLeft(column))).toEqual([0, 50, 110]);
    expect(skeleton.getCellRect(1, 2)).toEqual({ x: 110, y: 20, width: 70, height: 35 });
    expect(skeleton.getRangeRect({ startRow: 1, startColumn: 1, endRow: 2, endColumn: 2 })).toEqual({
      x: 50,
      y: 20,
      width: 130,
      height: 55
    });
  });

  it("does not change when the source LayoutEngine mutates later", () => {
    const layout = createLayout();
    const skeleton = SkeletonSnapshot.fromLayout(layout, 1);

    layout.setRowHeight(1, 80);
    layout.setColumnWidth(0, 100);
    layout.insertRows(0, 1);
    layout.insertColumns(0, [25]);

    expect(skeleton.layoutVersion).toBe(0);
    expect(skeleton.rowCount).toBe(4);
    expect(skeleton.columnCount).toBe(3);
    expect(skeleton.getCellRect(1, 1)).toEqual({ x: 50, y: 20, width: 60, height: 35 });
    expect(skeleton.getTotalWidth()).toBe(180);
    expect(skeleton.getTotalHeight()).toBe(95);
  });

  it("shares immutable row metrics and column axis with the layout root in O(1)", () => {
    const layout = createLayout();
    const root = layout.root;
    const skeleton = SkeletonSnapshot.fromRoot(root, 1);

    expect(skeleton.rowMetrics).toBe(root.rowMetrics);
    expect(skeleton.columnAxis).toBe(root.columnAxis);
    expect(layout.prepareSetColumnWidth(0, 90).rowMetrics).toBe(root.rowMetrics);
    expect(layout.prepareHideColumn(0).rowMetrics).toBe(root.rowMetrics);
  });

  it("copies axis constructor arrays instead of retaining caller-owned mutable buffers", () => {
    const rowHeights = new Float64Array([10, 20, 30]);
    const columnWidths = [40, 50];
    const columnAxis = new ColumnAxisSnapshot({ columnWidths, frozenColumnCount: 0 });
    const skeleton = new SkeletonSnapshot({
      modelVersion: 1,
      layoutVersion: 2,
      rowHeights,
      columnAxis,
      frozenRowCount: 0
    });

    rowHeights[1] = 999;
    columnWidths[0] = 999;

    expect(skeleton.getRowHeight(1)).toBe(20);
    expect(skeleton.getColumnWidth(0)).toBe(40);
    expect(skeleton.getCellRect(2, 1)).toEqual({ x: 40, y: 30, width: 50, height: 30 });
  });

  it("computes visible ranges against frozen logical boundaries", () => {
    const skeleton = SkeletonSnapshot.fromLayout(createLayout(), 1);

    expect(skeleton.getVisibleRowRange({ scrollTop: 0, height: 60 })).toEqual({ start: 1, count: 2 });
    expect(skeleton.getVisibleColumnRange({ scrollLeft: 0, width: 120 })).toEqual({ start: 1, count: 2 });
  });

  it("applies viewport overscan with the same clamping semantics as LayoutEngine", () => {
    const layout = new LayoutEngine({
      rowCount: 10,
      columnCount: 5,
      rowHeight: 20,
      columnWidths: [50, 50, 50, 50, 50],
      freeze: { rows: 1, columns: 1 },
      overscanRows: 2,
      overscanColumns: 1
    });
    const skeleton = SkeletonSnapshot.fromLayout(layout, 1);
    const viewport = {
      scrollTop: 47,
      scrollLeft: 13,
      width: 180,
      height: 90,
      overscanRows: 2,
      overscanColumns: 1
    };

    expect(skeleton.getVisibleRowRange(viewport)).toEqual(layout.getVisibleRowRange(viewport));
    expect(skeleton.getVisibleColumnRange(viewport)).toEqual(layout.getVisibleColumnRange(viewport));
    expect(skeleton.getVisibleRowRange(viewport)).toEqual({ start: 1, count: 8 });
    expect(skeleton.getVisibleColumnRange(viewport)).toEqual({ start: 1, count: 4 });

    const pastExtent = { ...viewport, scrollTop: 10_000, scrollLeft: 10_000 };
    expect(skeleton.getVisibleRowRange(pastExtent)).toEqual(layout.getVisibleRowRange(pastExtent));
    expect(skeleton.getVisibleColumnRange(pastExtent)).toEqual(layout.getVisibleColumnRange(pastExtent));
  });

  it("rejects invalid metrics", () => {
    expect(
      () =>
        new SkeletonSnapshot({
          modelVersion: 0,
          layoutVersion: 0,
          rowHeights: [20, 0],
          columnAxis: new ColumnAxisSnapshot({ columnWidths: [50], frozenColumnCount: 0 }),
          frozenRowCount: 0
        })
    ).toThrow(RangeError);

    expect(
      () =>
        new SkeletonSnapshot({
          modelVersion: 0,
          layoutVersion: 0,
          rowHeights: [20],
          columnAxis: new ColumnAxisSnapshot({ columnWidths: [50], frozenColumnCount: 0 }),
          frozenRowCount: 2
        })
    ).toThrow(RangeError);

    expect(() => new ColumnAxisSnapshot({
      columnWidths: [50],
      frozenColumnCount: 2
    })).toThrow(RangeError);
  });

  it("preserves axis identity for row publications and replaces it for every column operation", () => {
    const layout = createLayout();
    const initial = layout.root.columnAxis;

    expect(layout.prepareSetRowHeight(0, 25).columnAxis).toBe(initial);
    expect(layout.prepareInsertRows(1, 1).columnAxis).toBe(initial);
    expect(layout.prepareDeleteRows(1, 1).columnAxis).toBe(initial);
    expect(layout.prepareMoveRows(0, 1, 2).columnAxis).toBe(initial);
    expect(SkeletonSnapshot.fromRoot(layout.root, 1).columnAxis).toBe(initial);
    expect(SkeletonSnapshot.fromRoot(layout.root, 2).columnAxis).toBe(initial);

    expect(layout.prepareSetColumnWidth(0, 55).columnAxis).not.toBe(initial);
    expect(layout.prepareHideColumn(0).columnAxis).not.toBe(initial);
    expect(layout.prepareShowColumn(0).columnAxis).not.toBe(initial);
    expect(layout.prepareInsertColumns(1, [45]).columnAxis).not.toBe(initial);
    expect(layout.prepareDeleteColumns(1, 1).columnAxis).not.toBe(initial);
    expect(layout.prepareMoveColumns(0, 1, 2).columnAxis).not.toBe(initial);
  });

  it("does not expose mutable width, offset, or hidden-map handles", () => {
    const layout = createLayout();
    layout.hideColumn(1);
    const axis = layout.root.columnAxis;
    const widths = axis.copyColumnWidths();
    const hidden = axis.copyHiddenColumnWidths();

    widths[0] = 999;
    hidden.set(1, 999);
    expect(axis.getColumnWidth(0)).toBe(50);
    expect(axis.getColumnRestoreWidth(1)).toBe(60);
    expect(Object.isFrozen(axis.columnOffsets)).toBe(true);
    expect(() => {
      (axis.columnOffsets as number[])[1] = 999;
    }).toThrow(TypeError);
    expect(axis.getColumnOffset(1)).toBe(50);
  });
});
