import { describe, expect, it } from 'vitest';

import {
  LayoutEngine,
} from '../../../../src/core/layout';
import { CyExcelError, CyExcelErrorCode } from '../../../../src/core/errors';

describe('LayoutEngine construction', () => {
  const base = {
    rowCount: 10,
    columnCount: 2,
    rowHeight: 20,
    columnWidths: [50, 60],
  };

  it.each([
    ['negative rowCount', { ...base, rowCount: -1 }],
    ['non-integer rowCount', { ...base, rowCount: 1.5 }],
    ['negative columnCount', { ...base, columnCount: -1 }],
    ['zero rowHeight', { ...base, rowHeight: 0 }],
    ['negative rowHeight', { ...base, rowHeight: -20 }],
    ['non-finite rowHeight', { ...base, rowHeight: Number.POSITIVE_INFINITY }],
    ['columnWidths length mismatch', { ...base, columnWidths: [50] }],
    ['negative column width', { ...base, columnWidths: [50, -1] }],
    ['non-finite column width', { ...base, columnWidths: [50, Number.NaN] }],
    ['non-integer overscanRows', { ...base, overscanRows: 1.5 }],
    ['negative overscanColumns', { ...base, overscanColumns: -1 }],
    ['non-integer freeze.rows', { ...base, freeze: { rows: 1.5 } }],
    ['negative freeze.columns', { ...base, freeze: { columns: -1 } }],
  ])('rejects %s', (_label, options) => {
    expect(() => new LayoutEngine(options)).toThrow(CyExcelError);
  });

  it('clamps freeze counts that exceed the actual row/column count instead of throwing', () => {
    const layout = new LayoutEngine({
      ...base,
      freeze: { rows: 999, columns: 999 },
    });
    expect(layout.frozenRowCount).toBe(10);
    expect(layout.frozenColumnCount).toBe(2);
  });

  it('defaults freeze and overscan to 0 when not provided', () => {
    const layout = new LayoutEngine(base);
    expect(layout.frozenRowCount).toBe(0);
    expect(layout.frozenColumnCount).toBe(0);
  });

  it('exposes the constructor inputs through basic getters', () => {
    const layout = new LayoutEngine(base);
    expect(layout.rowCount).toBe(10);
    expect(layout.columnCount).toBe(2);
    expect(layout.rowHeight).toBe(20);
  });
});

describe('LayoutEngine geometry primitives', () => {
  const layout = new LayoutEngine({
    rowCount: 10,
    columnCount: 3,
    rowHeight: 20,
    columnWidths: [50, 100, 150],
  });

  it('computes cell rects from row height and cumulative column widths', () => {
    expect(layout.getCellRect(0, 0)).toEqual({
      x: 0,
      y: 0,
      width: 50,
      height: 20,
    });
    expect(layout.getCellRect(2, 1)).toEqual({
      x: 50,
      y: 40,
      width: 100,
      height: 20,
    });
    expect(layout.getCellRect(0, 2)).toEqual({
      x: 150,
      y: 0,
      width: 150,
      height: 20,
    });
  });

  it('reports total width and height', () => {
    expect(layout.getTotalWidth()).toBe(300);
    expect(layout.getTotalHeight()).toBe(200);
  });

  it('returns 0 width for an out-of-range column', () => {
    expect(layout.getColumnWidth(99)).toBe(0);
  });

  it('falls back to x=0 for an out-of-range column in getCellRect', () => {
    expect(layout.getCellRect(0, 99)).toEqual({
      x: 0,
      y: 0,
      width: 0,
      height: 20,
    });
  });

  it('getRowAtY/getColumnAtX clamp to valid bounds', () => {
    expect(layout.getRowAtY(-5)).toBe(0);
    expect(layout.getRowAtY(0)).toBe(0);
    expect(layout.getRowAtY(45)).toBe(2);
    expect(layout.getRowAtY(10_000)).toBe(9);

    expect(layout.getColumnAtX(-5)).toBe(0);
    expect(layout.getColumnAtX(60)).toBe(1);
    expect(layout.getColumnAtX(10_000)).toBe(2);
  });

  it('getRowAtY/getColumnAtX return 0 for an empty table', () => {
    const empty = new LayoutEngine({
      rowCount: 0,
      columnCount: 0,
      rowHeight: 20,
      columnWidths: [],
    });
    expect(empty.getRowAtY(50)).toBe(0);
    expect(empty.getColumnAtX(50)).toBe(0);
  });
});

describe('LayoutEngine variable row heights', () => {
  function createLayout(): LayoutEngine {
    return new LayoutEngine({
      rowCount: 5,
      columnCount: 1,
      rowHeight: 20,
      rowHeights: new Map([
        [1, 40],
        [3, 10],
      ]),
      columnWidths: [50],
    });
  }

  it('uses sparse constructor overrides for heights, offsets, rects, and total height', () => {
    const layout = createLayout();

    expect([0, 1, 2, 3, 4].map((row) => layout.getRowHeight(row))).toEqual([
      20, 40, 20, 10, 20,
    ]);
    expect([0, 1, 2, 3, 4, 5].map((row) => layout.getRowOffset(row))).toEqual([
      0, 20, 60, 80, 90, 110,
    ]);
    expect(layout.getCellRect(3, 0)).toEqual({
      x: 0,
      y: 80,
      width: 50,
      height: 10,
    });
    expect(layout.getTotalHeight()).toBe(110);
    expect(layout.getRowOffset(-1)).toBe(0);
    expect(layout.getRowOffset(6)).toBe(0);
  });

  it('setRowHeight updates downstream offsets and reverse lookups', () => {
    const layout = createLayout();

    layout.setRowHeight(1, 50);

    expect(layout.getRowHeight(1)).toBe(50);
    expect(layout.getRowOffset(2)).toBe(70);
    expect(layout.getRowAtY(69)).toBe(1);
    expect(layout.getRowAtY(70)).toBe(2);
    expect(
      layout.getCellAtPoint(
        { x: 5, y: 65 },
        { scrollTop: 0, scrollLeft: 0, width: 50, height: 100 },
      ),
    ).toEqual({ row: 1, column: 0 });
  });

  it('prepares several row-height changes in one immutable root/version', () => {
    const layout = createLayout();
    const before = layout.root;
    const prepared = layout.prepareSetRowHeights(new Map([
      [0, 25],
      [3, 35],
    ]));

    expect(layout.root).toBe(before);
    expect(prepared.version).toBe(before.version + 1);
    expect(prepared.rowMetrics.exportDenseValues()).toEqual(new Float64Array([25, 40, 20, 35, 20]));

    layout.replaceRoot(prepared);
    expect(layout.getRowOffset(4)).toBe(120);
  });

  it('returns the base root when every row-height update is equal', () => {
    const layout = createLayout();
    const before = layout.root;

    expect(layout.prepareSetRowHeights(new Map([
      [0, 20],
      [2, 20],
    ]))).toBe(before);
    expect(layout.root).toBe(before);
    expect(layout.version).toBe(0);
  });

  it('shares row metrics across column-only roots', () => {
    const layout = createLayout();
    const metrics = layout.root.rowMetrics;

    expect(layout.prepareSetColumnWidth(0, 80).rowMetrics).toBe(metrics);
    expect(layout.prepareHideColumn(0).rowMetrics).toBe(metrics);
    expect(layout.prepareInsertColumns(0, [30]).rowMetrics).toBe(metrics);
  });

  it('rejects an invalid row-height batch without changing the live root', () => {
    const layout = createLayout();
    const before = layout.root;

    expect(() => layout.prepareSetRowHeights(new Map([
      [0, 25],
      [99, 30],
    ]))).toThrow(CyExcelError);
    expect(layout.root).toBe(before);
  });

  it('maps visible ranges through variable-height boundaries', () => {
    const layout = createLayout();

    expect(
      layout.getVisibleRowRange({
        scrollTop: 25,
        scrollLeft: 0,
        width: 50,
        height: 50,
      }),
    ).toEqual({
      start: 1,
      count: 2,
    });
  });

  it('preserves custom heights through insert, delete, and move operations', () => {
    const layout = createLayout();

    layout.insertRows(2, 2, new Map([[0, 33]]));
    expect(
      Array.from({ length: 7 }, (_unused, row) => layout.getRowHeight(row)),
    ).toEqual([20, 40, 33, 20, 20, 10, 20]);

    layout.deleteRows(1, 2);
    expect(
      Array.from({ length: 5 }, (_unused, row) => layout.getRowHeight(row)),
    ).toEqual([20, 20, 20, 10, 20]);

    layout.moveRows(3, 1, 1);
    expect(
      Array.from({ length: 5 }, (_unused, row) => layout.getRowHeight(row)),
    ).toEqual([20, 10, 20, 20, 20]);
  });

  it('rejects invalid row structure changes before changing geometry version or row count', () => {
    const layout = createLayout();
    const version = layout.version;

    expect(() => layout.insertRows(1, 0)).toThrow(RangeError);
    expect(() => layout.deleteRows(4, 2)).toThrow(RangeError);
    expect(() => layout.moveRows(1, 2, 2)).toThrow(RangeError);

    expect(layout.version).toBe(version);
    expect(layout.rowCount).toBe(5);
    expect(
      Array.from({ length: 5 }, (_unused, row) => layout.getRowHeight(row)),
    ).toEqual([20, 40, 20, 10, 20]);
  });

  it('rejects invalid constructor overrides and row-height updates', () => {
    expect(
      () =>
        new LayoutEngine({
          rowCount: 1,
          columnCount: 1,
          rowHeight: 20,
          rowHeights: new Map([[1, 30]]),
          columnWidths: [50],
        }),
    ).toThrow(CyExcelError);

    const layout = createLayout();
    expect(() => layout.setRowHeight(99, 30)).toThrow(CyExcelError);
    expect(() => layout.setRowHeight(0, 0)).toThrow(CyExcelError);
  });
});

describe('LayoutEngine.getVisibleRowRange — test 1: scrollTop under fixed row height', () => {
  const layout = new LayoutEngine({
    rowCount: 100,
    columnCount: 1,
    rowHeight: 20,
    columnWidths: [50],
  });

  it('maps scrollTop 0 to the first rows', () => {
    expect(
      layout.getVisibleRowRange({
        scrollTop: 0,
        scrollLeft: 0,
        width: 50,
        height: 100,
      }),
    ).toEqual({
      start: 0,
      count: 6,
    });
  });

  it('maps a scrolled-down scrollTop to the corresponding rows', () => {
    expect(
      layout.getVisibleRowRange({
        scrollTop: 205,
        scrollLeft: 0,
        width: 50,
        height: 100,
      }),
    ).toEqual({
      start: 10,
      count: 6,
    });
  });
});

describe('LayoutEngine.getVisibleColumnRange — test 2: scrollLeft under variable column widths', () => {
  const layout = new LayoutEngine({
    rowCount: 1,
    columnCount: 3,
    rowHeight: 20,
    columnWidths: [50, 100, 150],
  });

  it('maps scrollLeft 0 to the leftmost visible columns', () => {
    expect(
      layout.getVisibleColumnRange({
        scrollTop: 0,
        scrollLeft: 0,
        width: 120,
        height: 20,
      }),
    ).toEqual({
      start: 0,
      count: 2,
    });
  });

  it('maps a scrolled scrollLeft to the correct visible columns', () => {
    expect(
      layout.getVisibleColumnRange({
        scrollTop: 0,
        scrollLeft: 60,
        width: 120,
        height: 20,
      }),
    ).toEqual({
      start: 1,
      count: 2,
    });
  });
});

describe('LayoutEngine overscan — test 3: never exceeds row/column bounds', () => {
  const layout = new LayoutEngine({
    rowCount: 10,
    columnCount: 1,
    rowHeight: 20,
    columnWidths: [50],
    overscanRows: 5,
  });

  it('clamps overscan at the top edge instead of going negative', () => {
    expect(
      layout.getVisibleRowRange({
        scrollTop: 0,
        scrollLeft: 0,
        width: 50,
        height: 40,
      }),
    ).toEqual({
      start: 0,
      count: 8,
    });
  });

  it('clamps overscan at the bottom edge instead of exceeding rowCount', () => {
    expect(
      layout.getVisibleRowRange({
        scrollTop: 160,
        scrollLeft: 0,
        width: 50,
        height: 40,
      }),
    ).toEqual({
      start: 3,
      count: 7,
    });
  });
});

describe('LayoutEngine freeze — test 4: frozen rows/columns occupy fixed viewport space', () => {
  it("reduces the scrollable visible row count by the frozen band's height", () => {
    const withoutFreeze = new LayoutEngine({
      rowCount: 20,
      columnCount: 1,
      rowHeight: 10,
      columnWidths: [50],
    });
    const withFreeze = new LayoutEngine({
      rowCount: 20,
      columnCount: 1,
      rowHeight: 10,
      columnWidths: [50],
      freeze: { rows: 3 },
    });
    const viewport = { scrollTop: 0, scrollLeft: 0, width: 50, height: 50 };

    expect(withFreeze.frozenRowCount).toBe(3);
    const withoutRange = withoutFreeze.getVisibleRowRange(viewport);
    const withRange = withFreeze.getVisibleRowRange(viewport);

    // 3 frozen rows * 10px row height = 30px less scrollable viewport space,
    // i.e. exactly 3 fewer visible rows at the same row height.
    expect(withoutRange.count - withRange.count).toBe(3);
  });

  it('frozen cell rects never depend on scroll position', () => {
    const layout = new LayoutEngine({
      rowCount: 20,
      columnCount: 1,
      rowHeight: 10,
      columnWidths: [50],
      freeze: { rows: 3 },
    });

    expect(layout.getCellRect(1, 0)).toEqual(layout.getCellRect(1, 0));
  });
});

describe('LayoutEngine — test 5: logical geometry never depends on DPR', () => {
  it('does not accept a dpr parameter and stays deterministic across repeated calls', () => {
    const layout = new LayoutEngine({
      rowCount: 10,
      columnCount: 2,
      rowHeight: 20,
      columnWidths: [50, 60],
    });
    const viewport = { scrollTop: 0, scrollLeft: 0, width: 100, height: 100 };

    expect(layout.getVisibleRowRange(viewport)).toEqual(
      layout.getVisibleRowRange(viewport),
    );
    expect(layout.getCellRect(0, 0)).toEqual(layout.getCellRect(0, 0));
  });
});

describe('LayoutEngine — test 6: no floating point drift at large row indices', () => {
  it('computes an exact cell rect near row 1,000,000', () => {
    const layout = new LayoutEngine({
      rowCount: 1_000_000,
      columnCount: 1,
      rowHeight: 24,
      columnWidths: [80],
    });

    expect(layout.getCellRect(999_998, 0).y).toBe(999_998 * 24);
  });
});

describe('LayoutEngine — test 7: no negative ranges for empty/zero/hidden edge cases', () => {
  it('returns a zero-count range for an empty table', () => {
    const empty = new LayoutEngine({
      rowCount: 0,
      columnCount: 0,
      rowHeight: 20,
      columnWidths: [],
    });
    const viewport = { scrollTop: 0, scrollLeft: 0, width: 100, height: 100 };

    const rowRange = empty.getVisibleRowRange(viewport);
    const columnRange = empty.getVisibleColumnRange(viewport);

    expect(rowRange).toEqual({ start: 0, count: 0 });
    expect(columnRange).toEqual({ start: 0, count: 0 });
  });

  it('returns a zero-count range for a zero-size viewport', () => {
    const layout = new LayoutEngine({
      rowCount: 10,
      columnCount: 2,
      rowHeight: 20,
      columnWidths: [50, 50],
    });
    const zeroViewport = { scrollTop: 0, scrollLeft: 0, width: 0, height: 0 };

    expect(layout.getVisibleRowRange(zeroViewport)).toEqual({
      start: 0,
      count: 0,
    });
    expect(layout.getVisibleColumnRange(zeroViewport)).toEqual({
      start: 0,
      count: 0,
    });
  });

  it('never returns a negative start/count when every column is hidden (width 0)', () => {
    const hidden = new LayoutEngine({
      rowCount: 5,
      columnCount: 3,
      rowHeight: 20,
      columnWidths: [0, 0, 0],
    });
    const range = hidden.getVisibleColumnRange({
      scrollTop: 0,
      scrollLeft: 0,
      width: 100,
      height: 100,
    });

    expect(range.start).toBeGreaterThanOrEqual(0);
    expect(range.count).toBeGreaterThanOrEqual(0);
  });
});

describe('LayoutEngine.getCellAtPoint — test 8: stable at boundary pixels', () => {
  const layout = new LayoutEngine({
    rowCount: 20,
    columnCount: 1,
    rowHeight: 10,
    columnWidths: [50],
    freeze: { rows: 2 },
  });
  const viewport = { scrollTop: 0, scrollLeft: 0, width: 50, height: 100 };

  it('resolves the exact frozen/scrollable boundary pixel consistently', () => {
    const atBoundary = layout.getCellAtPoint({ x: 0, y: 20 }, viewport);
    expect(atBoundary).toEqual({ row: 2, column: 0 });
    expect(layout.getCellAtPoint({ x: 0, y: 20 }, viewport)).toEqual(
      atBoundary,
    );
  });

  it('resolves the pixel just before the boundary to the frozen row', () => {
    expect(layout.getCellAtPoint({ x: 0, y: 19 }, viewport)).toEqual({
      row: 1,
      column: 0,
    });
  });

  it('resolves points inside a frozen column band without applying scrollLeft', () => {
    const withFrozenColumns = new LayoutEngine({
      rowCount: 5,
      columnCount: 4,
      rowHeight: 10,
      columnWidths: [30, 30, 30, 30],
      freeze: { columns: 2 },
    });
    const scrolledViewport = {
      scrollTop: 0,
      scrollLeft: 100,
      width: 120,
      height: 50,
    };

    // x=10 is inside the frozen band (0-60px); scrollLeft must not shift it.
    expect(
      withFrozenColumns.getCellAtPoint({ x: 10, y: 0 }, scrolledViewport),
    ).toEqual({
      row: 0,
      column: 0,
    });
  });

  it('returns null for negative coordinates', () => {
    expect(layout.getCellAtPoint({ x: -1, y: 0 }, viewport)).toBeNull();
    expect(layout.getCellAtPoint({ x: 0, y: -1 }, viewport)).toBeNull();
  });

  it('returns null for an empty table', () => {
    const empty = new LayoutEngine({
      rowCount: 0,
      columnCount: 0,
      rowHeight: 20,
      columnWidths: [],
    });
    expect(empty.getCellAtPoint({ x: 0, y: 0 }, viewport)).toBeNull();
  });

  it('returns null for a point beyond the last row or column', () => {
    const small = new LayoutEngine({
      rowCount: 2,
      columnCount: 1,
      rowHeight: 10,
      columnWidths: [50],
    });
    const smallViewport = {
      scrollTop: 0,
      scrollLeft: 0,
      width: 50,
      height: 200,
    };
    expect(small.getCellAtPoint({ x: 0, y: 500 }, smallViewport)).toBeNull();

    const narrow = new LayoutEngine({
      rowCount: 1,
      columnCount: 2,
      rowHeight: 10,
      columnWidths: [50, 50],
    });
    const narrowViewport = {
      scrollTop: 0,
      scrollLeft: 0,
      width: 200,
      height: 10,
    };
    expect(narrow.getCellAtPoint({ x: 500, y: 0 }, narrowViewport)).toBeNull();
  });
});

describe('LayoutEngine performance guard: 1,000,000 rows is not scanned', () => {
  it('computes visible ranges, a cell rect, and a hit test in well under linear time', () => {
    const layout = new LayoutEngine({
      rowCount: 1_000_000,
      columnCount: 50,
      rowHeight: 24,
      columnWidths: Array.from({ length: 50 }, () => 100),
    });
    const viewport = {
      scrollTop: 12_000_000,
      scrollLeft: 2000,
      width: 800,
      height: 600,
    };

    const start = performance.now();
    layout.getVisibleRowRange(viewport);
    layout.getVisibleColumnRange(viewport);
    layout.getCellRect(500_000, 25);
    layout.getCellAtPoint({ x: 400, y: 300 }, viewport);
    const elapsed = performance.now() - start;

    expect(elapsed).toBeLessThan(5);
  });
});

describe('LayoutEngine.getColumnLeft', () => {
  it('returns the same x offset as getCellRect without needing a row', () => {
    const layout = new LayoutEngine({
      rowCount: 5,
      columnCount: 3,
      rowHeight: 20,
      columnWidths: [50, 60, 70],
    });

    expect(layout.getColumnLeft(0)).toBe(0);
    expect(layout.getColumnLeft(1)).toBe(50);
    expect(layout.getColumnLeft(2)).toBe(110);
    expect(layout.getColumnLeft(2)).toBe(layout.getCellRect(0, 2).x);
  });

  it('falls back to 0 for an out-of-range column, same as getCellRect', () => {
    const layout = new LayoutEngine({
      rowCount: 5,
      columnCount: 3,
      rowHeight: 20,
      columnWidths: [50, 60, 70],
    });

    expect(layout.getColumnLeft(99)).toBe(0);
  });
});

describe('LayoutEngine.setColumnWidth', () => {
  function createLayout(): LayoutEngine {
    return new LayoutEngine({
      rowCount: 5,
      columnCount: 3,
      rowHeight: 20,
      columnWidths: [50, 60, 70],
    });
  }

  it("updates the column's own width", () => {
    const layout = createLayout();

    layout.setColumnWidth(0, 120);

    expect(layout.getColumnWidth(0)).toBe(120);
  });

  it("shifts every subsequent column's offset by the width delta, leaving earlier ones untouched", () => {
    const layout = createLayout();

    layout.setColumnWidth(0, 120); // was 50, delta +70

    expect(layout.getColumnLeft(0)).toBe(0); // unaffected — before the resized column
    expect(layout.getColumnLeft(1)).toBe(120); // was 50, now shifted by +70
    expect(layout.getColumnLeft(2)).toBe(180); // was 110, now shifted by +70
    expect(layout.getTotalWidth()).toBe(120 + 60 + 70);
  });

  it('getCellRect reflects the new width and shifted offsets immediately', () => {
    const layout = createLayout();

    layout.setColumnWidth(1, 200);

    expect(layout.getCellRect(2, 1)).toEqual({
      x: 50,
      y: 40,
      width: 200,
      height: 20,
    });
    expect(layout.getCellRect(2, 2).x).toBe(250);
  });

  it('bumps version by exactly 1 per successful call, starting at 0', () => {
    const layout = createLayout();
    expect(layout.version).toBe(0);

    layout.setColumnWidth(0, 100);
    expect(layout.version).toBe(1);

    layout.setColumnWidth(1, 90);
    expect(layout.version).toBe(2);
  });

  it('throws and leaves state/version unchanged for a negative or non-finite width', () => {
    const layout = createLayout();

    expect(() => layout.setColumnWidth(0, -1)).toThrow(CyExcelError);
    expect(() => layout.setColumnWidth(0, Number.NaN)).toThrow(CyExcelError);
    try {
      layout.setColumnWidth(0, -1);
    } catch (error) {
      expect((error as CyExcelError).code).toBe(
        CyExcelErrorCode.InvalidOptions,
      );
    }
    expect(layout.getColumnWidth(0)).toBe(50);
    expect(layout.version).toBe(0);
  });

  it('throws for an out-of-range column index', () => {
    const layout = createLayout();

    expect(() => layout.setColumnWidth(99, 100)).toThrow(CyExcelError);
    expect(layout.version).toBe(0);
  });

  it('allows a width of 0 (hidden column)', () => {
    const layout = createLayout();

    layout.setColumnWidth(0, 0);

    expect(layout.getColumnWidth(0)).toBe(0);
    expect(layout.getColumnLeft(1)).toBe(0);
  });
});

describe('LayoutEngine.setRowCount', () => {
  it('updates rowCount and getTotalHeight immediately', () => {
    const layout = new LayoutEngine({
      rowCount: 5,
      columnCount: 1,
      rowHeight: 20,
      columnWidths: [50],
    });

    layout.setRowCount(8);

    expect(layout.rowCount).toBe(8);
    expect(layout.getTotalHeight()).toBe(160);
  });

  it('bumps version by exactly 1 per successful call, starting at 0', () => {
    const layout = new LayoutEngine({
      rowCount: 5,
      columnCount: 1,
      rowHeight: 20,
      columnWidths: [50],
    });
    expect(layout.version).toBe(0);

    layout.setRowCount(8);
    expect(layout.version).toBe(1);

    layout.setRowCount(3);
    expect(layout.version).toBe(2);
  });

  it('re-clamps frozenRowCount against the new rowCount when it shrinks below the frozen band', () => {
    const layout = new LayoutEngine({
      rowCount: 10,
      columnCount: 1,
      rowHeight: 20,
      columnWidths: [50],
      freeze: { rows: 4 },
    });
    expect(layout.frozenRowCount).toBe(4);

    layout.setRowCount(2);

    expect(layout.frozenRowCount).toBe(2);
  });

  it('recovers frozen rows on a later grow, re-clamping against the originally requested freeze.rows', () => {
    const layout = new LayoutEngine({
      rowCount: 10,
      columnCount: 1,
      rowHeight: 20,
      columnWidths: [50],
      freeze: { rows: 4 },
    });

    layout.setRowCount(2); // frozenRowCount clamps to 2
    layout.setRowCount(10); // grows back — should recover the originally requested 4, not stay at 2

    expect(layout.frozenRowCount).toBe(4);
  });

  it('does not touch column state', () => {
    const layout = new LayoutEngine({
      rowCount: 5,
      columnCount: 2,
      rowHeight: 20,
      columnWidths: [50, 60],
    });

    layout.setRowCount(8);

    expect(layout.getColumnWidth(0)).toBe(50);
    expect(layout.getColumnLeft(1)).toBe(50);
  });

  it('throws and leaves state/version unchanged for a negative or non-integer rowCount', () => {
    const layout = new LayoutEngine({
      rowCount: 5,
      columnCount: 1,
      rowHeight: 20,
      columnWidths: [50],
    });

    expect(() => layout.setRowCount(-1)).toThrow(CyExcelError);
    expect(() => layout.setRowCount(1.5)).toThrow(CyExcelError);
    expect(layout.rowCount).toBe(5);
    expect(layout.version).toBe(0);
  });

  it('allows a rowCount of 0', () => {
    const layout = new LayoutEngine({
      rowCount: 5,
      columnCount: 1,
      rowHeight: 20,
      columnWidths: [50],
    });

    layout.setRowCount(0);

    expect(layout.rowCount).toBe(0);
    expect(layout.getTotalHeight()).toBe(0);
  });

  it('preserves retained overrides, appends defaults, and drops truncated overrides', () => {
    const layout = new LayoutEngine({
      rowCount: 4,
      columnCount: 1,
      rowHeight: 20,
      rowHeights: new Map([
        [1, 35],
        [3, 50],
      ]),
      columnWidths: [50],
    });

    layout.setRowCount(3);
    layout.setRowCount(5);

    expect(
      Array.from({ length: 5 }, (_unused, row) => layout.getRowHeight(row)),
    ).toEqual([20, 35, 20, 20, 20]);
  });
});

describe('LayoutEngine.insertColumns', () => {
  function createLayout(): LayoutEngine {
    return new LayoutEngine({
      rowCount: 5,
      columnCount: 3,
      rowHeight: 20,
      columnWidths: [50, 60, 70],
    });
  }

  it('inserts new columns at the given position with the given widths', () => {
    const layout = createLayout();

    layout.insertColumns(1, [30, 40]);

    expect(layout.columnCount).toBe(5);
    expect(
      [0, 1, 2, 3, 4].map((column) => layout.getColumnWidth(column)),
    ).toEqual([50, 30, 40, 60, 70]);
  });

  it("shifts every later column's offset to account for the inserted widths", () => {
    const layout = createLayout();

    layout.insertColumns(1, [30, 40]);

    expect(
      [0, 1, 2, 3, 4].map((column) => layout.getColumnLeft(column)),
    ).toEqual([0, 50, 80, 120, 180]);
    expect(layout.getTotalWidth()).toBe(50 + 30 + 40 + 60 + 70);
  });

  it('bumps version by exactly 1', () => {
    const layout = createLayout();
    expect(layout.version).toBe(0);

    layout.insertColumns(0, [10]);

    expect(layout.version).toBe(1);
  });

  it('does not touch row state', () => {
    const layout = createLayout();

    layout.insertColumns(0, [10]);

    expect(layout.rowCount).toBe(5);
    expect(layout.getTotalHeight()).toBe(100);
  });

  it('rejects zero-count and out-of-range inserts before changing geometry', () => {
    const layout = createLayout();
    const version = layout.version;

    expect(() => layout.insertColumns(1, [])).toThrow(RangeError);
    expect(() => layout.insertColumns(4, [10])).toThrow(RangeError);

    expect(layout.version).toBe(version);
    expect(layout.columnCount).toBe(3);
    expect([0, 1, 2].map((column) => layout.getColumnWidth(column))).toEqual([
      50, 60, 70,
    ]);
  });
});

describe('LayoutEngine.deleteColumns', () => {
  function createLayout(): LayoutEngine {
    return new LayoutEngine({
      rowCount: 5,
      columnCount: 4,
      rowHeight: 20,
      columnWidths: [50, 60, 70, 80],
    });
  }

  it('removes the given range of columns', () => {
    const layout = createLayout();

    layout.deleteColumns(1, 2);

    expect(layout.columnCount).toBe(2);
    expect([0, 1].map((column) => layout.getColumnWidth(column))).toEqual([
      50, 80,
    ]);
  });

  it('shifts remaining offsets to close the gap', () => {
    const layout = createLayout();

    layout.deleteColumns(1, 2);

    expect(layout.getColumnLeft(1)).toBe(50);
    expect(layout.getTotalWidth()).toBe(50 + 80);
  });

  it('bumps version by exactly 1', () => {
    const layout = createLayout();

    layout.deleteColumns(0, 1);

    expect(layout.version).toBe(1);
  });

  it('rejects zero-count and out-of-range deletes before changing geometry', () => {
    const layout = createLayout();
    const version = layout.version;

    expect(() => layout.deleteColumns(1, 0)).toThrow(RangeError);
    expect(() => layout.deleteColumns(3, 2)).toThrow(RangeError);

    expect(layout.version).toBe(version);
    expect(layout.columnCount).toBe(4);
    expect([0, 1, 2, 3].map((column) => layout.getColumnWidth(column))).toEqual(
      [50, 60, 70, 80],
    );
  });
});

describe('LayoutEngine.moveColumns', () => {
  it('moves a block of columns to a new position, widths/offsets following', () => {
    const layout = new LayoutEngine({
      rowCount: 5,
      columnCount: 4,
      rowHeight: 20,
      columnWidths: [10, 20, 30, 40],
    });

    layout.moveColumns(0, 1, 4); // [10,20,30,40] -> [20,30,40,10]

    expect([0, 1, 2, 3].map((column) => layout.getColumnWidth(column))).toEqual(
      [20, 30, 40, 10],
    );
    expect(layout.getTotalWidth()).toBe(100);
  });

  it('bumps version by exactly 1', () => {
    const layout = new LayoutEngine({
      rowCount: 5,
      columnCount: 3,
      rowHeight: 20,
      columnWidths: [10, 20, 30],
    });

    layout.moveColumns(0, 1, 2);

    expect(layout.version).toBe(1);
  });

  it('rejects invalid moves before changing geometry', () => {
    const layout = new LayoutEngine({
      rowCount: 5,
      columnCount: 3,
      rowHeight: 20,
      columnWidths: [10, 20, 30],
    });
    const version = layout.version;

    expect(() => layout.moveColumns(0, 0, 3)).toThrow(RangeError);
    expect(() => layout.moveColumns(0, 2, 1)).toThrow(RangeError);
    expect(() => layout.moveColumns(2, 2, 0)).toThrow(RangeError);

    expect(layout.version).toBe(version);
    expect(layout.columnCount).toBe(3);
    expect([0, 1, 2].map((column) => layout.getColumnWidth(column))).toEqual([
      10, 20, 30,
    ]);
  });
});

describe('LayoutEngine frozenColumnCount re-clamps after column structural changes', () => {
  function createLayout(): LayoutEngine {
    return new LayoutEngine({
      rowCount: 5,
      columnCount: 5,
      rowHeight: 20,
      columnWidths: [10, 10, 10, 10, 10],
      freeze: { columns: 4 },
    });
  }

  it('re-clamps frozenColumnCount when deleteColumns shrinks below the frozen band', () => {
    const layout = createLayout();
    expect(layout.frozenColumnCount).toBe(4);

    layout.deleteColumns(0, 3);

    expect(layout.frozenColumnCount).toBe(2);
  });

  it('recovers frozen columns on a later insert, re-clamping against the originally requested freeze.columns', () => {
    const layout = createLayout();

    layout.deleteColumns(0, 3); // clamps to 2
    layout.insertColumns(0, [10, 10, 10]); // grows back to 5 columns

    expect(layout.frozenColumnCount).toBe(4);
  });
});

describe('LayoutEngine.hideColumn/showColumn/isColumnHidden', () => {
  function createLayout(): LayoutEngine {
    return new LayoutEngine({
      rowCount: 5,
      columnCount: 3,
      rowHeight: 20,
      columnWidths: [50, 60, 70],
    });
  }

  it('hiding a column sets its width to 0 and marks it hidden', () => {
    const layout = createLayout();

    layout.hideColumn(1);

    expect(layout.getColumnWidth(1)).toBe(0);
    expect(layout.isColumnHidden(1)).toBe(true);
    expect(layout.getTotalWidth()).toBe(50 + 0 + 70);
  });

  it('showing a hidden column restores its exact previous width', () => {
    const layout = createLayout();
    layout.hideColumn(1);

    layout.showColumn(1);

    expect(layout.getColumnWidth(1)).toBe(60);
    expect(layout.isColumnHidden(1)).toBe(false);
    expect(layout.getTotalWidth()).toBe(50 + 60 + 70);
  });

  it('showing a column that was never hidden is a no-op', () => {
    const layout = createLayout();

    layout.showColumn(1);

    expect(layout.getColumnWidth(1)).toBe(60);
  });

  it('hiding an already-hidden column does not overwrite the remembered restore width', () => {
    const layout = createLayout();
    layout.hideColumn(1);

    layout.hideColumn(1); // already hidden — must not remember 0 as the "restore" width

    layout.showColumn(1);
    expect(layout.getColumnWidth(1)).toBe(60);
  });

  it('isColumnHidden is false for a never-hidden column', () => {
    const layout = createLayout();

    expect(layout.isColumnHidden(0)).toBe(false);
  });

  it('bumps version on both hide and show', () => {
    const layout = createLayout();
    expect(layout.version).toBe(0);

    layout.hideColumn(0);
    expect(layout.version).toBe(1);

    layout.showColumn(0);
    expect(layout.version).toBe(2);
  });
});

describe('LayoutEngine root prepare/replace', () => {
  function createLayout(): LayoutEngine {
    return new LayoutEngine({
      rowCount: 4,
      columnCount: 4,
      rowHeight: 20,
      rowHeights: new Map([[1, 35]]),
      columnWidths: [50, 60, 70, 80],
      freeze: { rows: 3, columns: 3 },
    });
  }

  function readRowHeights(layout: LayoutEngine): number[] {
    return Array.from({ length: layout.rowCount }, (_unused, row) =>
      layout.getRowHeight(row),
    );
  }

  function readColumnWidths(layout: LayoutEngine): number[] {
    return Array.from({ length: layout.columnCount }, (_unused, column) =>
      layout.getColumnWidth(column),
    );
  }

  it('prepareInsertRows computes the next root without writing live geometry', () => {
    const layout = createLayout();
    const version = layout.version;

    const prepared = layout.prepareInsertRows(1, 2, new Map([[0, 25]]));

    expect(layout.rowCount).toBe(4);
    expect(readRowHeights(layout)).toEqual([20, 35, 20, 20]);
    expect(layout.version).toBe(version);
    expect(prepared.rowCount).toBe(6);

    layout.replaceRoot(prepared);

    expect(layout.rowCount).toBe(6);
    expect(readRowHeights(layout)).toEqual([20, 25, 20, 35, 20, 20]);
    expect(layout.version).toBe(version + 1);
  });

  it('prepareDeleteColumns remaps hidden-column state without writing live geometry', () => {
    const layout = createLayout();
    layout.hideColumn(3);
    const version = layout.version;

    const prepared = layout.prepareDeleteColumns(1, 1);

    expect(readColumnWidths(layout)).toEqual([50, 60, 70, 0]);
    expect(layout.isColumnHidden(3)).toBe(true);
    expect(layout.version).toBe(version);

    layout.replaceRoot(prepared);

    expect(readColumnWidths(layout)).toEqual([50, 70, 0]);
    expect(layout.isColumnHidden(2)).toBe(true);
    expect(layout.getColumnRestoreWidth(2)).toBe(80);
    expect(layout.version).toBe(version + 1);
  });

  it('prepareMoveColumns remaps hidden-column state without writing live geometry', () => {
    const layout = createLayout();
    layout.hideColumn(1);
    const version = layout.version;

    const prepared = layout.prepareMoveColumns(1, 1, 4);

    expect(readColumnWidths(layout)).toEqual([50, 0, 70, 80]);
    expect(layout.isColumnHidden(1)).toBe(true);
    expect(layout.version).toBe(version);

    layout.replaceRoot(prepared);

    expect(readColumnWidths(layout)).toEqual([50, 70, 80, 0]);
    expect(layout.isColumnHidden(3)).toBe(true);
    expect(layout.getColumnRestoreWidth(3)).toBe(60);
    expect(layout.version).toBe(version + 1);
  });

  it('prepareStructureRoot can restore inserted hidden columns from one immutable base root', () => {
    const layout = createLayout();
    layout.hideColumn(3);
    const base = layout.root;
    const prepared = layout.prepareStructureRoot(
      { axis: 'column', operation: 'insert', start: 1, count: 2 },
      {
        columnWidths: [40, 0],
        hiddenColumnOffsets: new Map([[1, 45]]),
      },
      base,
    );

    expect(readColumnWidths(layout)).toEqual([50, 60, 70, 0]);
    expect(layout.isColumnHidden(2)).toBe(false);
    expect(layout.isColumnHidden(5)).toBe(false);

    layout.replaceRoot(prepared);

    expect(readColumnWidths(layout)).toEqual([50, 40, 0, 60, 70, 0]);
    expect(layout.isColumnHidden(2)).toBe(true);
    expect(layout.getColumnRestoreWidth(2)).toBe(45);
    expect(layout.isColumnHidden(5)).toBe(true);
    expect(layout.getColumnRestoreWidth(5)).toBe(80);
  });

  it('prepareHideColumn and prepareShowColumn do not write until replaceRoot swaps the root', () => {
    const layout = createLayout();
    const hidden = layout.prepareHideColumn(1);

    expect(layout.getColumnWidth(1)).toBe(60);
    expect(layout.isColumnHidden(1)).toBe(false);

    layout.replaceRoot(hidden);
    expect(layout.getColumnWidth(1)).toBe(0);
    expect(layout.isColumnHidden(1)).toBe(true);

    const version = layout.version;
    const shown = layout.prepareShowColumn(1);
    expect(layout.getColumnWidth(1)).toBe(0);
    expect(layout.isColumnHidden(1)).toBe(true);

    layout.replaceRoot(shown);
    expect(layout.getColumnWidth(1)).toBe(60);
    expect(layout.isColumnHidden(1)).toBe(false);
    expect(layout.version).toBe(version + 1);
  });

  it('prepareRoot reuses immutable snapshot roots without exposing mutable buffers', () => {
    const layout = createLayout();
    const snapshot = layout.snapshotRoot();
    const prepared = layout.prepareRoot(snapshot);

    const originalMetrics = snapshot.rowMetrics;
    const originalAxis = snapshot.columnAxis;
    const detachedWidths = snapshot.columnAxis.copyColumnWidths();
    detachedWidths[0] = 999;

    layout.replaceRoot(prepared);

    expect(layout.getRowHeight(0)).toBe(20);
    expect(layout.getColumnWidth(0)).toBe(50);
    expect(layout.root.rowMetrics).toBe(originalMetrics);
    expect(layout.root.columnAxis).toBe(originalAxis);
  });

  it('exposes root snapshots without mutable row index or hidden map handles', () => {
    const layout = createLayout();
    layout.hideColumn(1);
    const version = layout.version;
    const root = layout.root;

    expect(Object.isFrozen(root.rowMetrics)).toBe(true);
    const detachedHidden = root.columnAxis.copyHiddenColumnWidths();
    detachedHidden.set(1, 999);
    expect(root.columnAxis.getColumnRestoreWidth(1)).toBe(60);
    expect(Object.isFrozen(root.columnAxis.columnOffsets)).toBe(true);

    expect(readRowHeights(layout)).toEqual([20, 35, 20, 20]);
    expect(readColumnWidths(layout)).toEqual([50, 0, 70, 80]);
    expect(layout.version).toBe(version);
  });

  it('replaceRoot swaps row and column roots and bumps version exactly once', () => {
    const layout = createLayout();
    const prepared = layout.prepareMoveRows(1, 1, 4);
    const version = layout.version;
    const previousColumnWidths = layout.getColumnWidthsSnapshot();

    layout.replaceRoot(prepared);

    expect(readRowHeights(layout)).toEqual([20, 20, 20, 35]);
    expect(layout.getColumnWidthsSnapshot()).toEqual(previousColumnWidths);
    expect(layout.version).toBe(version + 1);
  });

  // 行高来源不再存在布局里：权威是 `rowData[row].hp`，`RowMetrics` 只留数字
  // （ADR 0004 阶段 4B T8，设计稿 §4.1/§4.2）。原先在这里的两条来源用例——
  // "来源与数字各自独立"和"来源跟着插/删/移重映射"——随 `getRowHeightProvenance`
  // 一起去掉了：前者已无对象可测，后者是 `rowData` 作为注册侧表的性质，由
  // `IndexedTable` 的重映射用例与 §4.2 的属性测试覆盖。

  it('prepare validation failures leave live state, hidden remaps, and version unchanged', () => {
    const layout = createLayout();
    layout.hideColumn(2);
    const version = layout.version;

    expect(() => layout.prepareInsertColumns(1, [Number.NaN])).toThrow(
      CyExcelError,
    );
    expect(() =>
      layout.prepareRoot({ ...layout.snapshotRoot(), rowCount: 1 }),
    ).toThrow(CyExcelError);
    expect(() => layout.prepareMoveRows(0, 2, 1)).toThrow(RangeError);

    expect(layout.rowCount).toBe(4);
    expect(layout.columnCount).toBe(4);
    expect(readRowHeights(layout)).toEqual([20, 35, 20, 20]);
    expect(readColumnWidths(layout)).toEqual([50, 60, 0, 80]);
    expect(layout.isColumnHidden(2)).toBe(true);
    expect(layout.getColumnRestoreWidth(2)).toBe(70);
    expect(layout.version).toBe(version);
  });
});

describe('LayoutEngine.prepareReorderRows', () => {
  // Asymmetric permutation on purpose: order[i] = order-inverse[i] would hold
  // for a self-inverse permutation like [1, 0], which can't distinguish a
  // correctly inverse-mapped row-height rope from an un-inverted one.
  function createLayout(): LayoutEngine {
    const layout = new LayoutEngine({
      rowCount: 3,
      columnCount: 1,
      rowHeight: 24,
      columnWidths: [100],
    });
    layout.setRowHeight(0, 60);
    return layout;
  }

  it('moves custom row heights along with their rows', () => {
    const layout = createLayout();

    layout.replaceRoot(layout.prepareReorderRows(Int32Array.from([2, 0, 1])));

    expect(layout.getRowHeight(1)).toBe(60);
    expect(layout.getRowHeight(0)).toBe(24);
    expect(layout.getRowHeight(2)).toBe(24);
  });

  it('keeps the total height unchanged', () => {
    const layout = createLayout();
    const before = layout.getTotalHeight();

    layout.replaceRoot(layout.prepareReorderRows(Int32Array.from([2, 0, 1])));

    expect(layout.getTotalHeight()).toBe(before);
  });

  it('does not mutate the current root', () => {
    const layout = createLayout();
    const version = layout.version;

    const prepared = layout.prepareReorderRows(Int32Array.from([2, 0, 1]));

    expect(layout.getRowHeight(0)).toBe(60);
    expect(layout.version).toBe(version);
    expect(prepared.version).toBe(version + 1);
  });

  it('rejects an order whose length does not match the row count', () => {
    const layout = createLayout();

    expect(() => layout.prepareReorderRows(Int32Array.from([1, 0]))).toThrow(RangeError);
  });
});
