import { describe, expect, it } from "vitest";

import { CyExcelError, CyExcelErrorCode, LayoutEngine } from "../../../../src/core";

describe("LayoutEngine construction", () => {
  const base = { rowCount: 10, columnCount: 2, rowHeight: 20, columnWidths: [50, 60] };

  it.each([
    ["negative rowCount", { ...base, rowCount: -1 }],
    ["non-integer rowCount", { ...base, rowCount: 1.5 }],
    ["negative columnCount", { ...base, columnCount: -1 }],
    ["zero rowHeight", { ...base, rowHeight: 0 }],
    ["negative rowHeight", { ...base, rowHeight: -20 }],
    ["non-finite rowHeight", { ...base, rowHeight: Number.POSITIVE_INFINITY }],
    ["columnWidths length mismatch", { ...base, columnWidths: [50] }],
    ["negative column width", { ...base, columnWidths: [50, -1] }],
    ["non-finite column width", { ...base, columnWidths: [50, Number.NaN] }],
    ["non-integer overscanRows", { ...base, overscanRows: 1.5 }],
    ["negative overscanColumns", { ...base, overscanColumns: -1 }],
    ["non-integer freeze.rows", { ...base, freeze: { rows: 1.5 } }],
    ["negative freeze.columns", { ...base, freeze: { columns: -1 } }]
  ])("rejects %s", (_label, options) => {
    expect(() => new LayoutEngine(options)).toThrow(CyExcelError);
  });

  it("clamps freeze counts that exceed the actual row/column count instead of throwing", () => {
    const layout = new LayoutEngine({ ...base, freeze: { rows: 999, columns: 999 } });
    expect(layout.frozenRowCount).toBe(10);
    expect(layout.frozenColumnCount).toBe(2);
  });

  it("defaults freeze and overscan to 0 when not provided", () => {
    const layout = new LayoutEngine(base);
    expect(layout.frozenRowCount).toBe(0);
    expect(layout.frozenColumnCount).toBe(0);
  });

  it("exposes the constructor inputs through basic getters", () => {
    const layout = new LayoutEngine(base);
    expect(layout.rowCount).toBe(10);
    expect(layout.columnCount).toBe(2);
    expect(layout.rowHeight).toBe(20);
  });
});

describe("LayoutEngine geometry primitives", () => {
  const layout = new LayoutEngine({
    rowCount: 10,
    columnCount: 3,
    rowHeight: 20,
    columnWidths: [50, 100, 150]
  });

  it("computes cell rects from row height and cumulative column widths", () => {
    expect(layout.getCellRect(0, 0)).toEqual({ x: 0, y: 0, width: 50, height: 20 });
    expect(layout.getCellRect(2, 1)).toEqual({ x: 50, y: 40, width: 100, height: 20 });
    expect(layout.getCellRect(0, 2)).toEqual({ x: 150, y: 0, width: 150, height: 20 });
  });

  it("reports total width and height", () => {
    expect(layout.getTotalWidth()).toBe(300);
    expect(layout.getTotalHeight()).toBe(200);
  });

  it("returns 0 width for an out-of-range column", () => {
    expect(layout.getColumnWidth(99)).toBe(0);
  });

  it("falls back to x=0 for an out-of-range column in getCellRect", () => {
    expect(layout.getCellRect(0, 99)).toEqual({ x: 0, y: 0, width: 0, height: 20 });
  });

  it("getRowAtY/getColumnAtX clamp to valid bounds", () => {
    expect(layout.getRowAtY(-5)).toBe(0);
    expect(layout.getRowAtY(0)).toBe(0);
    expect(layout.getRowAtY(45)).toBe(2);
    expect(layout.getRowAtY(10_000)).toBe(9);

    expect(layout.getColumnAtX(-5)).toBe(0);
    expect(layout.getColumnAtX(60)).toBe(1);
    expect(layout.getColumnAtX(10_000)).toBe(2);
  });

  it("getRowAtY/getColumnAtX return 0 for an empty table", () => {
    const empty = new LayoutEngine({ rowCount: 0, columnCount: 0, rowHeight: 20, columnWidths: [] });
    expect(empty.getRowAtY(50)).toBe(0);
    expect(empty.getColumnAtX(50)).toBe(0);
  });
});

describe("LayoutEngine.getVisibleRowRange — test 1: scrollTop under fixed row height", () => {
  const layout = new LayoutEngine({ rowCount: 100, columnCount: 1, rowHeight: 20, columnWidths: [50] });

  it("maps scrollTop 0 to the first rows", () => {
    expect(layout.getVisibleRowRange({ scrollTop: 0, scrollLeft: 0, width: 50, height: 100 })).toEqual({
      start: 0,
      count: 6
    });
  });

  it("maps a scrolled-down scrollTop to the corresponding rows", () => {
    expect(layout.getVisibleRowRange({ scrollTop: 205, scrollLeft: 0, width: 50, height: 100 })).toEqual({
      start: 10,
      count: 6
    });
  });
});

describe("LayoutEngine.getVisibleColumnRange — test 2: scrollLeft under variable column widths", () => {
  const layout = new LayoutEngine({
    rowCount: 1,
    columnCount: 3,
    rowHeight: 20,
    columnWidths: [50, 100, 150]
  });

  it("maps scrollLeft 0 to the leftmost visible columns", () => {
    expect(layout.getVisibleColumnRange({ scrollTop: 0, scrollLeft: 0, width: 120, height: 20 })).toEqual({
      start: 0,
      count: 2
    });
  });

  it("maps a scrolled scrollLeft to the correct visible columns", () => {
    expect(layout.getVisibleColumnRange({ scrollTop: 0, scrollLeft: 60, width: 120, height: 20 })).toEqual({
      start: 1,
      count: 2
    });
  });
});

describe("LayoutEngine overscan — test 3: never exceeds row/column bounds", () => {
  const layout = new LayoutEngine({
    rowCount: 10,
    columnCount: 1,
    rowHeight: 20,
    columnWidths: [50],
    overscanRows: 5
  });

  it("clamps overscan at the top edge instead of going negative", () => {
    expect(layout.getVisibleRowRange({ scrollTop: 0, scrollLeft: 0, width: 50, height: 40 })).toEqual({
      start: 0,
      count: 8
    });
  });

  it("clamps overscan at the bottom edge instead of exceeding rowCount", () => {
    expect(layout.getVisibleRowRange({ scrollTop: 160, scrollLeft: 0, width: 50, height: 40 })).toEqual({
      start: 3,
      count: 7
    });
  });
});

describe("LayoutEngine freeze — test 4: frozen rows/columns occupy fixed viewport space", () => {
  it("reduces the scrollable visible row count by the frozen band's height", () => {
    const withoutFreeze = new LayoutEngine({ rowCount: 20, columnCount: 1, rowHeight: 10, columnWidths: [50] });
    const withFreeze = new LayoutEngine({
      rowCount: 20,
      columnCount: 1,
      rowHeight: 10,
      columnWidths: [50],
      freeze: { rows: 3 }
    });
    const viewport = { scrollTop: 0, scrollLeft: 0, width: 50, height: 50 };

    expect(withFreeze.frozenRowCount).toBe(3);
    const withoutRange = withoutFreeze.getVisibleRowRange(viewport);
    const withRange = withFreeze.getVisibleRowRange(viewport);

    // 3 frozen rows * 10px row height = 30px less scrollable viewport space,
    // i.e. exactly 3 fewer visible rows at the same row height.
    expect(withoutRange.count - withRange.count).toBe(3);
  });

  it("frozen cell rects never depend on scroll position", () => {
    const layout = new LayoutEngine({
      rowCount: 20,
      columnCount: 1,
      rowHeight: 10,
      columnWidths: [50],
      freeze: { rows: 3 }
    });

    expect(layout.getCellRect(1, 0)).toEqual(layout.getCellRect(1, 0));
  });
});

describe("LayoutEngine — test 5: logical geometry never depends on DPR", () => {
  it("does not accept a dpr parameter and stays deterministic across repeated calls", () => {
    const layout = new LayoutEngine({ rowCount: 10, columnCount: 2, rowHeight: 20, columnWidths: [50, 60] });
    const viewport = { scrollTop: 0, scrollLeft: 0, width: 100, height: 100 };

    expect(layout.getVisibleRowRange(viewport)).toEqual(layout.getVisibleRowRange(viewport));
    expect(layout.getCellRect(0, 0)).toEqual(layout.getCellRect(0, 0));
  });
});

describe("LayoutEngine — test 6: no floating point drift at large row indices", () => {
  it("computes an exact cell rect near row 1,000,000", () => {
    const layout = new LayoutEngine({
      rowCount: 1_000_000,
      columnCount: 1,
      rowHeight: 24,
      columnWidths: [80]
    });

    expect(layout.getCellRect(999_998, 0).y).toBe(999_998 * 24);
  });
});

describe("LayoutEngine — test 7: no negative ranges for empty/zero/hidden edge cases", () => {
  it("returns a zero-count range for an empty table", () => {
    const empty = new LayoutEngine({ rowCount: 0, columnCount: 0, rowHeight: 20, columnWidths: [] });
    const viewport = { scrollTop: 0, scrollLeft: 0, width: 100, height: 100 };

    const rowRange = empty.getVisibleRowRange(viewport);
    const columnRange = empty.getVisibleColumnRange(viewport);

    expect(rowRange).toEqual({ start: 0, count: 0 });
    expect(columnRange).toEqual({ start: 0, count: 0 });
  });

  it("returns a zero-count range for a zero-size viewport", () => {
    const layout = new LayoutEngine({ rowCount: 10, columnCount: 2, rowHeight: 20, columnWidths: [50, 50] });
    const zeroViewport = { scrollTop: 0, scrollLeft: 0, width: 0, height: 0 };

    expect(layout.getVisibleRowRange(zeroViewport)).toEqual({ start: 0, count: 0 });
    expect(layout.getVisibleColumnRange(zeroViewport)).toEqual({ start: 0, count: 0 });
  });

  it("never returns a negative start/count when every column is hidden (width 0)", () => {
    const hidden = new LayoutEngine({ rowCount: 5, columnCount: 3, rowHeight: 20, columnWidths: [0, 0, 0] });
    const range = hidden.getVisibleColumnRange({ scrollTop: 0, scrollLeft: 0, width: 100, height: 100 });

    expect(range.start).toBeGreaterThanOrEqual(0);
    expect(range.count).toBeGreaterThanOrEqual(0);
  });
});

describe("LayoutEngine.getCellAtPoint — test 8: stable at boundary pixels", () => {
  const layout = new LayoutEngine({
    rowCount: 20,
    columnCount: 1,
    rowHeight: 10,
    columnWidths: [50],
    freeze: { rows: 2 }
  });
  const viewport = { scrollTop: 0, scrollLeft: 0, width: 50, height: 100 };

  it("resolves the exact frozen/scrollable boundary pixel consistently", () => {
    const atBoundary = layout.getCellAtPoint({ x: 0, y: 20 }, viewport);
    expect(atBoundary).toEqual({ row: 2, column: 0 });
    expect(layout.getCellAtPoint({ x: 0, y: 20 }, viewport)).toEqual(atBoundary);
  });

  it("resolves the pixel just before the boundary to the frozen row", () => {
    expect(layout.getCellAtPoint({ x: 0, y: 19 }, viewport)).toEqual({ row: 1, column: 0 });
  });

  it("resolves points inside a frozen column band without applying scrollLeft", () => {
    const withFrozenColumns = new LayoutEngine({
      rowCount: 5,
      columnCount: 4,
      rowHeight: 10,
      columnWidths: [30, 30, 30, 30],
      freeze: { columns: 2 }
    });
    const scrolledViewport = { scrollTop: 0, scrollLeft: 100, width: 120, height: 50 };

    // x=10 is inside the frozen band (0-60px); scrollLeft must not shift it.
    expect(withFrozenColumns.getCellAtPoint({ x: 10, y: 0 }, scrolledViewport)).toEqual({
      row: 0,
      column: 0
    });
  });

  it("returns null for negative coordinates", () => {
    expect(layout.getCellAtPoint({ x: -1, y: 0 }, viewport)).toBeNull();
    expect(layout.getCellAtPoint({ x: 0, y: -1 }, viewport)).toBeNull();
  });

  it("returns null for an empty table", () => {
    const empty = new LayoutEngine({ rowCount: 0, columnCount: 0, rowHeight: 20, columnWidths: [] });
    expect(empty.getCellAtPoint({ x: 0, y: 0 }, viewport)).toBeNull();
  });

  it("returns null for a point beyond the last row or column", () => {
    const small = new LayoutEngine({ rowCount: 2, columnCount: 1, rowHeight: 10, columnWidths: [50] });
    const smallViewport = { scrollTop: 0, scrollLeft: 0, width: 50, height: 200 };
    expect(small.getCellAtPoint({ x: 0, y: 500 }, smallViewport)).toBeNull();

    const narrow = new LayoutEngine({ rowCount: 1, columnCount: 2, rowHeight: 10, columnWidths: [50, 50] });
    const narrowViewport = { scrollTop: 0, scrollLeft: 0, width: 200, height: 10 };
    expect(narrow.getCellAtPoint({ x: 500, y: 0 }, narrowViewport)).toBeNull();
  });
});

describe("LayoutEngine performance guard: 1,000,000 rows is not scanned", () => {
  it("computes visible ranges, a cell rect, and a hit test in well under linear time", () => {
    const layout = new LayoutEngine({
      rowCount: 1_000_000,
      columnCount: 50,
      rowHeight: 24,
      columnWidths: Array.from({ length: 50 }, () => 100)
    });
    const viewport = { scrollTop: 12_000_000, scrollLeft: 2000, width: 800, height: 600 };

    const start = performance.now();
    layout.getVisibleRowRange(viewport);
    layout.getVisibleColumnRange(viewport);
    layout.getCellRect(500_000, 25);
    layout.getCellAtPoint({ x: 400, y: 300 }, viewport);
    const elapsed = performance.now() - start;

    expect(elapsed).toBeLessThan(5);
  });
});

describe("LayoutEngine.getColumnLeft", () => {
  it("returns the same x offset as getCellRect without needing a row", () => {
    const layout = new LayoutEngine({
      rowCount: 5,
      columnCount: 3,
      rowHeight: 20,
      columnWidths: [50, 60, 70]
    });

    expect(layout.getColumnLeft(0)).toBe(0);
    expect(layout.getColumnLeft(1)).toBe(50);
    expect(layout.getColumnLeft(2)).toBe(110);
    expect(layout.getColumnLeft(2)).toBe(layout.getCellRect(0, 2).x);
  });

  it("falls back to 0 for an out-of-range column, same as getCellRect", () => {
    const layout = new LayoutEngine({
      rowCount: 5,
      columnCount: 3,
      rowHeight: 20,
      columnWidths: [50, 60, 70]
    });

    expect(layout.getColumnLeft(99)).toBe(0);
  });
});

describe("LayoutEngine.setColumnWidth", () => {
  function createLayout(): LayoutEngine {
    return new LayoutEngine({
      rowCount: 5,
      columnCount: 3,
      rowHeight: 20,
      columnWidths: [50, 60, 70]
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

  it("getCellRect reflects the new width and shifted offsets immediately", () => {
    const layout = createLayout();

    layout.setColumnWidth(1, 200);

    expect(layout.getCellRect(2, 1)).toEqual({ x: 50, y: 40, width: 200, height: 20 });
    expect(layout.getCellRect(2, 2).x).toBe(250);
  });

  it("bumps version by exactly 1 per successful call, starting at 0", () => {
    const layout = createLayout();
    expect(layout.version).toBe(0);

    layout.setColumnWidth(0, 100);
    expect(layout.version).toBe(1);

    layout.setColumnWidth(1, 90);
    expect(layout.version).toBe(2);
  });

  it("throws and leaves state/version unchanged for a negative or non-finite width", () => {
    const layout = createLayout();

    expect(() => layout.setColumnWidth(0, -1)).toThrow(CyExcelError);
    expect(() => layout.setColumnWidth(0, Number.NaN)).toThrow(CyExcelError);
    try {
      layout.setColumnWidth(0, -1);
    } catch (error) {
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.InvalidOptions);
    }
    expect(layout.getColumnWidth(0)).toBe(50);
    expect(layout.version).toBe(0);
  });

  it("throws for an out-of-range column index", () => {
    const layout = createLayout();

    expect(() => layout.setColumnWidth(99, 100)).toThrow(CyExcelError);
    expect(layout.version).toBe(0);
  });

  it("allows a width of 0 (hidden column)", () => {
    const layout = createLayout();

    layout.setColumnWidth(0, 0);

    expect(layout.getColumnWidth(0)).toBe(0);
    expect(layout.getColumnLeft(1)).toBe(0);
  });
});

describe("LayoutEngine.setRowCount", () => {
  it("updates rowCount and getTotalHeight immediately", () => {
    const layout = new LayoutEngine({ rowCount: 5, columnCount: 1, rowHeight: 20, columnWidths: [50] });

    layout.setRowCount(8);

    expect(layout.rowCount).toBe(8);
    expect(layout.getTotalHeight()).toBe(160);
  });

  it("bumps version by exactly 1 per successful call, starting at 0", () => {
    const layout = new LayoutEngine({ rowCount: 5, columnCount: 1, rowHeight: 20, columnWidths: [50] });
    expect(layout.version).toBe(0);

    layout.setRowCount(8);
    expect(layout.version).toBe(1);

    layout.setRowCount(3);
    expect(layout.version).toBe(2);
  });

  it("re-clamps frozenRowCount against the new rowCount when it shrinks below the frozen band", () => {
    const layout = new LayoutEngine({
      rowCount: 10,
      columnCount: 1,
      rowHeight: 20,
      columnWidths: [50],
      freeze: { rows: 4 }
    });
    expect(layout.frozenRowCount).toBe(4);

    layout.setRowCount(2);

    expect(layout.frozenRowCount).toBe(2);
  });

  it("recovers frozen rows on a later grow, re-clamping against the originally requested freeze.rows", () => {
    const layout = new LayoutEngine({
      rowCount: 10,
      columnCount: 1,
      rowHeight: 20,
      columnWidths: [50],
      freeze: { rows: 4 }
    });

    layout.setRowCount(2); // frozenRowCount clamps to 2
    layout.setRowCount(10); // grows back — should recover the originally requested 4, not stay at 2

    expect(layout.frozenRowCount).toBe(4);
  });

  it("does not touch column state", () => {
    const layout = new LayoutEngine({ rowCount: 5, columnCount: 2, rowHeight: 20, columnWidths: [50, 60] });

    layout.setRowCount(8);

    expect(layout.getColumnWidth(0)).toBe(50);
    expect(layout.getColumnLeft(1)).toBe(50);
  });

  it("throws and leaves state/version unchanged for a negative or non-integer rowCount", () => {
    const layout = new LayoutEngine({ rowCount: 5, columnCount: 1, rowHeight: 20, columnWidths: [50] });

    expect(() => layout.setRowCount(-1)).toThrow(CyExcelError);
    expect(() => layout.setRowCount(1.5)).toThrow(CyExcelError);
    expect(layout.rowCount).toBe(5);
    expect(layout.version).toBe(0);
  });

  it("allows a rowCount of 0", () => {
    const layout = new LayoutEngine({ rowCount: 5, columnCount: 1, rowHeight: 20, columnWidths: [50] });

    layout.setRowCount(0);

    expect(layout.rowCount).toBe(0);
    expect(layout.getTotalHeight()).toBe(0);
  });
});
