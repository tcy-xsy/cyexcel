import { describe, expect, it } from "vitest";

import { LayoutEngine, getSelectionBottomRightRect } from "../../../../src/core/layout";

function createLayout(freeze?: { rows?: number; columns?: number }): LayoutEngine {
  return new LayoutEngine({
    rowCount: 20,
    columnCount: 20,
    rowHeight: 20,
    columnWidths: Array.from({ length: 20 }, () => 50),
    ...(freeze ? { freeze } : {})
  });
}

describe("getSelectionBottomRightRect", () => {
  it("matches SelectionLayerRenderer's fill-handle rect for a plain single-cell selection", () => {
    const layout = createLayout();
    const viewport = { scrollTop: 0, scrollLeft: 0, width: 200, height: 200 };

    const rect = getSelectionBottomRightRect({ startRow: 3, startColumn: 2, endRow: 3, endColumn: 2 }, layout, viewport);

    // Absolute rect for (row 3, col 2): x = 2*50=100, y = 3*20=60.
    expect(rect).toEqual({ x: 100, y: 60, width: 50, height: 20 });
  });

  it("returns the bounding rect of a multi-cell range's last (bottom-right) quadrant", () => {
    const layout = createLayout();
    const viewport = { scrollTop: 0, scrollLeft: 0, width: 200, height: 200 };

    const rect = getSelectionBottomRightRect({ startRow: 1, startColumn: 1, endRow: 2, endColumn: 3 }, layout, viewport);

    expect(rect).toEqual({ x: 50, y: 20, width: 150, height: 40 });
  });

  it("uses the scrollable-quadrant rect (not the frozen one) when the selection straddles the frozen boundary", () => {
    const layout = createLayout({ rows: 1 });
    const viewport = { scrollTop: 100, scrollLeft: 0, width: 200, height: 200 };

    const rect = getSelectionBottomRightRect({ startRow: 0, startColumn: 0, endRow: 1, endColumn: 0 }, layout, viewport);

    // Scrollable row 1: absolute y=20, minus scrollTop 100 = -80.
    expect(rect).toEqual({ x: 0, y: -80, width: 50, height: 20 });
  });

  it("stays unshifted when the whole selection is inside the frozen corner", () => {
    const layout = createLayout({ rows: 2, columns: 2 });
    const viewport = { scrollTop: 500, scrollLeft: 500, width: 200, height: 200 };

    const rect = getSelectionBottomRightRect({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }, layout, viewport);

    expect(rect).toEqual({ x: 0, y: 0, width: 50, height: 20 });
  });
});
