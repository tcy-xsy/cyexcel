import { describe, expect, it } from "vitest";

import { LayoutEngine, classifyContextMenuTarget } from "../../../../src/core";
import type { SelectionSnapshot, ViewportState } from "../../../../src/core";

const selection: SelectionSnapshot = {
  kind: "cell",
  active: { row: 0, column: 0 },
  anchor: { row: 0, column: 0 },
  range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }
};

function layout(): LayoutEngine {
  return new LayoutEngine({
    rowCount: 10,
    columnCount: 5,
    rowHeight: 20,
    columnWidths: [50, 50, 50, 50, 50]
  });
}

const viewport: ViewportState = { scrollTop: 0, scrollLeft: 0, width: 300, height: 300 };
const headerLayout = { rowHeaderWidth: 40, columnHeaderHeight: 24 };

describe("classifyContextMenuTarget", () => {
  it("classifies a point in the top-left corner band as 'corner'", () => {
    const context = classifyContextMenuTarget({ x: 10, y: 10 }, viewport, headerLayout, layout(), selection);
    expect(context.target).toBe("corner");
    expect(context.row).toBeUndefined();
    expect(context.column).toBeUndefined();
  });

  it("classifies a point in the left strip (below the corner) as 'row-header' and resolves the row", () => {
    const context = classifyContextMenuTarget({ x: 10, y: 24 + 45 }, viewport, headerLayout, layout(), selection);
    expect(context.target).toBe("row-header");
    expect(context.row).toBe(2); // (45) / 20 = row 2
    expect(context.column).toBeUndefined();
  });

  it("classifies a point in the top strip (right of the corner) as 'column-header' and resolves the column", () => {
    const context = classifyContextMenuTarget({ x: 40 + 55, y: 10 }, viewport, headerLayout, layout(), selection);
    expect(context.target).toBe("column-header");
    expect(context.column).toBe(1); // 55 / 50 = column 1
    expect(context.row).toBeUndefined();
  });

  it("classifies a point inside the grid as 'cell' and resolves both row and column", () => {
    const context = classifyContextMenuTarget(
      { x: 40 + 55, y: 24 + 45 },
      viewport,
      headerLayout,
      layout(),
      selection
    );
    expect(context.target).toBe("cell");
    expect(context.row).toBe(2);
    expect(context.column).toBe(1);
  });

  it("classifies a point past the data extent as 'blank'", () => {
    // 10 rows * 20px = 200px tall; well past that but still inside the grid area.
    const context = classifyContextMenuTarget({ x: 40 + 10, y: 24 + 500 }, viewport, headerLayout, layout(), selection);
    expect(context.target).toBe("blank");
    expect(context.row).toBeUndefined();
    expect(context.column).toBeUndefined();
  });

  it("always carries the selection snapshot and the raw point through, regardless of target", () => {
    const context = classifyContextMenuTarget({ x: 5, y: 5 }, viewport, headerLayout, layout(), selection);
    expect(context.selection).toBe(selection);
    expect(context.point).toEqual({ x: 5, y: 5 });
  });

  it("treats zero-size row/column headers as if there were no header bands at all", () => {
    const context = classifyContextMenuTarget(
      { x: 55, y: 45 },
      viewport,
      { rowHeaderWidth: 0, columnHeaderHeight: 0 },
      layout(),
      selection
    );
    expect(context.target).toBe("cell");
    expect(context.row).toBe(2);
    expect(context.column).toBe(1);
  });
});
