import { describe, expect, it } from "vitest";

import type { SelectionSnapshot } from "../../../../src/core/interaction";
import { LayoutEngine, SkeletonSnapshot, ViewportSnapshot } from "../../../../src/core/layout";
import type { ViewportState } from "../../../../src/core/layout";
import { classifyContextMenuTarget } from "../../../../src/core/menu";
import { SceneSnapshot, type SceneSurfaceMetrics } from "../../../../src/core/scene";

const selection: SelectionSnapshot = {
  kind: "cell",
  active: { row: 0, column: 0 },
  anchor: { row: 0, column: 0 },
  range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 },
  ranges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }]
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

function scene(
  layout: LayoutEngine,
  viewportState: ViewportState = viewport,
  surfaceMetrics: SceneSurfaceMetrics = { rowHeaderWidth: 40, columnHeaderHeight: 24 }
): SceneSnapshot {
  return new SceneSnapshot({
    version: 1,
    skeleton: SkeletonSnapshot.fromLayout(layout, 1),
    viewport: new ViewportSnapshot({ version: 1, ...viewportState }),
    surfaceMetrics
  });
}

describe("classifyContextMenuTarget", () => {
  it("classifies a point in the top-left corner band as 'corner'", () => {
    const context = classifyContextMenuTarget({ x: 10, y: 10 }, scene(layout()), selection);
    expect(context.target).toBe("corner");
    expect(context.row).toBeUndefined();
    expect(context.column).toBeUndefined();
  });

  it("classifies a point in the left strip (below the corner) as 'row-header' and resolves the row", () => {
    const context = classifyContextMenuTarget({ x: 10, y: 24 + 45 }, scene(layout()), selection);
    expect(context.target).toBe("row-header");
    expect(context.row).toBe(2); // (45) / 20 = row 2
    expect(context.column).toBeUndefined();
  });

  it("classifies a point in the top strip (right of the corner) as 'column-header' and resolves the column", () => {
    const context = classifyContextMenuTarget({ x: 40 + 55, y: 10 }, scene(layout()), selection);
    expect(context.target).toBe("column-header");
    expect(context.column).toBe(1); // 55 / 50 = column 1
    expect(context.row).toBeUndefined();
  });

  it("classifies a point inside the grid as 'cell' and resolves both row and column", () => {
    const context = classifyContextMenuTarget(
      { x: 40 + 55, y: 24 + 45 },
      scene(layout()),
      selection
    );
    expect(context.target).toBe("cell");
    expect(context.row).toBe(2);
    expect(context.column).toBe(1);
  });

  it("classifies a point past the data extent as 'blank'", () => {
    // 10 rows * 20px = 200px tall; well past that but still inside the grid area.
    const context = classifyContextMenuTarget({ x: 40 + 10, y: 24 + 500 }, scene(layout()), selection);
    expect(context.target).toBe("blank");
    expect(context.row).toBeUndefined();
    expect(context.column).toBeUndefined();
  });

  it("always carries the selection snapshot and the raw point through, regardless of target", () => {
    const context = classifyContextMenuTarget({ x: 5, y: 5 }, scene(layout()), selection);
    expect(context.selection).toBe(selection);
    expect(context.point).toEqual({ x: 5, y: 5 });
  });

  it("treats zero-size row/column headers as if there were no header bands at all", () => {
    const context = classifyContextMenuTarget(
      { x: 55, y: 45 },
      scene(layout(), viewport, { rowHeaderWidth: 0, columnHeaderHeight: 0 }),
      selection
    );
    expect(context.target).toBe("cell");
    expect(context.row).toBe(2);
    expect(context.column).toBe(1);
  });
});
