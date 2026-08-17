import fc from "fast-check";
import { describe, expect, it, vi } from "vitest";

import {
  RowReorderController,
  isPointOnRowReorderHandle
} from "../../../../src/core/interaction";
import type { SelectionSnapshot } from "../../../../src/core/interaction";
import { LayoutEngine, SkeletonSnapshot, ViewportSnapshot } from "../../../../src/core/layout";
import type { ViewportState } from "../../../../src/core/layout";
import { SceneSnapshot } from "../../../../src/core/scene";
import type { CellRange } from "../../../../src/core/types";

function createLayout(rowCount = 20, frozenRows = 0) {
  return new LayoutEngine({
    rowCount,
    columnCount: 5,
    rowHeight: 20,
    columnWidths: Array.from({ length: 5 }, () => 50),
    freeze: { rows: frozenRows, columns: 0 }
  });
}

const VIEWPORT: ViewportState = { scrollTop: 0, scrollLeft: 0, width: 500, height: 400 };

function createScene(layout: LayoutEngine, viewport: ViewportState = VIEWPORT): SceneSnapshot {
  return new SceneSnapshot({
    version: 1,
    skeleton: SkeletonSnapshot.fromLayout(layout, 1),
    viewport: new ViewportSnapshot({ version: 1, ...viewport })
  });
}

const STRUCTURE_VERSION = 9;

function createController(onIndicatorChange?: (row: number | null) => void): RowReorderController {
  return new RowReorderController({ ...(onIndicatorChange ? { onIndicatorChange } : {}) });
}

function rowSnapshot(startRow: number, endRow: number, kind: SelectionSnapshot["kind"] = "row"): SelectionSnapshot {
  const range: CellRange = { startRow, startColumn: 0, endRow, endColumn: 4 };
  return {
    kind,
    active: { row: startRow, column: 0 },
    anchor: { row: startRow, column: 0 },
    range,
    ranges: [range]
  };
}

describe("isPointOnRowReorderHandle", () => {
  it("hits near the top edge of a single-row selection", () => {
    const layout = createLayout();
    // Row 3 spans y=[60,80). Top edge is y=60.
    expect(isPointOnRowReorderHandle(60, rowSnapshot(3, 3), createScene(layout))).toBe(true);
    expect(isPointOnRowReorderHandle(62, rowSnapshot(3, 3), createScene(layout))).toBe(true);
  });

  it("hits near the bottom edge of a multi-row selection", () => {
    const layout = createLayout();
    // Rows 3-5 span y=[60,120). Bottom edge is y=120.
    expect(isPointOnRowReorderHandle(120, rowSnapshot(3, 5), createScene(layout))).toBe(true);
    expect(isPointOnRowReorderHandle(118, rowSnapshot(3, 5), createScene(layout))).toBe(true);
  });

  it("misses a point well inside the selection, away from either edge", () => {
    const layout = createLayout();
    expect(isPointOnRowReorderHandle(90, rowSnapshot(3, 5), createScene(layout))).toBe(false);
  });

  it("misses when the active selection isn't row-kind", () => {
    const layout = createLayout();
    expect(isPointOnRowReorderHandle(60, rowSnapshot(3, 3, "cell"), createScene(layout))).toBe(false);
  });

  it("accounts for scroll when the boundary is outside the frozen band", () => {
    const layout = createLayout(20, 1); // 1 frozen row
    const scrolled: ViewportState = { ...VIEWPORT, scrollTop: 40 };
    // Row 5 top edge (absolute y=100) minus scrollTop(40) = viewport y 60.
    expect(isPointOnRowReorderHandle(60, rowSnapshot(5, 5), createScene(layout, scrolled))).toBe(true);
  });
});

describe("RowReorderController drag", () => {
  it("does not start a drag when pointerdown misses both edges", () => {
    const layout = createLayout();
    const scene = createScene(layout);
    const controller = createController();

    expect(controller.handlePointerDown(90, rowSnapshot(3, 3), scene, STRUCTURE_VERSION)).toBe(false);
    expect(controller.dragging).toBe(false);
  });

  it("starts a drag when pointerdown hits an edge, but stays non-dragging until the jitter threshold is exceeded", () => {
    const layout = createLayout();
    const scene = createScene(layout);
    const controller = createController();

    expect(controller.handlePointerDown(60, rowSnapshot(3, 3), scene, STRUCTURE_VERSION)).toBe(true);
    expect(controller.dragging).toBe(false); // hit, but no movement yet

    controller.handlePointerMove(61, scene, STRUCTURE_VERSION); // well within the jitter threshold
    expect(controller.dragging).toBe(false);

    controller.handlePointerMove(70, scene, STRUCTURE_VERSION); // past the threshold
    expect(controller.dragging).toBe(true);
  });

  it("resolves the indicator to the hovered row's index when hovering its top half (insert before)", () => {
    const layout = createLayout();
    const onIndicatorChange = vi.fn();
    const scene = createScene(layout);
    const controller = createController(onIndicatorChange);
    controller.handlePointerDown(20, rowSnapshot(0, 0), scene, STRUCTURE_VERSION);

    // Row 8 spans y=[160,180); top half -> hover y=165.
    controller.handlePointerMove(165, scene, STRUCTURE_VERSION);

    expect(controller.indicatorRow).toBe(8);
    expect(onIndicatorChange).toHaveBeenLastCalledWith(8);
  });

  it("resolves the indicator to hoveredRow+1 when hovering its bottom half (insert after)", () => {
    const layout = createLayout();
    const scene = createScene(layout);
    const controller = createController();
    controller.handlePointerDown(20, rowSnapshot(0, 0), scene, STRUCTURE_VERSION);

    // Row 8 spans y=[160,180); bottom half -> hover y=175.
    controller.handlePointerMove(175, scene, STRUCTURE_VERSION);

    expect(controller.indicatorRow).toBe(9);
  });

  it("shows no indicator while hovering strictly inside the dragged block", () => {
    const layout = createLayout();
    const scene = createScene(layout);
    const controller = createController();
    controller.handlePointerDown(40, rowSnapshot(2, 5), scene, STRUCTURE_VERSION);

    controller.handlePointerMove(80, scene, STRUCTURE_VERSION); // row 4, well inside [2,5]

    expect(controller.indicatorRow).toBeNull();
  });

  it("returns an immutable moveRows intent on pointerup for a valid target outside the block", () => {
    const layout = createLayout();
    const scene = createScene(layout);
    const controller = createController();
    controller.handlePointerDown(20, rowSnapshot(0, 0), scene, STRUCTURE_VERSION);
    controller.handlePointerMove(165, scene, STRUCTURE_VERSION);

    const intent = controller.handlePointerUp(STRUCTURE_VERSION);

    expect(intent).toEqual({ start: 0, count: 1, target: 8, structureVersion: STRUCTURE_VERSION });
    expect(Object.isFrozen(intent)).toBe(true);
    expect(controller.dragging).toBe(false);
    expect(controller.indicatorRow).toBeNull();
  });

  it("does not commit when the target lands on either endpoint of the source block (no-op)", () => {
    const layout = createLayout();
    const scene = createScene(layout);
    const controller = createController();
    controller.handlePointerDown(40, rowSnapshot(2, 4), scene, STRUCTURE_VERSION);

    // Hover row 2's top half -> resolves to boundary 2, the block's own start.
    controller.handlePointerMove(45, scene, STRUCTURE_VERSION);

    expect(controller.handlePointerUp(STRUCTURE_VERSION)).toBeNull();
  });

  it("does not commit when the pointer never leaves the block being dragged", () => {
    const layout = createLayout();
    const scene = createScene(layout);
    const controller = createController();
    controller.handlePointerDown(40, rowSnapshot(2, 5), scene, STRUCTURE_VERSION);
    controller.handlePointerMove(80, scene, STRUCTURE_VERSION); // strictly inside

    expect(controller.handlePointerUp(STRUCTURE_VERSION)).toBeNull();
  });

  it("cancel() aborts the drag without committing", () => {
    const layout = createLayout();
    const onIndicatorChange = vi.fn();
    const scene = createScene(layout);
    const controller = createController(onIndicatorChange);
    controller.handlePointerDown(20, rowSnapshot(0, 0), scene, STRUCTURE_VERSION);
    controller.handlePointerMove(165, scene, STRUCTURE_VERSION);

    controller.cancel();
    const intent = controller.handlePointerUp(STRUCTURE_VERSION);

    expect(intent).toBeNull();
    expect(controller.dragging).toBe(false);
    expect(controller.indicatorRow).toBeNull();
    expect(onIndicatorChange).toHaveBeenLastCalledWith(null);
  });

  it("ignores pointermove/pointerup/cancel when no drag is in progress", () => {
    const layout = createLayout();
    const scene = createScene(layout);
    const controller = createController();

    controller.handlePointerMove(165, scene, STRUCTURE_VERSION);
    expect(controller.handlePointerUp(STRUCTURE_VERSION)).toBeNull();
    controller.cancel();

    expect(controller.indicatorRow).toBeNull();
  });

  it("cancels and returns no intent when the structure generation changes", () => {
    const layout = createLayout();
    const scene = createScene(layout);
    const onIndicatorChange = vi.fn();
    const controller = createController(onIndicatorChange);
    controller.handlePointerDown(20, rowSnapshot(0, 0), scene, STRUCTURE_VERSION);
    controller.handlePointerMove(165, scene, STRUCTURE_VERSION);

    expect(controller.handlePointerUp(STRUCTURE_VERSION + 1)).toBeNull();
    expect(controller.indicatorRow).toBeNull();
    expect(controller.dragging).toBe(false);
    expect(onIndicatorChange).toHaveBeenLastCalledWith(null);
  });
});

describe("RowReorderController — property: the resolved target is always outside the dragged block when non-null", () => {
  it("holds for random (start, count, hoveredRow) triples", () => {
    const ROW_COUNT = 30;
    const layout = createLayout(ROW_COUNT);
    fc.assert(
      fc.property(
        fc.integer({ min: 0, max: ROW_COUNT - 1 }),
        fc.integer({ min: 1, max: ROW_COUNT }),
        fc.integer({ min: 0, max: ROW_COUNT - 1 }),
        fc.boolean(),
        (start, rawCount, hoveredRow, bottomHalf) => {
          const count = Math.min(rawCount, ROW_COUNT - start);
          fc.pre(count >= 1);
          const endRow = start + count - 1;

          const scene = createScene(layout);
          const controller = createController();
          controller.handlePointerDown(start * 20, rowSnapshot(start, endRow), scene, STRUCTURE_VERSION);
          const y = hoveredRow * 20 + (bottomHalf ? 15 : 5);
          controller.handlePointerMove(y, scene, STRUCTURE_VERSION);

          const target = controller.indicatorRow;
          if (target !== null) {
            expect(target <= start || target >= start + count).toBe(true);
          }
        }
      )
    );
  });
});
