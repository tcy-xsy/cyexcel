import { describe, expect, it, vi } from "vitest";

import type { CanvasCellRenderer, RuntimeColumn } from "../../../../src/core/cells";
import { LayoutEngine, SkeletonSnapshot, ViewportSnapshot } from "../../../../src/core/layout";
import { WrapMeasurementService } from "../../../../src/core/measurement";
import { SceneSnapshot } from "../../../../src/core/scene";
import { RecordingCanvasContext } from "../../../helpers/RecordingCanvasContext";
import { provenanceViewOf } from "../../../helpers/rowHeightAuthority";
import type { CellValue } from "../../../../src/core/sheet/types";

// RecordingCanvasContext measures a deterministic 7px per character, and the
// theme wraps at CELL_LINE_HEIGHT 16 + CELL_PADDING 4 on each side. So a
// 100px column fits floor((100 - 8) / 7) = 13 characters per line, and an
// n-line cell needs n * 16 + 8 px.
const DEFAULT_ROW_HEIGHT = 24;
const lineHeightFor = (lines: number): number => lines * 16 + 8;

/** A column that paints itself, so `wrap` must be ignored for it. */
const customCanvasRenderer: CanvasCellRenderer = {
  kind: "canvas",
  paint: () => undefined
};

function textColumn(overrides: Partial<RuntimeColumn> = {}): RuntimeColumn {
  return {
    
    
    width: 100,
    typeName: "text",
    format: (value) => (value === undefined || value === null ? "" : String(value)),
    clipboard: { copy: true, paste: true },
    wrap: false,
    sortable: true,
    filterable: true,
    version: 1,
    ...overrides
  };
}

function createLayout(rowCount = 10, columnWidths = [100]): LayoutEngine {
  return new LayoutEngine({
    rowCount,
    columnCount: columnWidths.length,
    rowHeight: DEFAULT_ROW_HEIGHT,
    columnWidths
  });
}

function createScene(layout: LayoutEngine, modelVersion = 1, height = 48): SceneSnapshot {
  return new SceneSnapshot({
    version: 1,
    skeleton: SkeletonSnapshot.fromLayout(layout, modelVersion),
    viewport: new ViewportSnapshot({
      version: 1,
      scrollTop: 0,
      scrollLeft: 0,
      width: 100,
      height
    })
  });
}

interface Harness {
  readonly service: WrapMeasurementService;
  readonly onError: ReturnType<typeof vi.fn>;
  readonly layout: LayoutEngine;
  readonly scene: SceneSnapshot;
}

function createService(options: {
  readonly columns?: readonly RuntimeColumn[];
  readonly cells?: ReadonlyMap<string, CellValue>;
  readonly layout?: LayoutEngine;
  readonly sceneHeight?: number;
} = {}): Harness {
  const columns = options.columns ?? [textColumn({ wrap: true })];
  const cells = options.cells ?? new Map<string, CellValue>();
  const layout = options.layout ?? createLayout();
  const onError = vi.fn();
  const service = new WrapMeasurementService({
    measureContext: new RecordingCanvasContext(),
    getColumns: () => columns,
    getCell: (row, column) => cells.get(`${row}:${column}`),
    onError
  });
  return { service, onError, layout, scene: createScene(layout, 1, options.sceneHeight) };
}

function basisFor(harness: Harness, provenance = new Map<number, "auto" | "manual">()) {
  return {
    scene: harness.scene,
    modelVersion: harness.scene.modelVersion,
    layoutVersion: harness.scene.layoutVersion,
    defaultRowHeight: DEFAULT_ROW_HEIGHT,
    rowHeightProvenance: provenanceViewOf(provenance)
  };
}

describe("WrapMeasurementService.hasWrapColumns", () => {
  it("is false when no column opted into wrap", () => {
    const { service } = createService({ columns: [textColumn()] });
    expect(service.hasWrapColumns()).toBe(false);
  });

  it("is true for a wrap column on the built-in paint path", () => {
    const { service } = createService({ columns: [textColumn({ wrap: true })] });
    expect(service.hasWrapColumns()).toBe(true);
  });

  it("ignores a wrap column that owns its own renderer", () => {
    const { service } = createService({
      columns: [textColumn({ wrap: true, renderer: customCanvasRenderer })]
    });
    expect(service.hasWrapColumns()).toBe(false);
  });
});

describe("WrapMeasurementService.computeRequiredRowHeight", () => {
  it("returns the default height when content fits on one line", () => {
    const harness = createService({ cells: new Map([["0:0", "short"]]) });
    expect(harness.service.computeRequiredRowHeight(0, harness.scene, DEFAULT_ROW_HEIGHT))
      .toBe(DEFAULT_ROW_HEIGHT);
  });

  it("grows to fit wrapped content", () => {
    // 26 characters at 13 per line = 2 lines.
    const harness = createService({ cells: new Map([["0:0", "aaaaaaaaaaaa bbbbbbbbbbbb"]]) });
    expect(harness.service.computeRequiredRowHeight(0, harness.scene, DEFAULT_ROW_HEIGHT))
      .toBe(lineHeightFor(2));
  });

  it("takes the tallest column, not the last one", () => {
    const harness = createService({
      columns: [
        textColumn({ wrap: true }),
        textColumn({ wrap: true })
      ],
      cells: new Map<string, CellValue>([
        ["0:0", "aaaaaaaaaaaa bbbbbbbbbbbb cccccccccccc"],
        ["0:1", "short"]
      ]),
      layout: createLayout(10, [100, 100])
    });
    expect(harness.service.computeRequiredRowHeight(0, harness.scene, DEFAULT_ROW_HEIGHT))
      .toBe(lineHeightFor(3));
  });

  it("skips columns owned by a custom renderer", () => {
    const harness = createService({
      columns: [textColumn({ wrap: true, renderer: customCanvasRenderer })],
      cells: new Map([["0:0", "aaaaaaaaaaaa bbbbbbbbbbbb"]])
    });
    expect(harness.service.computeRequiredRowHeight(0, harness.scene, DEFAULT_ROW_HEIGHT))
      .toBe(DEFAULT_ROW_HEIGHT);
  });

  it("skips a column too narrow to hold any text", () => {
    const harness = createService({
      columns: [textColumn({ wrap: true, width: 8 })],
      cells: new Map([["0:0", "aaaaaaaaaaaa bbbbbbbbbbbb"]]),
      layout: createLayout(10, [8])
    });
    expect(harness.service.computeRequiredRowHeight(0, harness.scene, DEFAULT_ROW_HEIGHT))
      .toBe(DEFAULT_ROW_HEIGHT);
  });
});

describe("WrapMeasurementService cache", () => {
  it("reports visible rows once and then treats them as measured", () => {
    const harness = createService({ sceneHeight: 48 });
    const first = harness.service.visibleUnmeasuredRows(harness.scene);
    expect(first.length).toBeGreaterThan(0);

    harness.service.markMeasured(first);
    expect(harness.service.visibleUnmeasuredRows(harness.scene)).toEqual([]);
  });

  it("forgetRow re-exposes exactly one row", () => {
    const harness = createService();
    const rows = harness.service.visibleUnmeasuredRows(harness.scene);
    harness.service.markMeasured(rows);

    harness.service.forgetRow(rows[0] as number);
    expect(harness.service.visibleUnmeasuredRows(harness.scene)).toEqual([rows[0]]);
  });

  it("forgetRowRange re-exposes only rows inside the span", () => {
    const harness = createService({ sceneHeight: 240 });
    const rows = harness.service.visibleUnmeasuredRows(harness.scene);
    harness.service.markMeasured(rows);

    harness.service.forgetRowRange(1, 2);
    expect(harness.service.visibleUnmeasuredRows(harness.scene)).toEqual([1, 2]);
  });

  it("forgetAll drops the whole epoch", () => {
    const harness = createService();
    const rows = harness.service.visibleUnmeasuredRows(harness.scene);
    harness.service.markMeasured(rows);

    harness.service.forgetAll();
    expect(harness.service.measuredRowCount).toBe(0);
    expect(harness.service.visibleUnmeasuredRows(harness.scene)).toEqual([...rows]);
  });
});

describe("WrapMeasurementService.planUpdates", () => {
  it("rejects a Scene that is not the current model/layout basis", () => {
    const harness = createService();
    expect(() =>
      harness.service.planUpdates(new Set([0]), {
        ...basisFor(harness),
        modelVersion: harness.scene.modelVersion + 1
      })
    ).toThrow(/does not match model\/layout/);
  });

  it("reports a resize for a row whose content no longer fits", () => {
    const harness = createService({ cells: new Map([["0:0", "aaaaaaaaaaaa bbbbbbbbbbbb"]]) });
    const plan = harness.service.planUpdates(new Set([0]), basisFor(harness));

    expect(plan.updates).toEqual([{ row: 0, height: lineHeightFor(2) }]);
    expect(plan.measuredWithoutResize).toEqual([]);
  });

  it("marks a row measured when it already has the right height, without an update", () => {
    const harness = createService({ cells: new Map([["0:0", "short"]]) });
    const plan = harness.service.planUpdates(new Set([0]), basisFor(harness));

    expect(plan.updates).toEqual([]);
    expect(plan.measuredWithoutResize).toEqual([0]);
    expect(harness.service.measuredRowCount).toBe(1);
  });

  it("never measures a manually resized row", () => {
    const harness = createService({ cells: new Map([["0:0", "aaaaaaaaaaaa bbbbbbbbbbbb"]]) });
    const plan = harness.service.planUpdates(
      new Set([0]),
      basisFor(harness, new Map([[0, "manual"]]))
    );

    expect(plan.updates).toEqual([]);
    expect(plan.measuredWithoutResize).toEqual([0]);
  });

  it("reports a throwing formatter and keeps measuring the other rows", () => {
    const harness = createService({
      columns: [
        textColumn({
          wrap: true,
          format: (_value, context) => {
            if (context.row === 0) {
              throw new Error("bad formatter");
            }
            return "aaaaaaaaaaaa bbbbbbbbbbbb";
          }
        })
      ]
    });
    const plan = harness.service.planUpdates(new Set([0, 1]), basisFor(harness));

    expect(harness.onError).toHaveBeenCalledTimes(1);
    expect(harness.onError.mock.calls[0]?.[1]).toBe("finalize:wrap-measurement");
    expect(plan.updates).toEqual([{ row: 1, height: lineHeightFor(2) }]);
  });

  it("freezes the plan so a caller cannot mutate measured geometry", () => {
    const harness = createService({ cells: new Map([["0:0", "aaaaaaaaaaaa bbbbbbbbbbbb"]]) });
    const plan = harness.service.planUpdates(new Set([0]), basisFor(harness));

    expect(Object.isFrozen(plan)).toBe(true);
    expect(Object.isFrozen(plan.updates)).toBe(true);
    expect(Object.isFrozen(plan.updates[0])).toBe(true);
  });
});

describe("WrapMeasurementService.rowsNeedingDefaultReset", () => {
  it("finds auto-owned rows still taller than the default", () => {
    const layout = createLayout();
    layout.setRowHeight(2, 60);
    const harness = createService({ layout });

    expect(
      harness.service.rowsNeedingDefaultReset(
        harness.scene,
        provenanceViewOf(new Map([[2, "auto"]])),
        DEFAULT_ROW_HEIGHT
      )
    ).toEqual([2]);
  });

  it("leaves manually resized rows alone", () => {
    const layout = createLayout();
    layout.setRowHeight(2, 60);
    const harness = createService({ layout });

    expect(
      harness.service.rowsNeedingDefaultReset(
        harness.scene,
        provenanceViewOf(new Map([[2, "manual"]])),
        DEFAULT_ROW_HEIGHT
      )
    ).toEqual([]);
  });
});

describe("WrapMeasurementService.invalidateFor", () => {
  function measuredHarness(): Harness {
    const harness = createService({ sceneHeight: 240 });
    harness.service.markMeasured(harness.service.visibleUnmeasuredRows(harness.scene));
    return harness;
  }

  it("forgets exactly the edited row for a cell-value fact", () => {
    const harness = measuredHarness();
    harness.service.invalidateFor([{ type: "cell-value", row: 2, column: 0 }]);
    expect(harness.service.visibleUnmeasuredRows(harness.scene)).toEqual([2]);
  });

  it("forgets the whole span for a range-value fact", () => {
    const harness = measuredHarness();
    harness.service.invalidateFor([
      { type: "range-value", range: { startRow: 1, startColumn: 0, endRow: 3, endColumn: 0 } }
    ]);
    expect(harness.service.visibleUnmeasuredRows(harness.scene)).toEqual([1, 2, 3]);
  });

  it("forgets the whole epoch for any axis-level fact", () => {
    for (const change of [
      { type: "row-structure", operation: "insert", start: 0, count: 1 },
      { type: "column-structure", operation: "delete", start: 0, count: 1 },
      { type: "column-layout", column: 0 },
      { type: "full-reset" }
    ] as const) {
      const harness = measuredHarness();
      expect(harness.service.measuredRowCount).toBeGreaterThan(0);
      harness.service.invalidateFor([change]);
      expect(harness.service.measuredRowCount).toBe(0);
    }
  });

  it("forgets every measured row after a reorder", () => {
    const harness = measuredHarness();
    expect(harness.service.measuredRowCount).toBeGreaterThan(0);

    harness.service.invalidateFor([{ type: "row-reorder", order: Int32Array.from([1, 0]) }]);

    expect(harness.service.measuredRowCount).toBe(0);
  });

  it("leaves the cache alone when nothing in the list touches it", () => {
    // 合并区变化不改变任何一行里文字要占多高，所以不该丢测量结果。
    // （这条原本用的是 `cell-style`——Step 24 之前它是个从未被产出过的占位变体；
    // 现在样式改动会让测量失效，见下面两条。）
    const harness = measuredHarness();
    const before = harness.service.measuredRowCount;
    harness.service.invalidateFor([{ type: "merge-layout" }]);
    expect(harness.service.measuredRowCount).toBe(before);
  });

  it("forgets the row when its style changed — 改字号会改变文字占的高度", () => {
    const harness = measuredHarness();
    const before = harness.service.measuredRowCount;
    harness.service.invalidateFor([{ type: "cell-style", row: 0, column: 0 }]);
    expect(harness.service.measuredRowCount).toBe(before - 1);
  });

  it("forgets the whole range when a range's style changed", () => {
    const harness = measuredHarness();
    const before = harness.service.measuredRowCount;
    harness.service.invalidateFor([
      { type: "range-style", range: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 0 } }
    ]);
    expect(harness.service.measuredRowCount).toBe(before - 2);
  });
});

describe("WrapMeasurementService.rowsToMeasureFor", () => {
  const noReset = {
    hasColumnStructureChange: false,
    rowHeightProvenance: provenanceViewOf(new Map()),
    defaultRowHeight: DEFAULT_ROW_HEIGHT
  };

  it("returns the visible unmeasured rows when a wrap column exists", () => {
    const harness = createService({ sceneHeight: 240 });
    const result = harness.service.rowsToMeasureFor(harness.scene, noReset);

    expect(result.rows.size).toBeGreaterThan(0);
    expect(result.visibleUnmeasuredCount).toBe(result.rows.size);
  });

  it("skips the visible scan entirely when nothing wraps", () => {
    const harness = createService({ columns: [textColumn()], sceneHeight: 240 });
    const result = harness.service.rowsToMeasureFor(harness.scene, noReset);

    expect(result.rows.size).toBe(0);
    expect(result.visibleUnmeasuredCount).toBe(0);
  });

  it("still finds auto-owned rows to shrink after the last wrap column is deleted", () => {
    // The half that is easy to forget: these rows are already marked measured,
    // so the visible scan can never surface them, and without this branch they
    // would keep their wrapped height forever.
    const layout = createLayout();
    layout.setRowHeight(2, 60);
    const harness = createService({ columns: [textColumn()], layout, sceneHeight: 240 });

    const result = harness.service.rowsToMeasureFor(harness.scene, {
      hasColumnStructureChange: true,
      rowHeightProvenance: provenanceViewOf(new Map([[2, "auto"]])),
      defaultRowHeight: DEFAULT_ROW_HEIGHT
    });

    expect([...result.rows]).toEqual([2]);
    // Not from the visible scan, so it must not extend the follow-up pass.
    expect(result.visibleUnmeasuredCount).toBe(0);
  });

  it("does not look for shrink candidates while a wrap column still exists", () => {
    const layout = createLayout();
    layout.setRowHeight(2, 60);
    const harness = createService({ layout, sceneHeight: 240 });

    const result = harness.service.rowsToMeasureFor(harness.scene, {
      hasColumnStructureChange: true,
      rowHeightProvenance: provenanceViewOf(new Map([[2, "auto"]])),
      defaultRowHeight: DEFAULT_ROW_HEIGHT
    });

    expect(result.visibleUnmeasuredCount).toBe(result.rows.size);
  });

  it("reports the visible count separately from the merged set", () => {
    const layout = createLayout();
    layout.setRowHeight(7, 60);
    const harness = createService({ columns: [textColumn()], layout, sceneHeight: 48 });

    const result = harness.service.rowsToMeasureFor(harness.scene, {
      hasColumnStructureChange: true,
      rowHeightProvenance: provenanceViewOf(new Map([[7, "auto"]])),
      defaultRowHeight: DEFAULT_ROW_HEIGHT
    });

    // Row 7 is off-screen: it belongs in the merged set but must not make the
    // caller schedule another visible-measurement pass.
    expect(result.rows.has(7)).toBe(true);
    expect(result.visibleUnmeasuredCount).toBe(0);
  });
});

describe("WrapMeasurementService measured-row cache bound", () => {
  // Far above any plausible cap: a 4K-tall viewport at the 16px minimum row
  // height shows ~135 rows, so a cap has to be a few orders of magnitude below
  // this to still leave the current screenful cached.
  const OVERFLOW = 50_000;
  const SCREENFUL = 200;

  function markRange(service: WrapMeasurementService, start: number, end: number): void {
    const rows: number[] = [];
    for (let row = start; row <= end; row += 1) {
      rows.push(row);
    }
    service.markMeasured(rows);
  }

  it("caps the cache instead of growing with every row ever made visible", () => {
    const { service } = createService();

    markRange(service, 0, OVERFLOW - 1);

    // Scrolling a wrap-enabled table through a million rows must not retain a
    // million cache entries — Milestone 0 requires memory not to grow with
    // time spent scrolling.
    expect(service.measuredRowCount).toBeLessThan(OVERFLOW);
  });

  it("evicts the oldest rows first, so the most recently measured screenful survives", () => {
    const { service } = createService();

    markRange(service, 0, OVERFLOW - 1);

    // forgetRowRange only deletes rows that are actually cached, so the drop in
    // measuredRowCount reports precisely how many of that span survived.
    const beforeNewest = service.measuredRowCount;
    service.forgetRowRange(OVERFLOW - SCREENFUL, OVERFLOW - 1);
    expect(beforeNewest - service.measuredRowCount).toBe(SCREENFUL);

    const beforeOldest = service.measuredRowCount;
    service.forgetRowRange(0, SCREENFUL - 1);
    expect(beforeOldest - service.measuredRowCount).toBe(0);
  });
});


it("does not enqueue filtered rows for wrap measurement", () => {
  const layout = createLayout(100_000);
  layout.replaceRoot(layout.prepareSetFilteredRows(new Set(
    Array.from({ length: 99_998 }, (_, index) => index + 1)
  )));
  const { service, scene } = createService({ layout, sceneHeight: 100 });
  expect(service.visibleUnmeasuredRows(scene)).toEqual([0, 99_999]);
});
