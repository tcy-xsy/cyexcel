import { describe, expect, it } from "vitest";

import { compileDirtySet, createChangeSet, summarizeStructuralChanges } from "../../../../src/core/data";
import type { SemanticChange } from "../../../../src/core/data";

describe("ChangeSet", () => {
  it("preserves duplicate structural changes in their original order", () => {
    const first = { type: "row-structure" as const, operation: "insert" as const, start: 2, count: 1 };
    const second = { type: "row-structure" as const, operation: "insert" as const, start: 2, count: 1 };

    const changes = createChangeSet(7, [first, second]);

    expect(changes.modelVersion).toBe(7);
    expect(changes.changes).toEqual([first, second]);
    expect(changes.changes).toHaveLength(2);
    expect(Object.isFrozen(changes.changes)).toBe(true);
  });

  it("deep-snapshots the rows and range of a row-layout fact", () => {
    const rows: [number, ...number[]] = [1, 3];
    const range = { startRow: 1, startColumn: 0, endRow: 8, endColumn: 4 };
    const changes = createChangeSet(2, [{ type: "row-layout", rows, range }]);

    rows.push(5);
    range.startRow = 7;

    expect(changes.changes).toEqual([{
      type: "row-layout",
      rows: [1, 3],
      range: { startRow: 1, startColumn: 0, endRow: 8, endColumn: 4 }
    }]);
    const rowLayout = changes.changes[0];
    expect(rowLayout?.type).toBe("row-layout");
    if (rowLayout?.type !== "row-layout") {
      throw new Error("expected row-layout change");
    }
    expect(Object.isFrozen(rowLayout)).toBe(true);
    expect(Object.isFrozen(rowLayout.rows)).toBe(true);
    expect(Object.isFrozen(rowLayout.range)).toBe(true);
  });

  it("deep-snapshots the order of a row-reorder fact", () => {
    const order = Int32Array.from([1, 0, 2]);
    const changes = createChangeSet(3, [{ type: "row-reorder", order }]);

    order[0] = 99;

    expect(changes.changes).toEqual([{ type: "row-reorder", order: Int32Array.from([1, 0, 2]) }]);
    const rowReorder = changes.changes[0];
    expect(rowReorder?.type).toBe("row-reorder");
    if (rowReorder?.type !== "row-reorder") {
      throw new Error("expected row-reorder change");
    }
    expect(Object.isFrozen(rowReorder)).toBe(true);
  });
});

describe("compileDirtySet", () => {
  it("derives one full repaint from repeated structure facts without deleting those facts", () => {
    const dirty = compileDirtySet([
      { type: "row-structure", operation: "insert", start: 2, count: 1 },
      { type: "row-structure", operation: "insert", start: 2, count: 1 }
    ]);

    expect(dirty).toEqual({
      content: "full",
      skeleton: "row-axis",
      selection: true,
      presentation: { domCells: true, columnHeader: false, rowHeader: true, domHeaders: false }
    });
  });

  it("keeps local cell changes as independently derived dirty regions", () => {
    const dirty = compileDirtySet([
      { type: "cell-value", row: 1, column: 2 },
      { type: "range-value", range: { startRow: 3, startColumn: 4, endRow: 5, endColumn: 6 } }
    ]);

    expect(dirty).toEqual({
      content: [
        { range: { startRow: 1, startColumn: 2, endRow: 1, endColumn: 2 } },
        { range: { startRow: 3, startColumn: 4, endRow: 5, endColumn: 6 } }
      ],
      skeleton: "none",
      selection: false,
      presentation: { domCells: true, columnHeader: false, rowHeader: false, domHeaders: false }
    });
  });

  it("derives one shifted-tail region from one batched row-layout fact", () => {
    const dirty = compileDirtySet([{
      type: "row-layout",
      rows: [2, 5],
      range: { startRow: 2, startColumn: 0, endRow: 9, endColumn: 3 }
    }]);

    expect(dirty).toEqual({
      content: [{ range: { startRow: 2, startColumn: 0, endRow: 9, endColumn: 3 } }],
      skeleton: "row-axis",
      selection: false,
      presentation: { domCells: true, columnHeader: false, rowHeader: true, domHeaders: false }
    });
  });

  it("reports no dirty content when nothing touched a cell", () => {
    // "none" used to be unreachable: an empty region list was returned as `[]`,
    // so the type advertised a case the function never produced.
    expect(compileDirtySet([])).toEqual({
      content: "none",
      skeleton: "none",
      selection: false,
      presentation: { domCells: false, columnHeader: false, rowHeader: false, domHeaders: false }
    });
  });

  it("treats a row-reorder like a row-structure fact, including the row-header repaint", () => {
    // Row numbers are positional and do not themselves change, but each row's
    // height travels with it under a reorder, so band boundaries move and the
    // gutter must repaint (doc/scene-mutation-migration-plan.md §14.7).
    const dirty = compileDirtySet([{ type: "row-reorder", order: Int32Array.from([1, 0]) }]);

    expect(dirty).toEqual({
      content: "full",
      skeleton: "row-axis",
      selection: true,
      presentation: { domCells: true, columnHeader: false, rowHeader: true, domHeaders: false }
    });
  });

  it("treats a row-visibility fact like a row-reorder, including the row-header repaint", () => {
    // A filter apply/clear collapses/restores row heights across the sheet
    // without changing any row's identity — same shape of consequence as a
    // reorder (band boundaries move, content/gutter/selection all repaint),
    // not a structural axis change.
    const dirty = compileDirtySet([{ type: "row-visibility" }]);

    expect(dirty).toEqual({
      content: "full",
      skeleton: "row-axis",
      selection: true,
      presentation: { domCells: true, columnHeader: false, rowHeader: true, domHeaders: false }
    });
  });

  it("marks only the content layer dirty for a validation-layout fact", () => {
    // 规则只影响"无效"角标画不画：不动几何、表头、选区，也不需要重排 DOM 单元格
    const dirty = compileDirtySet([{ type: "validation-layout" }]);

    expect(dirty).toEqual({
      content: "full",
      skeleton: "none",
      selection: false,
      presentation: { domCells: false, columnHeader: false, rowHeader: false, domHeaders: false }
    });
  });

  it("marks content and selection dirty for a merge-layout fact, but leaves headers untouched", () => {
    // Narrower than column-layout: merges don't move row/column geometry, so
    // unlike the column-axis facts below, the header bands stay clean.
    const dirty = compileDirtySet([{ type: "merge-layout" }]);

    expect(dirty).toEqual({
      content: "full",
      skeleton: "none",
      selection: true,
      presentation: { domCells: true, columnHeader: false, rowHeader: false, domHeaders: false }
    });
  });

  it("escalates every band for a column-axis fact, matching the old per-patch invalidateAll", () => {
    const columnAxisFacts: readonly SemanticChange[] = [
      { type: "column-structure", operation: "insert", start: 0, count: 1 },
      { type: "column-layout", column: 0 },
      { type: "full-reset" }
    ];
    for (const change of columnAxisFacts) {
      const dirty = compileDirtySet([change]);
      expect(dirty.content).toBe("full");
      expect(dirty.presentation).toEqual({
        domCells: true,
        columnHeader: true,
        rowHeader: true,
        domHeaders: true
      });
    }
  });
});

describe("summarizeStructuralChanges", () => {
  it("reports nothing for a change list with no axis facts", () => {
    expect(
      summarizeStructuralChanges([
        { type: "cell-value", row: 0, column: 0 },
        { type: "range-value", range: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 } }
      ])
    ).toEqual({
      hasRowStructure: false,
      hasColumnStructure: false,
      hasColumnLayout: false,
      hasRowLayout: false
    });
  });

  it("reports each axis fact independently, and all of them together", () => {
    expect(
      summarizeStructuralChanges([
        { type: "row-structure", operation: "insert", start: 0, count: 1 },
        { type: "column-structure", operation: "delete", start: 1, count: 2 },
        { type: "column-layout", column: 3 },
        {
          type: "row-layout",
          rows: [4],
          range: { startRow: 4, startColumn: 0, endRow: 9, endColumn: 2 }
        }
      ])
    ).toEqual({
      hasRowStructure: true,
      hasColumnStructure: true,
      hasColumnLayout: true,
      hasRowLayout: true
    });
  });

  it("stays false for an empty list", () => {
    expect(summarizeStructuralChanges([])).toEqual({
      hasRowStructure: false,
      hasColumnStructure: false,
      hasColumnLayout: false,
      hasRowLayout: false
    });
  });

  it("counts a row-visibility fact as row-layout, not row-structure", () => {
    // No row was inserted/deleted/moved, so #structure.apply() (gated on
    // hasRowStructure/hasColumnStructure) must not run for a filter change.
    expect(summarizeStructuralChanges([{ type: "row-visibility" }])).toEqual({
      hasRowStructure: false,
      hasColumnStructure: false,
      hasColumnLayout: false,
      hasRowLayout: true
    });
  });

  it("counts a row-reorder as a row-structure fact", () => {
    expect(
      summarizeStructuralChanges([{ type: "row-reorder", order: Int32Array.from([1, 0]) }])
    ).toEqual({
      hasRowStructure: true,
      hasColumnStructure: false,
      hasColumnLayout: false,
      hasRowLayout: false
    });
  });
});
