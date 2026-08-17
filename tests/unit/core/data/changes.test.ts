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
});
