import { describe, expect, it } from "vitest";

import {
  CellTypeRegistry,
  ColumnManager,
  registerBuiltinCellTypes
} from "../../../../src/core/cells";
import { buildCellGrid } from "../../../../src/core/clipboard";
import type { CellValue } from "../../../../src/core/sheet/types";

/** 46037 = 2026-01-15。日期在这个模型里就是序列号（见 cells/builtins.ts 的 "date"）。 */
const JAN_15_2026 = 46037;
const JAN_16_2026 = 46038;

function manager(): ColumnManager {
  const registry = new CellTypeRegistry();
  registerBuiltinCellTypes(registry);
  return new ColumnManager({
    columns: [
      { type: "text" },
      // display 与 raw 真正分岔的一列：`select` 的 format() 画标签，没有
      // serialize()，于是 raw 落到 String(value) 拿到原值。
      { type: "select", options: [{ value: "draft", label: "草稿" }] },
      { type: "date" },
      { type: "text", clipboard: { copy: false } }
    ],
    registry
  });
}

function values(): Record<string, CellValue[]> {
  return {
    name: ["Alice", "Bob"],
    status: ["draft", "published"],
    due: [JAN_15_2026, JAN_16_2026],
    secret: ["s1", "s2"]
  };
}

function contextFor(columnManager: ColumnManager) {
  const rows = values();
  const columnIds = ["name", "status", "due", "secret"] as const;
  return {
    getCellValue: (row: number, column: number) => rows[columnIds[column] as string]?.[row],
    getColumn: (column: number) => columnManager.getColumn(column)
  };
}

describe("buildCellGrid — display mode", () => {
  it("uses each column's format() output, not the stored value", () => {
    const columnManager = manager();
    const grid = buildCellGrid(
      { startRow: 0, startColumn: 0, endRow: 0, endColumn: 2 },
      "display",
      contextFor(columnManager)
    );

    expect(grid[0]?.[0]).toBe("Alice");
    expect(grid[0]?.[1]).toBe("草稿");
    expect(grid[0]?.[2]).toBe("2026-01-15");
  });
});

describe("buildCellGrid — raw mode", () => {
  it("uses the stored value rather than the display form", () => {
    const columnManager = manager();
    const grid = buildCellGrid(
      { startRow: 0, startColumn: 0, endRow: 0, endColumn: 2 },
      "raw",
      contextFor(columnManager)
    );

    expect(grid[0]?.[0]).toBe("Alice");
    // Premise: display mode draws "草稿" here, so raw mode really is a
    // different path and not the same string by coincidence.
    expect(grid[0]?.[1]).toBe("draft");
    expect(grid[0]?.[2]).toBe("2026-01-15");
  });

  it("falls back to String(value) for a type with no serialize()", () => {
    const registry = new CellTypeRegistry();
    registry.register("plain", { format: (v) => String(v ?? "") });
    const columnManager = new ColumnManager({ columns: [{ type: "plain" }], registry });

    const grid = buildCellGrid(
      { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 },
      "raw",
      { getCellValue: () => 42, getColumn: () => columnManager.getColumn(0) }
    );

    expect(grid[0]?.[0]).toBe("42");
  });
});

describe("buildCellGrid — structured mode", () => {
  it("includes the column's type name alongside the raw value", () => {
    const columnManager = manager();
    const grid = buildCellGrid(
      { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 },
      "structured",
      contextFor(columnManager)
    );

    expect(JSON.parse(grid[0]?.[0] ?? "")).toEqual({ type: "text", value: "Alice" });
  });
});

describe("buildCellGrid — column clipboard policy", () => {
  it("leaves a clipboard.copy:false column's cells blank but keeps its column position", () => {
    const columnManager = manager();
    const grid = buildCellGrid(
      { startRow: 0, startColumn: 0, endRow: 1, endColumn: 3 },
      "display",
      contextFor(columnManager)
    );

    expect(grid[0]?.length).toBe(4);
    expect(grid[0]?.[3]).toBe("");
    expect(grid[1]?.[3]).toBe("");
    // The other columns in the same rows are unaffected.
    expect(grid[0]?.[0]).toBe("Alice");
    expect(grid[1]?.[0]).toBe("Bob");
  });

  it("never reads the cell value of a clipboard.copy:false column", () => {
    const registry = new CellTypeRegistry();
    registerBuiltinCellTypes(registry);
    const columnManager = new ColumnManager({
      columns: [{ type: "text", clipboard: { copy: false } }],
      registry
    });

    let reads = 0;
    buildCellGrid(
      { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 },
      "display",
      {
        getCellValue: () => {
          reads += 1;
          return "value";
        },
        getColumn: () => columnManager.getColumn(0)
      }
    );

    expect(reads).toBe(0);
  });
});

describe("buildCellGrid — merge awareness", () => {
  it("serializes a covered-but-not-anchor cell as empty even though it has a stored value", () => {
    const columnManager = manager();
    // Merge spans (0,0)-(1,0): row 1 column 0 is covered but not the anchor.
    const merge = { startRow: 0, startColumn: 0, endRow: 1, endColumn: 0 };
    const grid = buildCellGrid(
      { startRow: 0, startColumn: 0, endRow: 1, endColumn: 0 },
      "display",
      {
        ...contextFor(columnManager),
        findMergeAt: (row, column) =>
          row >= merge.startRow && row <= merge.endRow && column >= merge.startColumn && column <= merge.endColumn
            ? merge
            : undefined
      }
    );

    expect(grid[0]?.[0]).toBe("Alice"); // anchor keeps its value
    expect(grid[1]?.[0]).toBe(""); // covered cell serializes to empty
  });

  it("does not call findMergeAt for cells outside any merge, and leaves them untouched", () => {
    const columnManager = manager();
    const calls: Array<[number, number]> = [];
    const grid = buildCellGrid(
      { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 },
      "display",
      {
        ...contextFor(columnManager),
        findMergeAt: (row, column) => {
          calls.push([row, column]);
          return undefined;
        }
      }
    );

    expect(calls).toEqual([[0, 0]]);
    expect(grid[0]?.[0]).toBe("Alice");
  });
});
