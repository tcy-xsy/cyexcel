import { describe, expect, it } from "vitest";

import {
  CellTypeRegistry,
  ColumnManager,
  registerBuiltinCellTypes
} from "../../../../src/core/cells";
import { buildCellGrid } from "../../../../src/core/clipboard";

function manager(): ColumnManager {
  const registry = new CellTypeRegistry();
  registerBuiltinCellTypes(registry);
  return new ColumnManager({
    columns: [
      { type: "text" },
      { type: "array" },
      { type: "json" },
      { type: "text", clipboard: { copy: false } }
    ],
    registry
  });
}

function values(): Record<string, unknown[]> {
  return {
    name: ["Alice", "Bob"],
    tags: [
      ["red", "green", "blue", "yellow"],
      ["x"]
    ],
    meta: [{ a: 1, b: 2 }, { c: 3 }],
    secret: ["s1", "s2"]
  };
}

function contextFor(columnManager: ColumnManager) {
  const rows = values();
  const columnIds = ["name", "tags", "meta", "secret"] as const;
  return {
    getCellValue: (row: number, column: number) => rows[columnIds[column] as string]?.[row],
    getColumn: (column: number) => columnManager.getColumn(column)
  };
}

describe("buildCellGrid — display mode", () => {
  it("uses each column's format() output, which previews arrays/json instead of fully serializing them", () => {
    const columnManager = manager();
    const grid = buildCellGrid(
      { startRow: 0, startColumn: 0, endRow: 0, endColumn: 2 },
      "display",
      contextFor(columnManager)
    );

    expect(grid[0]?.[0]).toBe("Alice");
    expect(grid[0]?.[1]).toBe("[4] red, green, blue, ...");
    expect(grid[0]?.[2]).toBe('{"a":1,"b":2}');
  });
});

describe("buildCellGrid — raw mode", () => {
  it("uses each column's serialize() output, which fully serializes arrays/json", () => {
    const columnManager = manager();
    const grid = buildCellGrid(
      { startRow: 0, startColumn: 0, endRow: 0, endColumn: 2 },
      "raw",
      contextFor(columnManager)
    );

    expect(grid[0]?.[0]).toBe("Alice");
    expect(grid[0]?.[1]).toBe(JSON.stringify(["red", "green", "blue", "yellow"]));
    expect(grid[0]?.[2]).toBe(JSON.stringify({ a: 1, b: 2 }));
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
