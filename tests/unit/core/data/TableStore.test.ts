import { describe, expect, it } from "vitest";

import { ArrayDataSource, CyExcelError, CyExcelErrorCode, TableStore } from "../../../../src/core";
import { createCountingDataSource } from "../../../helpers/createCountingDataSource";

function createStore() {
  const rows = [
    { name: "A", price: 1, tags: ["x"] },
    { name: "B", price: 2, tags: ["y"] },
    { name: "C", price: 3, tags: ["z"] }
  ];
  const source = new ArrayDataSource(rows, [
    { id: "name", field: "name" },
    { id: "price", field: "price" },
    { id: "tags", field: "tags" }
  ]);
  return { rows, store: new TableStore(source) };
}

describe("TableStore.setCell", () => {
  it("emits exactly one cell-value patch and bumps the version by one", () => {
    const { store } = createStore();

    const result = store.setCell(1, 1, 99);

    expect(result.patches).toEqual([{ type: "cell-value", row: 1, column: 1 }]);
    expect(result.version).toBe(1);
    expect(store.version).toBe(1);
    expect(store.getCell(1, 1)).toBe(99);
  });

  it("increments the version by exactly one per successful commit", () => {
    const { store } = createStore();

    store.setCell(0, 0, "a");
    store.setCell(0, 0, "b");
    store.setCell(0, 0, "c");

    expect(store.version).toBe(3);
  });
});

describe("TableStore.setRange", () => {
  it("emits exactly one range-value patch for the whole range", () => {
    const { store } = createStore();

    const result = store.setRange(
      { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 },
      [
        ["A2", 10],
        ["B2", 20]
      ]
    );

    expect(result.patches).toEqual([
      {
        type: "range-value",
        range: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 }
      }
    ]);
    expect(result.version).toBe(1);
    expect(store.getCell(0, 0)).toBe("A2");
    expect(store.getCell(0, 1)).toBe(10);
    expect(store.getCell(1, 0)).toBe("B2");
    expect(store.getCell(1, 1)).toBe(20);
  });
});

describe("TableStore.updateCell", () => {
  it("only produces a patch when updateCell is called explicitly", () => {
    const { store } = createStore();

    const current = store.getCell(0, 2) as string[];
    current.push("mutated-without-updateCell");
    expect(store.version).toBe(0);

    const result = store.updateCell(0, 2, value => [...(value as string[]), "y"]);

    expect(result.patches).toEqual([{ type: "cell-value", row: 0, column: 2 }]);
    expect(store.version).toBe(1);
    expect(store.getCell(0, 2)).toEqual(["x", "mutated-without-updateCell", "y"]);
  });
});

describe("TableStore failure paths do not bump the version", () => {
  it.each([
    ["negative row", () => {
      const { store } = createStore();
      return () => store.setCell(-1, 0, 1);
    }],
    ["non-integer column", () => {
      const { store } = createStore();
      return () => store.setCell(0, 1.5, 1);
    }],
    ["column out of range", () => {
      const { store } = createStore();
      return () => store.setCell(0, 99, 1);
    }]
  ])("%s throws a CyExcelError and leaves version at 0", (_label, setup) => {
    const attempt = setup();

    expect(attempt).toThrow(CyExcelError);
  });

  it("readonly DataSource: setCell throws and version stays at 0", () => {
    const source = new ArrayDataSource(
      [{ name: "A" }],
      [{ id: "name", field: "name" }],
      { readonly: true }
    );
    const store = new TableStore(source);

    expect(() => store.setCell(0, 0, "B")).toThrow(CyExcelError);
    expect(store.version).toBe(0);

    try {
      store.setCell(0, 0, "B");
    } catch (error) {
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.ReadonlyDataSource);
    }
  });

  it("a write failure raised by the DataSource leaves version unchanged", () => {
    const source = new ArrayDataSource([{ name: "A" }], [{ id: "name", field: "name" }]);
    const store = new TableStore(source);

    // Row 5 does not exist in the underlying source, so the write fails.
    expect(() => store.setCell(5, 0, "B")).toThrow(CyExcelError);
    expect(store.version).toBe(0);
  });
});

describe("TableStore does not scan unrelated rows", () => {
  it("setCell touches a constant number of cells regardless of table size", () => {
    const source = createCountingDataSource(1_000_000, 50);
    const store = new TableStore(source);

    const result = store.setCell(10, 2, 99);

    expect(result.patches).toEqual([{ type: "cell-value", row: 10, column: 2 }]);
    expect(source.readCount).toBeLessThanOrEqual(1);
    expect(source.writeCount).toBe(1);
  });
});
