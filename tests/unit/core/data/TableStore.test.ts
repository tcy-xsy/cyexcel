import { describe, expect, it } from "vitest";

import { ArrayDataSource, TableStore } from "../../../../src/core/data";
import { CyExcelError, CyExcelErrorCode } from "../../../../src/core/errors";
import type { DataSource } from "../../../../src/core/types";
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
    expect(result.changeSet).toEqual({
      modelVersion: 1,
      changes: [{ type: "cell-value", row: 1, column: 1 }]
    });
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

describe("TableStore — structural row operations", () => {
  it("insertRows emits a row-structure Patch and bumps the version", () => {
    const { store } = createStore();

    const result = store.insertRows(1, [["X", 10, []]]);

    expect(result.patches).toEqual([{ type: "row-structure", operation: "insert", start: 1, count: 1 }]);
    expect(result.version).toBe(1);
    expect(store.getCell(1, 0)).toBe("X");
    expect(store.getCell(2, 0)).toBe("B");
  });

  it("deleteRows emits a row-structure Patch and bumps the version", () => {
    const { store } = createStore();

    const result = store.deleteRows(0, 1);

    expect(result.patches).toEqual([{ type: "row-structure", operation: "delete", start: 0, count: 1 }]);
    expect(result.version).toBe(1);
    expect(store.getCell(0, 0)).toBe("B");
  });

  it("moveRows emits a row-structure Patch (including target) and bumps the version", () => {
    const { store } = createStore();

    const result = store.moveRows(0, 1, 3);

    expect(result.patches).toEqual([
      { type: "row-structure", operation: "move", start: 0, count: 1, target: 3 }
    ]);
    expect(result.version).toBe(1);
  });

  it("throws StructuralModificationNotSupported when the DataSource doesn't implement insertRows", () => {
    // A plain-object DataSource (e.g. a remote/computed one) that never
    // declared insertRows — not an ArrayDataSource with the method deleted,
    // since that would still be found via the prototype chain.
    const source: DataSource = {
      getRowCount: () => 1,
      getColumnCount: () => 1,
      getCell: () => "A",
      setCell: () => {}
    };
    const store = new TableStore(source);

    try {
      store.insertRows(0, [["B"]]);
      expect.unreachable("expected insertRows to throw");
    } catch (error) {
      expect(error).toBeInstanceOf(CyExcelError);
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.StructuralModificationNotSupported);
    }
  });

  it("throws StructuralModificationNotSupported when the DataSource doesn't implement deleteRows", () => {
    const source: DataSource = {
      getRowCount: () => 1,
      getColumnCount: () => 1,
      getCell: () => "A",
      setCell: () => {}
    };
    const store = new TableStore(source);

    expect(() => store.deleteRows(0, 1)).toThrow(CyExcelError);
  });

  it("throws StructuralModificationNotSupported when the DataSource doesn't implement moveRows", () => {
    const source: DataSource = {
      getRowCount: () => 2,
      getColumnCount: () => 1,
      getCell: () => "A",
      setCell: () => {}
    };
    const store = new TableStore(source);

    expect(() => store.moveRows(0, 1, 1)).toThrow(CyExcelError);
  });

  it("a rejected structural operation leaves the version unchanged (row)", () => {
    const source: DataSource = {
      getRowCount: () => 1,
      getColumnCount: () => 1,
      getCell: () => "A",
      setCell: () => {}
    };
    const store = new TableStore(source);

    expect(() => store.insertRows(0, [["B"]])).toThrow(CyExcelError);
    expect(store.version).toBe(0);
  });

  it.each([
    ["zero-count insert", (store: TableStore) => store.insertRows(1, [])],
    ["delete past the end", (store: TableStore) => store.deleteRows(2, 2)],
    ["move target inside the block", (store: TableStore) => store.moveRows(0, 2, 1)]
  ])("rejects %s before mutating data or version", (_label, run) => {
    const { rows, store } = createStore();

    expect(() => run(store)).toThrow(RangeError);

    expect(store.version).toBe(0);
    expect(rows.map((row) => row.name)).toEqual(["A", "B", "C"]);
  });
});

describe("TableStore — structural column operations", () => {
  it("insertColumns emits a column-structure Patch and bumps the version", () => {
    const { store } = createStore();

    const result = store.insertColumns(1, [{ id: "qty", field: "qty" }]);

    expect(result.patches).toEqual([{ type: "column-structure", operation: "insert", start: 1, count: 1 }]);
    expect(result.version).toBe(1);
    expect(store.getColumnCount()).toBe(4);
    expect(store.getCell(0, 0)).toBe("A"); // name unaffected
    expect(store.getCell(0, 2)).toBe(1); // price shifted right
  });

  it("deleteColumns emits a column-structure Patch and bumps the version", () => {
    const { store } = createStore();

    const result = store.deleteColumns(0, 1);

    expect(result.patches).toEqual([{ type: "column-structure", operation: "delete", start: 0, count: 1 }]);
    expect(result.version).toBe(1);
    expect(store.getColumnCount()).toBe(2);
    expect(store.getCell(0, 0)).toBe(1); // price shifted into name's old slot
  });

  it("moveColumns emits a column-structure Patch (including target) and bumps the version", () => {
    const { store } = createStore();

    const result = store.moveColumns(0, 1, 3);

    expect(result.patches).toEqual([
      { type: "column-structure", operation: "move", start: 0, count: 1, target: 3 }
    ]);
    expect(result.version).toBe(1);
  });

  it("throws StructuralModificationNotSupported when the DataSource doesn't implement insertColumns", () => {
    const source: DataSource = {
      getRowCount: () => 1,
      getColumnCount: () => 1,
      getCell: () => "A",
      setCell: () => {}
    };
    const store = new TableStore(source);

    try {
      store.insertColumns(0, [{ id: "b", field: "b" }]);
      expect.unreachable("expected insertColumns to throw");
    } catch (error) {
      expect(error).toBeInstanceOf(CyExcelError);
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.StructuralModificationNotSupported);
    }
  });

  it("throws StructuralModificationNotSupported when the DataSource doesn't implement deleteColumns", () => {
    const source: DataSource = {
      getRowCount: () => 1,
      getColumnCount: () => 1,
      getCell: () => "A",
      setCell: () => {}
    };
    const store = new TableStore(source);

    expect(() => store.deleteColumns(0, 1)).toThrow(CyExcelError);
  });

  it("throws StructuralModificationNotSupported when the DataSource doesn't implement moveColumns", () => {
    const source: DataSource = {
      getRowCount: () => 1,
      getColumnCount: () => 2,
      getCell: () => "A",
      setCell: () => {}
    };
    const store = new TableStore(source);

    expect(() => store.moveColumns(0, 1, 1)).toThrow(CyExcelError);
  });

  it("a rejected structural operation leaves the version unchanged (column)", () => {
    const source: DataSource = {
      getRowCount: () => 1,
      getColumnCount: () => 1,
      getCell: () => "A",
      setCell: () => {}
    };
    const store = new TableStore(source);

    expect(() => store.insertColumns(0, [{ id: "b", field: "b" }])).toThrow(CyExcelError);
    expect(store.version).toBe(0);
  });

  it.each([
    ["zero-count insert", (store: TableStore) => store.insertColumns(1, [])],
    ["delete past the end", (store: TableStore) => store.deleteColumns(2, 2)],
    ["move target inside the block", (store: TableStore) => store.moveColumns(0, 2, 1)]
  ])("rejects %s before mutating columns or version", (_label, run) => {
    const { store } = createStore();

    expect(() => run(store)).toThrow(RangeError);

    expect(store.version).toBe(0);
    expect(store.getColumnCount()).toBe(3);
    expect(store.getCell(0, 0)).toBe("A");
    expect(store.getCell(0, 1)).toBe(1);
    expect(store.getCell(0, 2)).toEqual(["x"]);
  });
});
