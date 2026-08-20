import { describe, expect, it } from "vitest";

import { ArrayDataSource } from "../../../../src/core/data";
import { CyExcelError, CyExcelErrorCode } from "../../../../src/core/errors";
import type { ModelRevision } from "../../../../src/core/model";

describe("ArrayDataSource", () => {
  it("reads and writes by numeric index through a stable column field", () => {
    const rows = [
      { name: "A", price: 1 },
      { name: "B", price: 2 }
    ];
    const source = new ArrayDataSource(rows, [
      { id: "name", field: "name" },
      { id: "price", field: "price" }
    ]);

    expect(source.getRowCount()).toBe(2);
    expect(source.getColumnCount()).toBe(2);
    expect(source.getCell(0, 0)).toBe("A");
    expect(source.getCell(1, 1)).toBe(2);

    source.setCell(1, 1, 20);
    expect(source.getCell(1, 1)).toBe(20);
    expect(rows[1]?.price).toBe(20);
  });

  it("keeps complex values as the same reference instead of proxying them", () => {
    const rows = [{ name: "A", tags: ["x"] }];
    const source = new ArrayDataSource(rows, [
      { id: "name", field: "name" },
      { id: "tags", field: "tags" }
    ]);
    const tags = ["a", "b"];

    source.setCell(0, 1, tags);

    expect(source.getCell(0, 1)).toBe(tags);

    tags.push("c");
    expect(source.getCell(0, 1)).toBe(tags);
    expect((source.getCell(0, 1) as string[]).length).toBe(3);
  });

  it("returns undefined when reading a column index that does not exist", () => {
    const source = new ArrayDataSource([{ name: "A" }], [{ id: "name", field: "name" }]);

    expect(source.getCell(0, 5)).toBeUndefined();
  });

  it("throws a stable error code when writing to a column index that does not exist", () => {
    const source = new ArrayDataSource([{ name: "A" }], [{ id: "name", field: "name" }]);

    try {
      source.setCell(0, 5, "value");
      expect.unreachable("expected setCell to throw for an unknown column");
    } catch (error) {
      expect(error).toBeInstanceOf(CyExcelError);
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.ColumnNotFound);
    }
  });

  it("throws a stable error code when writing to a row index that does not exist", () => {
    const source = new ArrayDataSource([{ name: "A" }], [{ id: "name", field: "name" }]);

    try {
      source.setCell(5, 0, "value");
      expect.unreachable("expected setCell to throw for an unknown row");
    } catch (error) {
      expect(error).toBeInstanceOf(CyExcelError);
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.InvalidCoordinate);
    }
  });

  it("throws a stable error code when writing to a readonly source", () => {
    const source = new ArrayDataSource(
      [{ name: "A" }],
      [{ id: "name", field: "name" }],
      { readonly: true }
    );

    try {
      source.setCell(0, 0, "B");
      expect.unreachable("expected setCell to throw on a readonly source");
    } catch (error) {
      expect(error).toBeInstanceOf(CyExcelError);
      expect((error as CyExcelError).code).toBe(
        CyExcelErrorCode.ReadonlyDataSource
      );
    }
  });
});

describe("ArrayDataSource — insertRows", () => {
  function columns() {
    return [
      { id: "name", field: "name" },
      { id: "price", field: "price" }
    ];
  }

  it("inserts column-ordered value rows at the given position, converted via this source's own field mapping", () => {
    const source = new ArrayDataSource([{ name: "A", price: 1 }, { name: "C", price: 3 }], columns());

    source.insertRows(1, [["B", 2]]);

    expect(source.getRowCount()).toBe(3);
    expect(source.getCell(1, 0)).toBe("B");
    expect(source.getCell(1, 1)).toBe(2);
    expect(source.getCell(2, 0)).toBe("C");
  });

  it("inserts multiple rows at once, preserving their relative order", () => {
    const source = new ArrayDataSource([{ name: "A", price: 1 }], columns());

    source.insertRows(0, [
      ["X", 10],
      ["Y", 20]
    ]);

    expect(source.getRowCount()).toBe(3);
    expect(source.getCell(0, 0)).toBe("X");
    expect(source.getCell(1, 0)).toBe("Y");
    expect(source.getCell(2, 0)).toBe("A");
  });

  it("inserting into a large source does not rewrite the caller-owned top-level rows array", () => {
    const rows = Array.from({ length: 1_000_000 }, (_unused, index) => ({ name: `R${index}`, price: index }));
    const source = new ArrayDataSource(rows, columns());

    source.insertRows(0, [["new", 0]]);

    // Structural ownership is internal; only row objects remain shared.
    expect(source.getRowCount()).toBe(1_000_001);
    expect(rows.length).toBe(1_000_000);
  });

  it("throws a stable error code when inserting into a readonly source", () => {
    const source = new ArrayDataSource([{ name: "A" }], [{ id: "name", field: "name" }], { readonly: true });

    expect(() => source.insertRows(0, [["B"]])).toThrow(CyExcelError);
  });

  it("rejects zero-count and out-of-range inserts before splicing the caller-owned rows array", () => {
    const rows = [{ name: "A" }, { name: "B" }];
    const source = new ArrayDataSource(rows, columns());

    expect(() => source.insertRows(1, [])).toThrow(RangeError);
    expect(() => source.insertRows(3, [["C"]])).toThrow(RangeError);

    expect(rows.map((row) => row.name)).toEqual(["A", "B"]);
    expect(source.getRowCount()).toBe(2);
  });
});

describe("ArrayDataSource — deleteRows", () => {
  function columns() {
    return [{ id: "name", field: "name" }];
  }

  it("removes the given range of rows", () => {
    const source = new ArrayDataSource(
      [{ name: "A" }, { name: "B" }, { name: "C" }, { name: "D" }],
      columns()
    );

    source.deleteRows(1, 2);

    expect(source.getRowCount()).toBe(2);
    expect(source.getCell(0, 0)).toBe("A");
    expect(source.getCell(1, 0)).toBe("D");
  });

  it("throws a stable error code when deleting from a readonly source", () => {
    const source = new ArrayDataSource([{ name: "A" }], columns(), { readonly: true });

    expect(() => source.deleteRows(0, 1)).toThrow(CyExcelError);
  });

  it("rejects zero-count and out-of-range deletes before mutating rows", () => {
    const rows = [{ name: "A" }, { name: "B" }];
    const source = new ArrayDataSource(rows, columns());

    expect(() => source.deleteRows(1, 0)).toThrow(RangeError);
    expect(() => source.deleteRows(1, 2)).toThrow(RangeError);

    expect(rows.map((row) => row.name)).toEqual(["A", "B"]);
    expect(source.getRowCount()).toBe(2);
  });
});

describe("ArrayDataSource — moveRows", () => {
  function columns() {
    return [{ id: "name", field: "name" }];
  }

  function names(source: ArrayDataSource): string[] {
    return Array.from({ length: source.getRowCount() }, (_unused, row) => source.getCell(row, 0) as string);
  }

  it("moves a block of rows later in the sheet", () => {
    const source = new ArrayDataSource(
      ["A", "B", "C", "D", "E"].map((name) => ({ name })),
      columns()
    );

    source.moveRows(1, 2, 4); // move B,C to just before E

    expect(names(source)).toEqual(["A", "D", "B", "C", "E"]);
  });

  it("moves a block of rows earlier in the sheet", () => {
    const source = new ArrayDataSource(
      ["A", "B", "C", "D", "E"].map((name) => ({ name })),
      columns()
    );

    source.moveRows(3, 1, 0); // move D to the very start

    expect(names(source)).toEqual(["D", "A", "B", "C", "E"]);
  });

  it("preserves each moved row's own object identity — not a copy with the same values", () => {
    const rowB = { name: "B" };
    const rows = [{ name: "A" }, rowB, { name: "C" }];
    const source = new ArrayDataSource(rows, columns());

    source.moveRows(1, 1, 3);

    // The caller's array order is isolated, while the moved row object remains shared.
    source.setCell(2, 0, "B2");
    expect(rowB.name).toBe("B2");
    expect(rows[1]).toBe(rowB);
  });

  it("throws a stable error code when moving rows in a readonly source", () => {
    const source = new ArrayDataSource([{ name: "A" }, { name: "B" }], columns(), { readonly: true });

    expect(() => source.moveRows(0, 1, 1)).toThrow(CyExcelError);
  });

  it("rejects invalid moves before mutating row order", () => {
    const rows = ["A", "B", "C"].map((name) => ({ name }));
    const source = new ArrayDataSource(rows, columns());

    expect(() => source.moveRows(0, 0, 3)).toThrow(RangeError);
    expect(() => source.moveRows(0, 2, 1)).toThrow(RangeError);
    expect(() => source.moveRows(2, 2, 0)).toThrow(RangeError);

    expect(names(source)).toEqual(["A", "B", "C"]);
  });
});

describe("ArrayDataSource — insertColumns", () => {
  function rows() {
    return [
      { name: "A", price: 1 },
      { name: "B", price: 2 }
    ];
  }
  function baseColumns() {
    return [
      { id: "name", field: "name" },
      { id: "price", field: "price" }
    ];
  }

  it("inserts column definitions at the given position", () => {
    const source = new ArrayDataSource(rows(), baseColumns());

    source.insertColumns(1, [{ id: "qty", field: "qty" }]);

    expect(source.getColumnCount()).toBe(3);
    expect(source.getCell(0, 0)).toBe("A"); // name unaffected
    expect(source.getCell(0, 2)).toBe(1); // price shifted right
  });

  it("does not touch any row's data — a newly inserted column reads as undefined until written", () => {
    const source = new ArrayDataSource(rows(), baseColumns());

    source.insertColumns(0, [{ id: "new", field: "new" }]);

    expect(source.getCell(0, 0)).toBeUndefined();
    expect(source.getCell(1, 0)).toBeUndefined();
  });

  it("throws a stable error code when inserting into a readonly source", () => {
    const source = new ArrayDataSource(rows(), baseColumns(), { readonly: true });

    expect(() => source.insertColumns(0, [{ id: "new", field: "new" }])).toThrow(CyExcelError);
  });

  it("rejects zero-count and out-of-range inserts before changing column mappings", () => {
    const source = new ArrayDataSource(rows(), baseColumns());

    expect(() => source.insertColumns(1, [])).toThrow(RangeError);
    expect(() => source.insertColumns(3, [{ id: "qty", field: "qty" }])).toThrow(RangeError);

    expect(source.getColumnCount()).toBe(2);
    expect(source.getCell(0, 0)).toBe("A");
    expect(source.getCell(0, 1)).toBe(1);
  });
});

describe("ArrayDataSource — deleteColumns", () => {
  function rows() {
    return [{ name: "A", price: 1, qty: 5 }];
  }
  function baseColumns() {
    return [
      { id: "name", field: "name" },
      { id: "price", field: "price" },
      { id: "qty", field: "qty" }
    ];
  }

  it("removes the given range of columns", () => {
    const source = new ArrayDataSource(rows(), baseColumns());

    source.deleteColumns(1, 1); // remove "price"

    expect(source.getColumnCount()).toBe(2);
    expect(source.getCell(0, 0)).toBe("A");
    expect(source.getCell(0, 1)).toBe(5); // qty shifted left into price's old slot
  });

  it("never actually deletes the underlying row data — re-inserting the same column definition restores access to it", () => {
    const source = new ArrayDataSource(rows(), baseColumns());
    const priceColumn = baseColumns()[1]!;

    source.deleteColumns(1, 1);
    source.insertColumns(1, [priceColumn]);

    expect(source.getCell(0, 1)).toBe(1);
  });

  it("throws a stable error code when deleting from a readonly source", () => {
    const source = new ArrayDataSource(rows(), baseColumns(), { readonly: true });

    expect(() => source.deleteColumns(0, 1)).toThrow(CyExcelError);
  });

  it("rejects zero-count and out-of-range deletes before changing column mappings", () => {
    const source = new ArrayDataSource(rows(), baseColumns());

    expect(() => source.deleteColumns(1, 0)).toThrow(RangeError);
    expect(() => source.deleteColumns(2, 2)).toThrow(RangeError);

    expect(source.getColumnCount()).toBe(3);
    expect(source.getCell(0, 0)).toBe("A");
    expect(source.getCell(0, 1)).toBe(1);
    expect(source.getCell(0, 2)).toBe(5);
  });
});

describe("ArrayDataSource.reorderRows", () => {
  it("permutes rows without touching any cell value", () => {
    const source = new ArrayDataSource(
      [{ name: "c" }, { name: "a" }, { name: "b" }],
      [{ id: "name", field: "name" }]
    );

    source.reorderRows(Int32Array.from([1, 2, 0]));

    expect([source.getCell(0, 0), source.getCell(1, 0), source.getCell(2, 0)]).toEqual(["a", "b", "c"]);
  });

  it("keeps the caller's own array untouched", () => {
    const rows = [{ name: "c" }, { name: "a" }, { name: "b" }];
    const source = new ArrayDataSource(rows, [{ id: "name", field: "name" }]);

    source.reorderRows(Int32Array.from([1, 2, 0]));

    // The constructor shallow-copies the container, so structural changes
    // never reorder an array the caller still holds.
    expect(rows.map((row) => row.name)).toEqual(["c", "a", "b"]);
  });

  it("rejects an order that is not a permutation", () => {
    const source = new ArrayDataSource([{ name: "a" }, { name: "b" }], [{ id: "name", field: "name" }]);

    expect(() => source.reorderRows(Int32Array.from([0, 0]))).toThrow(/not a permutation/);
  });

  it("throws a stable error code when reordering rows in a readonly source", () => {
    const source = new ArrayDataSource([{ name: "A" }, { name: "B" }], [{ id: "name", field: "name" }], { readonly: true });

    expect(() => source.reorderRows(Int32Array.from([1, 0]))).toThrow(CyExcelError);
  });

  // Controller Ruling 4: prepareMutation's "reorder-rows" case must keep the
  // transaction's own in-progress row view current, the same way "move-rows"
  // calls moveRowPieces() before returning — otherwise a transaction that
  // reorders and then reads sees the pre-reorder order.
  it("keeps the transaction's own read view current after staging a reorder", () => {
    const source = new ArrayDataSource(
      [{ name: "c" }, { name: "a" }, { name: "b" }],
      [{ id: "name", field: "name" }]
    );

    const transaction = source.beginTransaction(0 as ModelRevision);
    transaction.prepareMutation({ type: "reorder-rows", order: Int32Array.from([1, 2, 0]) });

    expect([
      transaction.getCell(0, 0),
      transaction.getCell(1, 0),
      transaction.getCell(2, 0)
    ]).toEqual(["a", "b", "c"]);

    transaction.dispose();
  });
});

describe("ArrayDataSource — moveColumns", () => {
  function rows() {
    return [{ a: "A", b: "B", c: "C" }];
  }
  function baseColumns() {
    return [
      { id: "a", field: "a" },
      { id: "b", field: "b" },
      { id: "c", field: "c" }
    ];
  }

  it("moves a block of columns, data following via field name rather than position", () => {
    const source = new ArrayDataSource(rows(), baseColumns());

    source.moveColumns(0, 1, 3); // move "a" to the end: [b, c, a]

    expect(source.getCell(0, 0)).toBe("B");
    expect(source.getCell(0, 1)).toBe("C");
    expect(source.getCell(0, 2)).toBe("A");
  });

  it("throws a stable error code when moving columns in a readonly source", () => {
    const source = new ArrayDataSource(rows(), baseColumns(), { readonly: true });

    expect(() => source.moveColumns(0, 1, 1)).toThrow(CyExcelError);
  });

  it("rejects invalid moves before changing column mappings", () => {
    const source = new ArrayDataSource(rows(), baseColumns());

    expect(() => source.moveColumns(0, 0, 3)).toThrow(RangeError);
    expect(() => source.moveColumns(0, 2, 1)).toThrow(RangeError);
    expect(() => source.moveColumns(2, 2, 0)).toThrow(RangeError);

    expect(source.getColumnCount()).toBe(3);
    expect(source.getCell(0, 0)).toBe("A");
    expect(source.getCell(0, 1)).toBe("B");
    expect(source.getCell(0, 2)).toBe("C");
  });
});
