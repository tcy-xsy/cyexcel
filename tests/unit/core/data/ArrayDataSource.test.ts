import { describe, expect, it } from "vitest";

import { ArrayDataSource, CyExcelError, CyExcelErrorCode } from "../../../../src/core";

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

  it("inserting into a large source shifts the existing array in place instead of copying it", () => {
    const rows = Array.from({ length: 1_000_000 }, (_unused, index) => ({ name: `R${index}`, price: index }));
    const source = new ArrayDataSource(rows, columns());

    source.insertRows(0, [["new", 0]]);

    // splice() mutates `rows` in place — it's still the exact same array instance.
    expect(source.getRowCount()).toBe(1_000_001);
    expect(rows.length).toBe(1_000_001);
  });

  it("throws a stable error code when inserting into a readonly source", () => {
    const source = new ArrayDataSource([{ name: "A" }], [{ id: "name", field: "name" }], { readonly: true });

    expect(() => source.insertRows(0, [["B"]])).toThrow(CyExcelError);
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

    // The exact same object reference ends up at the new position.
    expect(rows[2]).toBe(rowB);
  });

  it("throws a stable error code when moving rows in a readonly source", () => {
    const source = new ArrayDataSource([{ name: "A" }, { name: "B" }], columns(), { readonly: true });

    expect(() => source.moveRows(0, 1, 1)).toThrow(CyExcelError);
  });
});
