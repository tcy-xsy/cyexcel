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
