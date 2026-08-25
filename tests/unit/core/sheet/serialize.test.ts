import { describe, expect, it } from "vitest";

import {
  CellMatrix,
  SNAPSHOT_VERSION,
  assertValidWorkbookSnapshot,
  deserializeSheet,
  serializeSheet
} from "../../../../src/core/sheet";
import type { WorkbookSnapshot } from "../../../../src/core/sheet";
import { CyExcelError, CyExcelErrorCode } from "../../../../src/core/errors";

function buildInput(matrix: CellMatrix) {
  return {
    id: "sheet-1",
    name: "Sheet1",
    blankRows: 1000,
    blankColumns: 26,
    defaultRowHeight: 28,
    defaultColumnWidth: 120,
    matrix,
    rowData: {},
    columnData: {},
    mergeData: []
  };
}

describe("serializeSheet", () => {
  it("writes occupied cells as nested sparse objects", () => {
    const matrix = new CellMatrix();
    matrix.set(2, 1, { v: 42 });
    const snapshot = serializeSheet(buildInput(matrix));
    expect(snapshot.cellData).toEqual({ 2: { 1: { v: 42 } } });
  });

  it("omits unoccupied rows entirely rather than writing empty buckets", () => {
    const matrix = new CellMatrix();
    matrix.set(5, 0, { v: "a" });
    const snapshot = serializeSheet(buildInput(matrix));
    expect(Object.keys(snapshot.cellData)).toEqual(["5"]);
  });

  it("does not write rowCount or columnCount", () => {
    const matrix = new CellMatrix();
    matrix.set(0, 0, { v: "a" });
    const snapshot = serializeSheet(buildInput(matrix)) as unknown as Record<string, unknown>;
    // spec §11 不变量 5：坐标空间是派生量，不得出现在快照里
    expect(snapshot["rowCount"]).toBeUndefined();
    expect(snapshot["columnCount"]).toBeUndefined();
  });

  it("survives a JSON round trip unchanged", () => {
    const matrix = new CellMatrix();
    matrix.set(0, 0, { v: "text" });
    matrix.set(1, 1, { v: 3.5, t: "money", s: "s7" });
    const snapshot = serializeSheet(buildInput(matrix));
    expect(JSON.parse(JSON.stringify(snapshot))).toEqual(snapshot);
  });
});

describe("deserializeSheet", () => {
  it("rebuilds the matrix and its maxOccupied", () => {
    const matrix = new CellMatrix();
    matrix.set(4, 2, { v: "a" });
    const restored = deserializeSheet(serializeSheet(buildInput(matrix)));
    expect(restored.matrix.get(4, 2)).toEqual({ v: "a" });
    expect(restored.matrix.maxOccupiedRow()).toBe(4);
    expect(restored.matrix.maxOccupiedColumn()).toBe(2);
    expect(restored.matrix.occupiedCount()).toBe(1);
  });

  it("round-trips an empty sheet", () => {
    const restored = deserializeSheet(serializeSheet(buildInput(new CellMatrix())));
    expect(restored.matrix.occupiedCount()).toBe(0);
    expect(restored.matrix.maxOccupiedRow()).toBeUndefined();
    expect(restored.blankRows).toBe(1000);
  });

  it("round-trips a date cell through its semantic type name", () => {
    // spec §4.1：不需要额外的原始类型标记，t 已经携带了这个信息
    const matrix = new CellMatrix();
    matrix.set(0, 0, { v: "2026-01-01T00:00:00.000Z", t: "date" });
    const restored = deserializeSheet(
      JSON.parse(JSON.stringify(serializeSheet(buildInput(matrix)))) as ReturnType<typeof serializeSheet>
    );
    expect(restored.matrix.get(0, 0)).toEqual({ v: "2026-01-01T00:00:00.000Z", t: "date" });
  });
});

describe("assertValidWorkbookSnapshot", () => {
  function buildWorkbook(overrides: Partial<WorkbookSnapshot> = {}): WorkbookSnapshot {
    return {
      version: SNAPSHOT_VERSION,
      sheetOrder: ["sheet-1"],
      sheets: { "sheet-1": serializeSheet(buildInput(new CellMatrix())) },
      styles: {},
      resources: [],
      ...overrides
    };
  }

  it("accepts a well-formed workbook", () => {
    expect(() => assertValidWorkbookSnapshot(buildWorkbook())).not.toThrow();
  });

  it("rejects an unknown schema version", () => {
    expect(() => assertValidWorkbookSnapshot(buildWorkbook({ version: 999 }))).toThrow(CyExcelError);
  });

  it("rejects more than one sheet until multi-sheet opens in M3", () => {
    // spec §11 不变量 6
    const sheet = serializeSheet(buildInput(new CellMatrix()));
    expect(() =>
      assertValidWorkbookSnapshot(
        buildWorkbook({
          sheetOrder: ["sheet-1", "sheet-2"],
          sheets: { "sheet-1": sheet, "sheet-2": { ...sheet, id: "sheet-2" } }
        })
      )
    ).toThrow(CyExcelError);
  });

  it("rejects a sheetOrder entry with no matching sheet", () => {
    expect(() => assertValidWorkbookSnapshot(buildWorkbook({ sheetOrder: ["missing"] }))).toThrow(
      CyExcelError
    );
  });

  it("rejects a negative blankRows", () => {
    const sheet = serializeSheet(buildInput(new CellMatrix()));
    expect(() =>
      assertValidWorkbookSnapshot(
        buildWorkbook({ sheets: { "sheet-1": { ...sheet, blankRows: -1 } } })
      )
    ).toThrow(CyExcelError);
  });

  it("reports InvalidOptions for every rejection", () => {
    try {
      assertValidWorkbookSnapshot(buildWorkbook({ version: 999 }));
    } catch (error) {
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.InvalidOptions);
    }
  });
});
