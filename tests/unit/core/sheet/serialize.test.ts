import { describe, expect, it } from "vitest";

import {
  CellMatrix,
  SNAPSHOT_VERSION,
  assertValidWorkbookSnapshot,
  deserializeSheet,
  serializeSheet
} from "../../../../src/core/sheet";
import type { SheetSnapshot, WorkbookSnapshot } from "../../../../src/core/sheet";
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
    // toStrictEqual：稀疏性（spec §11 不变量 1）说的是"未占用的格子不出现"，
    // 而 toEqual 把 `{ 3: undefined }` 当作与"没有键 3"相等，正好看不见它。
    expect(snapshot.cellData).toStrictEqual({ 2: { 1: { v: 42 } } });
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
    // spec §11 不变量 5：坐标空间是派生量，不得出现在快照里。
    // 断言"键不存在"而不是 toBeUndefined()——后者对 `rowCount: undefined`
    // 也放行，那正是这条用例要抓的东西。
    expect("rowCount" in snapshot).toBe(false);
    expect("columnCount" in snapshot).toBe(false);
  });

  it("survives a JSON round trip unchanged", () => {
    const matrix = new CellMatrix();
    matrix.set(0, 0, { v: "text" });
    matrix.set(1, 1, { v: 3.5, t: "money", s: "s7" });
    const snapshot = serializeSheet(buildInput(matrix));
    // JSON.stringify 恰好丢掉值为 undefined 的键，而 toEqual 也正好把那种键
    // 当作不存在——两者叠加，这条断言对**往返唯一会改变的那件事**是瞎的。
    expect(JSON.parse(JSON.stringify(snapshot))).toStrictEqual(snapshot);
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

  // cellData 的键在真实输入里来自 JSON.parse，因此在类型层面就是 `any`。这里用
  // 一次显式的双重断言构造带有非法键的测试夹具，模拟那条未经类型系统把关的路径，
  // 而不是放宽 SheetSnapshot 的类型来迁就它。
  function withMalformedCellData(sheet: SheetSnapshot, cellData: unknown): SheetSnapshot {
    return { ...sheet, cellData } as unknown as SheetSnapshot;
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

  it("rejects a sheets key that disagrees with the sheet's own id", () => {
    // 不查这一条的后果是静默改键：sheetOrder 与 sheets 的键用外面那个 id，而
    // WorkbookModel.toSnapshot() 输出的是 sheet 自己的 id，往返之后
    // sheetOrder 从 ["alias"] 变成 ["renamed"]。
    const sheet = serializeSheet(buildInput(new CellMatrix()));
    expect(() =>
      assertValidWorkbookSnapshot(
        buildWorkbook({
          sheetOrder: ["alias"],
          sheets: { alias: { ...sheet, id: "renamed" } }
        })
      )
    ).toThrow(CyExcelError);
  });

  it("rejects a negative blankRows", () => {
    const sheet = serializeSheet(buildInput(new CellMatrix()));
    expect(() =>
      assertValidWorkbookSnapshot(
        buildWorkbook({ sheets: { "sheet-1": { ...sheet, blankRows: -1 } } })
      )
    ).toThrow(CyExcelError);
  });

  it("rejects a non-numeric row key in cellData", () => {
    const sheet = serializeSheet(buildInput(new CellMatrix()));
    expect(() =>
      assertValidWorkbookSnapshot(
        buildWorkbook({
          sheets: { "sheet-1": withMalformedCellData(sheet, { abc: { 0: { v: 1 } } }) }
        })
      )
    ).toThrow(CyExcelError);
  });

  it("rejects a negative row key in cellData", () => {
    const sheet = serializeSheet(buildInput(new CellMatrix()));
    expect(() =>
      assertValidWorkbookSnapshot(
        buildWorkbook({
          sheets: { "sheet-1": withMalformedCellData(sheet, { "-3": { 0: { v: 1 } } }) }
        })
      )
    ).toThrow(CyExcelError);
  });

  it("rejects a non-numeric column key in cellData", () => {
    const sheet = serializeSheet(buildInput(new CellMatrix()));
    expect(() =>
      assertValidWorkbookSnapshot(
        buildWorkbook({
          sheets: { "sheet-1": withMalformedCellData(sheet, { 0: { abc: { v: 1 } } }) }
        })
      )
    ).toThrow(CyExcelError);
  });

  it("rejects a fractional row key in cellData", () => {
    const sheet = serializeSheet(buildInput(new CellMatrix()));
    expect(() =>
      assertValidWorkbookSnapshot(
        buildWorkbook({
          sheets: { "sheet-1": withMalformedCellData(sheet, { "1.5": { 0: { v: 1 } } }) }
        })
      )
    ).toThrow(CyExcelError);
  });

  it("reports InvalidOptions for every rejection", () => {
    expect.hasAssertions();
    const sheet = serializeSheet(buildInput(new CellMatrix()));
    const rejections: Array<[string, WorkbookSnapshot]> = [
      ["unknown version", buildWorkbook({ version: 999 })],
      [
        "more than one sheet",
        buildWorkbook({
          sheetOrder: ["sheet-1", "sheet-2"],
          sheets: { "sheet-1": sheet, "sheet-2": { ...sheet, id: "sheet-2" } }
        })
      ],
      ["sheetOrder entry with no matching sheet", buildWorkbook({ sheetOrder: ["missing"] })],
      ["negative blankRows", buildWorkbook({ sheets: { "sheet-1": { ...sheet, blankRows: -1 } } })],
      [
        "non-numeric row key in cellData",
        buildWorkbook({
          sheets: { "sheet-1": withMalformedCellData(sheet, { abc: { 0: { v: 1 } } }) }
        })
      ],
      [
        "negative row key in cellData",
        buildWorkbook({
          sheets: { "sheet-1": withMalformedCellData(sheet, { "-3": { 0: { v: 1 } } }) }
        })
      ],
      [
        "non-numeric column key in cellData",
        buildWorkbook({
          sheets: { "sheet-1": withMalformedCellData(sheet, { 0: { abc: { v: 1 } } }) }
        })
      ],
      [
        "fractional row key in cellData",
        buildWorkbook({
          sheets: { "sheet-1": withMalformedCellData(sheet, { "1.5": { 0: { v: 1 } } }) }
        })
      ]
    ];
    for (const [label, workbook] of rejections) {
      let threw = false;
      try {
        assertValidWorkbookSnapshot(workbook);
      } catch (error) {
        threw = true;
        expect((error as CyExcelError).code, label).toBe(CyExcelErrorCode.InvalidOptions);
      }
      expect(threw, label).toBe(true);
    }
  });
});
