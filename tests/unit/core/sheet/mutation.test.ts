import { describe, expect, it } from "vitest";
import { isStructuralSheetMutation, type SheetMutation } from "../../../../src/core/sheet/mutation";
import { CellValueType } from "../../../../src/core/sheet/types";

describe("SheetMutation — 格子", () => {
  it("set-cell 带的是记录，不是裸值", () => {
    const m = {
      type: "set-cell",
      row: 0,
      column: 0,
      record: { v: 1, t: CellValueType.NUMBER }
    } satisfies SheetMutation;
    expect(m.record.t).toBe(CellValueType.NUMBER);
  });

  it("set-cell 的 record 可以是 undefined，表示删格", () => {
    const m = { type: "set-cell", row: 0, column: 0, record: undefined } satisfies SheetMutation;
    expect(m.record).toBeUndefined();
  });

  it("格子变更不是结构变更", () => {
    expect(isStructuralSheetMutation({ type: "set-cell", row: 0, column: 0, record: undefined })).toBe(false);
    expect(
      isStructuralSheetMutation({ type: "set-range", range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }, records: [[undefined]] })
    ).toBe(false);
  });
});

describe("SheetMutation — 行", () => {
  it("insert-rows 带每行的记录，才能承载删行的逆", () => {
    const m: SheetMutation = { type: "insert-rows", start: 1, rows: [{ 0: { v: "a" } }] };
    expect(m.rows[0]?.[0]?.v).toBe("a");
  });

  it("delete-rows 只要 start/count——记录由逆变更负责", () => {
    const m: SheetMutation = { type: "delete-rows", start: 1, count: 2 };
    expect(m.count).toBe(2);
  });

  it("四个行操作都是结构变更", () => {
    const rows: SheetMutation[] = [
      { type: "insert-rows", start: 0, rows: [] },
      { type: "delete-rows", start: 0, count: 1 },
      { type: "move-rows", start: 0, count: 1, target: 2 },
      { type: "reorder-rows", order: Int32Array.from([0]) },
      { type: "restore-rows", start: 0, rows: [], meta: [], tables: {}, mergesBefore: [] }
    ];
    for (const m of rows) {
      expect(isStructuralSheetMutation(m)).toBe(true);
    }
  });
});

describe("SheetMutation — 列", () => {
  it("insert-columns 只要个数，不带 id/field（ADR §3.8.2 列不再有身份）", () => {
    const m: SheetMutation = { type: "insert-columns", start: 0, count: 2 };
    expect(m.count).toBe(2);
    // 名字承诺的是"没有身份"，就得真的测到——多余属性检查会让这行编译失败
    // @ts-expect-error 列不再有 id/field
    const withIdentity: SheetMutation = { type: "insert-columns", start: 0, count: 1, id: "c1" };
    expect(withIdentity.type).toBe("insert-columns");
  });

  it("delete-columns 的逆要带回被删列的格子与列元数据", () => {
    const m: SheetMutation = {
      type: "restore-columns",
      start: 1,
      count: 1,
      cells: [{ 0: { v: "x" } }],
      meta: [{ w: 140 }],
      tables: {},
      mergesBefore: []
    };
    expect(m.meta[0]?.w).toBe(140);
  });

  it("三个列操作都是结构变更", () => {
    const columns: SheetMutation[] = [
      { type: "insert-columns", start: 0, count: 1 },
      { type: "delete-columns", start: 0, count: 1 },
      { type: "move-columns", start: 0, count: 1, target: 2 },
      { type: "restore-columns", start: 0, count: 1, cells: [], meta: [], tables: {}, mergesBefore: [] }
    ];
    for (const m of columns) {
      expect(isStructuralSheetMutation(m)).toBe(true);
    }
  });
});
