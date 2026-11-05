import { describe, expect, it } from "vitest";
import { createFormulaEngine } from "../../../src/formula";
import { SheetModel } from "../../../src/core/sheet/SheetModel";
import { TableStore } from "../../../src/core/data/TableStore";

function setup(options: Parameters<typeof createFormulaEngine>[0] = {}) {
  const model = SheetModel.createEmpty({ blankRows: 100, blankColumns: 10 });
  model.enableFormulas(createFormulaEngine(options));
  return { model, store: new TableStore(model) };
}

describe("formula calculation", () => {
  it("stores formulas separately, calculates precedence and follows transitive dependencies", () => {
    const { model, store } = setup();
    store.setCell(0, 0, 2); store.setCell(1, 0, 3);
    store.setCell(0, 1, "=SUM(A1:A2)"); store.setCell(0, 2, "=B1*2+1");
    expect(store.getCell(0, 2)).toBe(11);
    expect(model.getRecord(0, 1)?.f).toBe("=SUM(A1:A2)");
    store.setCell(0, 0, 7);
    expect(model.takeFormulaInvalidations()).toEqual(expect.arrayContaining([{ row: 0, column: 1 }, { row: 0, column: 2 }]));
    expect(store.getCell(0, 2)).toBe(21);
    store.setCell(0, 1, 9);
    expect(store.getCell(0, 2)).toBe(19);
    expect(model.getRecord(0, 1)?.f).toBeUndefined();
  });

  it.each([
    ["=AVERAGE(A1:A3)", 2], ["=MIN(A1:A3)", 1], ["=MAX(A1:A3)", 3],
    ["=COUNT(A1:A3)", 3], ["=IF(A1>0,TRUE,FALSE)", true], ["=ROUND(1.234,2)", 1.23],
    ["=SUM($A$1:A3)", 6], ["=A1+A2*3", 7], ["=CONCATENATE(\"a\",\"b\")", "ab"],
    ["=COUNTIF(A1:A3,\">1\")", 2], ["=SUMIF(A1:A3,\">1\")", 5], ["=IFERROR(1/0,99)", 99], ["=ABS(-3)", 3]
  ])("evaluates %s", (formula, result) => {
    const { store } = setup();
    store.setRange({ startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 }, [[1], [2], [3]]);
    store.setCell(0, 1, formula);
    expect(store.getCell(0, 1)).toBe(result);
  });

  it("handles blank values, formula errors and recovery", () => {
    const { model, store } = setup();
    store.setCell(0, 4, "=A1"); expect(store.getCell(0, 4)).toBe(0);
    store.setCell(0, 1, "=A1+2"); expect(store.getCell(0, 1)).toBe(2);
    store.setCell(0, 2, "=1/A1"); expect(store.getCell(0, 2)).toBe("#DIV/0!");
    store.setCell(0, 3, "=C1+1"); expect(store.getCell(0, 3)).toBe("#DIV/0!");
    expect(model.getFormulaError(0, 2)).toBe("#DIV/0!");
    store.setCell(0, 0, 2);
    expect(store.getCell(0, 2)).toBe(0.5); expect(store.getCell(0, 3)).toBe(1.5);
    expect(model.getFormulaError(0, 2)).toBeUndefined();
    store.setCell(0, 1, "=UNKNOWN(1)"); expect(store.getCell(0, 1)).toBe("#NAME?");
    store.setCell(0, 1, "=SUM("); expect(typeof store.getCell(0, 1)).toBe("string");
    expect(model.getFormulaError(0, 1)).toBeDefined();
    store.setCell(0, 1, "=SUM(A1)"); expect(store.getCell(0, 1)).toBe(2);
  });

  it("detects cycles and permits correcting a cyclic dependency", () => {
    const { store } = setup();
    store.setCell(0, 0, "=B1+1"); store.setCell(0, 1, "=A1+1");
    expect(store.getCell(0, 0)).toBe("#CYCLE!");
    store.setCell(0, 1, 5); expect(store.getCell(0, 0)).toBe(6);
  });

  it("updates dependency edges when a formula is replaced", () => {
    const { model, store } = setup();
    store.setCell(0, 0, 2); store.setCell(1, 0, 5); store.setCell(0, 1, "=A1");
    expect(store.getCell(0, 1)).toBe(2);
    store.setCell(0, 1, "=A2"); expect(store.getCell(0, 1)).toBe(5);
    model.takeFormulaInvalidations();
    store.setCell(0, 0, 9); expect(model.takeFormulaInvalidations()).toEqual([]);
    store.setCell(1, 0, 8); expect(store.getCell(0, 1)).toBe(8);
  });

  it("isolates preview clones and saves fresh cached results without mutating raw records", () => {
    const { model, store } = setup();
    store.setCell(0, 0, 3); store.setCell(0, 1, "=A1*2"); expect(store.getCell(0, 1)).toBe(6);
    const copy = model.clone(); copy.setCell(0, 0, 9);
    expect(copy.getCell(0, 1)).toBe(18); expect(model.getCell(0, 1)).toBe(6);
    const snapshot = model.toSnapshot();
    expect(snapshot.cellData[0]?.[1]).toMatchObject({ f: "=A1*2", v: 6 });
    expect(model.getRecord(0, 1)?.v).toBeUndefined();
    const restored = SheetModel.fromSnapshot(snapshot); restored.enableFormulas(createFormulaEngine());
    restored.setCell(0, 0, 7); expect(restored.getCell(0, 1)).toBe(14);
  });

  it("keeps opt-out strings literal and supports escaped formula-like text", () => {
    const model = SheetModel.createEmpty({ blankRows: 10, blankColumns: 2 });
    const store = new TableStore(model); store.setCell(0, 0, "=1+1");
    expect(store.getCell(0, 0)).toBe("=1+1");
    model.enableFormulas(createFormulaEngine());
    store.setCell(0, 1, "'=1+1"); expect(store.getCell(0, 1)).toBe("=1+1");
    expect(model.getRecord(0, 1)?.f).toBeUndefined();
  });

  it("bounds range work and explicitly rejects unsupported dynamic or cross-sheet references", () => {
    const { store } = setup({ maxRangeCells: 10 });
    store.setCell(0, 1, "=SUM(A1:A11)"); expect(store.getCell(0, 1)).toBe("#NUM!");
    store.setCell(0, 1, "=Other!A1"); expect(store.getCell(0, 1)).toBe("#REF!");
    store.setCell(0, 1, '=INDIRECT("A1")'); expect(store.getCell(0, 1)).toBe("#NAME?");
    store.setCell(0, 1, '=WEBSERVICE("https://example.com")'); expect(store.getCell(0, 1)).toBe("#NAME?");
  });

  it("blocks unsupported structural edits atomically while formulas exist", () => {
    const { model, store } = setup(); store.setCell(0, 0, "=1+1");
    const before = model.toSnapshot();
    expect(() => store.insertRows(0, [[]])).toThrow("while formulas exist");
    expect(() => model.deleteColumns(0, 1)).toThrow("while formulas exist");
    expect(model.toSnapshot()).toEqual(before);
    store.setCell(0, 0, 2); expect(() => model.insertRows(0, 1)).not.toThrow();
  });

  it("preserves styles and images when replacing formulas with values or clearing", () => {
    const { model, store } = setup();
    model.setRecord(0, 0, { v: 3, s: "bold", img: "test.png" });
    store.setCell(0, 0, "=2+3");
    expect(model.getRecord(0, 0)).toEqual({ v: undefined, f: "=2+3", s: "bold", img: "test.png" });
    store.setCell(0, 0, 7);
    expect(model.getRecord(0, 0)).toMatchObject({ v: 7, s: "bold", img: "test.png" });
    expect(model.getRecord(0, 0)?.f).toBeUndefined();
    store.setCell(0, 0, undefined);
    expect(model.getRecord(0, 0)).toMatchObject({ v: undefined, s: "bold", img: "test.png" });
  });

  it("bounds dependency depth and recovers after shortening the chain", () => {
    const { store } = setup({ maxDepth: 3 });
    for (let row = 0; row < 5; row++) store.setCell(row, 0, `=A${row + 2}+1`);
    expect(store.getCell(0, 0)).toBe("#NUM!");
    store.setCell(2, 0, 4);
    expect(store.getCell(0, 0)).toBe(6);
  });
});
