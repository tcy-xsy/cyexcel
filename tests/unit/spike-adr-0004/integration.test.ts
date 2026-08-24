import { describe, expect, it } from "vitest";

import { DeleteRowsCommand, ReorderRowsCommand, SetCellCommand } from "../../../src/core/commands";
import { TableStore } from "../../../src/core/data/TableStore";
import { TableCommandManager } from "../../../src/core/integration";
import type { ModelRevision } from "../../../src/core/model";
import { SpikeSheetDataSource } from "../../spike/adr-0004/SpikeSheetDataSource";

/** ModelRevision is a branded number; the literal 0 is not assignable. */
const REV = (n: number): ModelRevision => n as unknown as ModelRevision;

function createStore(rowCount = 10, columnCount = 5): { store: TableStore; source: SpikeSheetDataSource } {
  const source = new SpikeSheetDataSource(rowCount, columnCount);
  const store = new TableStore(source);
  return { store, source };
}

function createManaged(
  rowCount = 10,
  columnCount = 5
): { store: TableStore; source: SpikeSheetDataSource; commands: TableCommandManager } {
  const { store, source } = createStore(rowCount, columnCount);
  return { store, source, commands: new TableCommandManager(store) };
}

describe("SpikeSheetDataSource against the real TableStore", () => {
  it("reads a bare value through getCell, not a record", () => {
    const { store, source } = createStore();
    source.setRecord(2, 3, { v: 42, t: 2 });
    // ADR §3.6 / invariant 1: getCell always returns the bare value.
    expect(store.getCell(2, 3)).toBe(42);
  });

  it("commits a set-cell mutation through the real store", () => {
    const { store } = createStore();
    store.setCell(1, 1, "hello");
    expect(store.getCell(1, 1)).toBe("hello");
  });

  it("produces a record-native inverse, not a PropertyDescriptor one", () => {
    const { source } = createStore();
    source.setRecord(1, 1, { v: "before", t: 1 });
    const transaction = source.beginTransaction(REV(0));
    const prepared = transaction.prepareMutation({ type: "set-cell", row: 1, column: 1, value: "after" });
    expect(prepared.inverse.type).toBe("restore-records");
  });

  it("produces an inverse that restores the previous value", () => {
    const { source } = createStore();
    source.setRecord(1, 1, { v: "before", t: 1 });
    const transaction = source.beginTransaction(REV(0));
    const prepared = transaction.prepareMutation({ type: "set-cell", row: 1, column: 1, value: "after" });
    prepared.commit();
    expect(source.getCell(1, 1)).toBe("after");

    const undo = source.beginTransaction(REV(1));
    undo.prepareMutation(prepared.inverse).commit();
    expect(source.getCell(1, 1)).toBe("before");
  });

  it("restores a cell that was previously unoccupied by deleting it", () => {
    const { source } = createStore();
    const transaction = source.beginTransaction(REV(0));
    const prepared = transaction.prepareMutation({ type: "set-cell", row: 4, column: 4, value: "new" });
    prepared.commit();
    const undo = source.beginTransaction(REV(1));
    undo.prepareMutation(prepared.inverse).commit();
    // `record: undefined` means "delete it", not "write undefined into it".
    expect(source.getRecord(4, 4)).toBeUndefined();
    expect(source.occupiedCount()).toBe(0);
  });

  it("preserves non-value attributes across an undo", () => {
    const { source } = createStore();
    source.setRecord(1, 1, { v: "before", t: 1, s: "style-7" });
    const transaction = source.beginTransaction(REV(0));
    const prepared = transaction.prepareMutation({ type: "set-cell", row: 1, column: 1, value: "after" });
    prepared.commit();
    const undo = source.beginTransaction(REV(1));
    undo.prepareMutation(prepared.inverse).commit();
    expect(source.getRecord(1, 1)).toEqual({ v: "before", t: 1, s: "style-7" });
  });

  it("rolls back without touching the matrix", () => {
    const { source } = createStore();
    source.setRecord(1, 1, { v: "before" });
    const transaction = source.beginTransaction(REV(0));
    const prepared = transaction.prepareMutation({ type: "set-cell", row: 1, column: 1, value: "after" });
    prepared.commit();
    prepared.rollback();
    expect(source.getCell(1, 1)).toBe("before");
  });

  it("reports an estimatedBytes that scales with the record, not a hardcoded 64", () => {
    const { source } = createStore();
    const transaction = source.beginTransaction(REV(0));
    const prepared = transaction.prepareMutation({ type: "set-cell", row: 1, column: 1, value: "x" });
    expect(prepared.estimatedBytes).toBeGreaterThan(0);
    expect(Number.isFinite(prepared.estimatedBytes)).toBe(true);
  });

  it("charges nothing for the record when the overwritten cell was unoccupied", () => {
    const { source } = createStore();
    source.setRecord(0, 0, { v: "occupied" });
    const transaction = source.beginTransaction(REV(0));
    // 64 base + 28 entry + 36 record (Task 2's measured 35.78 B/cell, rounded up).
    expect(transaction.prepareMutation({ type: "set-cell", row: 0, column: 0, value: "x" }).estimatedBytes).toBe(128);
    // 64 base + 28 entry; the inverse retains no record, it retains "delete me".
    expect(transaction.prepareMutation({ type: "set-cell", row: 2, column: 2, value: "x" }).estimatedBytes).toBe(92);
  });

  it("reorders rows through the real store and keeps attributes attached", () => {
    const { store, source } = createStore(3, 2);
    source.setRecord(0, 0, { v: "a", s: "style-a" });
    source.setRecord(2, 0, { v: "c", s: "style-c" });
    store.reorderRows(Int32Array.from([2, 1, 0]));
    expect(store.getCell(0, 0)).toBe("c");
    expect(source.getRecord(0, 0)).toEqual({ v: "c", s: "style-c" });
  });

  it("inserts rows through the real store", () => {
    const { store } = createStore(3, 2);
    store.setCell(0, 0, "keep");
    store.insertRows(0, [["new-a", "new-b"]]);
    expect(store.getRowCount()).toBe(4);
    expect(store.getCell(0, 0)).toBe("new-a");
    expect(store.getCell(1, 0)).toBe("keep");
  });

  it("inserts a blank row with no payload at all", () => {
    const { store, source } = createStore(3, 2);
    source.setRecord(0, 0, { v: "keep" });
    // Q2: a sparse row that nobody has written to occupies no cells, so the
    // payload here carries no information beyond its own length.
    store.insertRows(1, [[undefined, undefined]]);
    expect(store.getRowCount()).toBe(4);
    expect(source.occupiedCount()).toBe(1);
    expect(source.getRecord(0, 0)).toEqual({ v: "keep" });
  });

  it("deletes rows through the real store", () => {
    const { store } = createStore(3, 2);
    store.setCell(2, 0, "last");
    store.deleteRows(0, 1);
    expect(store.getRowCount()).toBe(2);
    expect(store.getCell(1, 0)).toBe("last");
  });
});

describe("the record matrix inside a transaction", () => {
  it("shows staged writes through the transaction read view before commit", () => {
    const { source } = createStore(3, 2);
    source.setRecord(0, 0, { v: "base" });
    const transaction = source.beginTransaction(REV(0));
    transaction.prepareMutation({ type: "set-cell", row: 0, column: 0, value: "staged" });
    expect(transaction.getCell(0, 0)).toBe("staged");
    // The live matrix must not have moved yet.
    expect(source.getCell(0, 0)).toBe("base");
  });

  it("shifts the staged read view when a structural mutation is staged", () => {
    const { source } = createStore(3, 2);
    source.setRecord(0, 0, { v: "base" });
    const transaction = source.beginTransaction(REV(0));
    transaction.prepareMutation({ type: "insert-rows", start: 0, rows: [["fresh", undefined]] });
    expect(transaction.getRowCount()).toBe(4);
    expect(transaction.getCell(0, 0)).toBe("fresh");
    expect(transaction.getCell(1, 0)).toBe("base");
  });

  it("moves rows through the real store and carries the record", () => {
    const { store, source } = createStore(3, 2);
    source.setRecord(0, 0, { v: "a", s: "style-a" });
    store.moveRows(0, 1, 3);
    expect(source.getRecord(2, 0)).toEqual({ v: "a", s: "style-a" });
  });
});

describe("undo and redo through the real CommandManager and HistoryManager", () => {
  it("undoes a set-cell and puts the whole record back", () => {
    const { source, commands } = createManaged(3, 2);
    source.setRecord(0, 0, { v: "before", t: 1, s: "style-7" });
    commands.execute(new SetCellCommand(0, 0, "after"));
    expect(source.getCell(0, 0)).toBe("after");
    // The style survived the forward write too - set-cell must merge, not replace.
    expect(source.getRecord(0, 0)).toEqual({ v: "after", t: 1, s: "style-7" });

    commands.undo();
    expect(source.getRecord(0, 0)).toEqual({ v: "before", t: 1, s: "style-7" });
  });

  it("redoes a set-cell after undoing it", () => {
    const { source, commands } = createManaged(3, 2);
    source.setRecord(0, 0, { v: "before", s: "style-7" });
    commands.execute(new SetCellCommand(0, 0, "after"));
    commands.undo();
    commands.redo();
    expect(source.getRecord(0, 0)).toEqual({ v: "after", s: "style-7" });
  });

  it("undoes a write into a previously unoccupied cell by deleting the record", () => {
    const { source, commands } = createManaged(3, 2);
    commands.execute(new SetCellCommand(1, 1, "new"));
    expect(source.occupiedCount()).toBe(1);
    commands.undo();
    expect(source.getRecord(1, 1)).toBeUndefined();
    expect(source.occupiedCount()).toBe(0);
  });

  it("emits a cell-value Patch for the record-native inverse", () => {
    const { source, commands } = createManaged(3, 2);
    source.setRecord(0, 1, { v: "before" });
    commands.execute(new SetCellCommand(0, 1, "after"));
    const patches = commands.undo();
    // restore-records has no Patch shape of its own; TableStore.ts:59's
    // `return fallback` hands back the originating set-cell's Patch.
    expect(patches).toEqual([{ type: "cell-value", row: 0, column: 1 }]);
  });

  it("undoes a reorder and the records travel back with their styles", () => {
    const { source, commands } = createManaged(3, 2);
    source.setRecord(0, 0, { v: "a", s: "style-a" });
    source.setRecord(2, 0, { v: "c", s: "style-c" });
    commands.execute(new ReorderRowsCommand(Int32Array.from([2, 1, 0])));
    expect(source.getRecord(0, 0)).toEqual({ v: "c", s: "style-c" });
    commands.undo();
    expect(source.getRecord(0, 0)).toEqual({ v: "a", s: "style-a" });
    expect(source.getRecord(2, 0)).toEqual({ v: "c", s: "style-c" });
  });

  it("undoes a delete-rows and gets the row's values back", () => {
    const { source, commands } = createManaged(3, 2);
    source.setRecord(1, 0, { v: "gone" });
    commands.execute(new DeleteRowsCommand(1, 1));
    expect(source.getRowCount()).toBe(2);
    commands.undo();
    expect(source.getRowCount()).toBe(3);
    expect(source.getCell(1, 0)).toBe("gone");
  });

  it.fails(
    // FINDING (Q2), and it is a DESIGN STANCE, not a type-level blocker.
    // delete-rows' inverse must be a *single* DataMutation, and the only
    // row-inserting variants are shaped for values (`insert-rows`,
    // src/core/types.ts:86) or caller row objects (`insert-row-objects`,
    // src/core/types.ts:87). A CellRecord would in fact fit through
    // `insert-rows` - its payload is `unknown[][]` - but that gives one
    // variant two meanings ("these are values" forward, "these are records"
    // inverse) with nothing in the type to say which. This source refuses
    // that, so `t` and `s` are dropped on the way through history.
    // Rollback inside the transaction keeps them (closures hold the records);
    // only the replayed-through-history path loses them.
    // To fix, src/core/types.ts:86-87 must gain a record-carrying variant
    // (e.g. `insert-records { start, rows: readonly (readonly (CellRecord |
    // undefined)[])[] }`), and src/core/data/TableStore.ts:68-70 must map it
    // to a row-structure "insert" Patch.
    "undoes a delete-rows and keeps the deleted cells' non-value attributes",
    () => {
      const { source, commands } = createManaged(3, 2);
      source.setRecord(1, 0, { v: "gone", t: 1, s: "style-9" });
      commands.execute(new DeleteRowsCommand(1, 1));
      commands.undo();
      expect(source.getRecord(1, 0)).toEqual({ v: "gone", t: 1, s: "style-9" });
    }
  );

  it("pins exactly what a delete-rows undo loses", () => {
    const { source, commands } = createManaged(3, 2);
    source.setRecord(1, 0, { v: "gone", t: 1, s: "style-9" });
    commands.execute(new DeleteRowsCommand(1, 1));
    commands.undo();
    // The value survives; `t` and `s` do not. This is the positive statement
    // of the it.fails() above: the inverse round-trips through a bare value
    // array, so everything that is not `v` is dropped.
    expect(source.getRecord(1, 0)).toEqual({ v: "gone" });
  });

  it("keeps the deleted records across an in-transaction rollback", () => {
    const { source } = createStore(3, 2);
    source.setRecord(1, 0, { v: "gone", t: 1, s: "style-9" });
    const transaction = source.beginTransaction(REV(0));
    const prepared = transaction.prepareMutation({ type: "delete-rows", start: 1, count: 1 });
    prepared.commit();
    prepared.rollback();
    // rollback() closes over the captured records, so it is lossless. Only the
    // history path, which must serialise the inverse into one DataMutation,
    // is lossy. The gap is in the union's shapes, not in the record model.
    expect(source.getRecord(1, 0)).toEqual({ v: "gone", t: 1, s: "style-9" });
  });

  it("charges history bytes that grow with the number of restored records", () => {
    const { store, source } = createStore(3, 3);
    source.setRecord(0, 0, { v: "a" });
    const transaction = source.beginTransaction(REV(0));
    const single = transaction.prepareMutation({ type: "set-cell", row: 0, column: 0, value: "x" });
    const range = transaction.prepareMutation({
      type: "set-range",
      range: { startRow: 0, startColumn: 0, endRow: 2, endColumn: 2 },
      values: [
        [1, 2, 3],
        [4, 5, 6],
        [7, 8, 9]
      ]
    });
    expect(range.estimatedBytes).toBeGreaterThan(single.estimatedBytes);
    // The old hardcoded 64 did not move at all between these two.
    expect(single.estimatedBytes).not.toBe(64);
    expect(store.getRowCount()).toBe(3);
  });

  it("keeps a whole-axis reorder cheap in the history budget", () => {
    const { source } = createStore(1000, 50);
    const transaction = source.beginTransaction(REV(0));
    const order = Int32Array.from({ length: 1000 }, (_unused, index) => 999 - index);
    const prepared = transaction.prepareMutation({ type: "reorder-rows", order });
    // ADR 0003 §5.1: a sort whose inverse is a permutation costs O(rows * 4)
    // bytes, not O(occupied cells). Charging per moved record here would put
    // a 100k-row sort far past HistoryManager's 5 MB budget and empty the
    // undo stack, which is exactly the recorded incident.
    expect(prepared.estimatedBytes).toBeLessThan(10_000);
  });
});
