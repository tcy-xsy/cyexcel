import fc from "fast-check";
import { describe, expect, it } from "vitest";

import {
  ArrayDataSource,
  CommandManager,
  CyExcelError,
  DeleteRowsCommand,
  HookManager,
  InsertRowsCommand,
  MoveRowsCommand,
  PRODUCTION_HOOK_REGISTRY,
  TableStore
} from "../../../../src/core";
import type { BeforeRowsChangeContext } from "../../../../src/core";

const CONTEXT = { transactionId: "t-test" };

function createStore(names: readonly string[] = ["A", "B", "C"]) {
  const rows = names.map((name) => ({ name }));
  const source = new ArrayDataSource(rows, [{ id: "name", field: "name" }]);
  return { rows, store: new TableStore(source) };
}

function readNames(store: TableStore, count: number): string[] {
  return Array.from({ length: count }, (_unused, row) => store.getCell(row, 0) as string);
}

describe("InsertRowsCommand", () => {
  it("inserts the given rows and produces a row-structure insert patch", () => {
    const { store } = createStore();
    const command = new InsertRowsCommand(store, 1, [["X"]]);

    const patches = command.execute(CONTEXT);

    expect(patches).toEqual([{ type: "row-structure", operation: "insert", start: 1, count: 1 }]);
    expect(readNames(store, 4)).toEqual(["A", "X", "B", "C"]);
  });

  it("undo removes exactly the rows it inserted", () => {
    const { store } = createStore();
    const command = new InsertRowsCommand(store, 1, [["X"], ["Y"]]);
    command.execute(CONTEXT);

    command.undo(CONTEXT);

    expect(readNames(store, 3)).toEqual(["A", "B", "C"]);
  });

  it("redo (execute again) reinserts the same rows", () => {
    const { store } = createStore();
    const command = new InsertRowsCommand(store, 1, [["X"]]);
    command.execute(CONTEXT);
    command.undo(CONTEXT);

    command.execute(CONTEXT);

    expect(readNames(store, 4)).toEqual(["A", "X", "B", "C"]);
  });
});

describe("DeleteRowsCommand", () => {
  it("deletes the given rows and produces a row-structure delete patch", () => {
    const { store } = createStore(["A", "B", "C", "D"]);
    const command = new DeleteRowsCommand(store, 1, 2);

    const patches = command.execute(CONTEXT);

    expect(patches).toEqual([{ type: "row-structure", operation: "delete", start: 1, count: 2 }]);
    expect(readNames(store, 2)).toEqual(["A", "D"]);
  });

  it("undo restores exactly the deleted rows' original values", () => {
    const { store } = createStore(["A", "B", "C", "D"]);
    const command = new DeleteRowsCommand(store, 1, 2);
    command.execute(CONTEXT);

    command.undo(CONTEXT);

    expect(readNames(store, 4)).toEqual(["A", "B", "C", "D"]);
  });

  it("redo (execute again) deletes the same rows again", () => {
    const { store } = createStore(["A", "B", "C", "D"]);
    const command = new DeleteRowsCommand(store, 1, 2);
    command.execute(CONTEXT);
    command.undo(CONTEXT);

    command.execute(CONTEXT);

    expect(readNames(store, 2)).toEqual(["A", "D"]);
  });
});

describe("MoveRowsCommand", () => {
  it("moves the given block and produces a row-structure move patch", () => {
    const { store } = createStore(["A", "B", "C", "D", "E"]);
    const command = new MoveRowsCommand(store, 1, 2, 4);

    const patches = command.execute(CONTEXT);

    expect(patches).toEqual([{ type: "row-structure", operation: "move", start: 1, count: 2, target: 4 }]);
    expect(readNames(store, 5)).toEqual(["A", "D", "B", "C", "E"]);
  });

  it("undo moves the block back to its original position", () => {
    const { store } = createStore(["A", "B", "C", "D", "E"]);
    const command = new MoveRowsCommand(store, 1, 2, 4);
    command.execute(CONTEXT);

    command.undo(CONTEXT);

    expect(readNames(store, 5)).toEqual(["A", "B", "C", "D", "E"]);
  });

  it("redo (execute again) moves the block forward again", () => {
    const { store } = createStore(["A", "B", "C", "D", "E"]);
    const command = new MoveRowsCommand(store, 1, 2, 4);
    command.execute(CONTEXT);
    command.undo(CONTEXT);

    command.execute(CONTEXT);

    expect(readNames(store, 5)).toEqual(["A", "D", "B", "C", "E"]);
  });

  it("preserves each moved row's own object identity", () => {
    const rowB = { name: "B" };
    const rows = [{ name: "A" }, rowB, { name: "C" }];
    const source = new ArrayDataSource(rows, [{ id: "name", field: "name" }]);
    const store = new TableStore(source);
    const command = new MoveRowsCommand(store, 1, 1, 3);

    command.execute(CONTEXT);

    expect(rows[2]).toBe(rowB);
  });

  it("undoing a move-earlier round-trips back to the original order", () => {
    const { store } = createStore(["A", "B", "C", "D", "E"]);
    const command = new MoveRowsCommand(store, 3, 1, 0);
    command.execute(CONTEXT);
    expect(readNames(store, 5)).toEqual(["D", "A", "B", "C", "E"]);

    command.undo(CONTEXT);

    expect(readNames(store, 5)).toEqual(["A", "B", "C", "D", "E"]);
  });
});

describe("MoveRowsCommand — property: execute then undo always restores the original order", () => {
  it("holds for random (start, count, target) triples on a fixed-size sheet", () => {
    const ROW_COUNT = 7;
    fc.assert(
      fc.property(
        fc.integer({ min: 0, max: ROW_COUNT - 1 }),
        fc.integer({ min: 1, max: ROW_COUNT }),
        fc.integer({ min: 0, max: ROW_COUNT }),
        (start, rawCount, target) => {
          const count = Math.min(rawCount, ROW_COUNT - start);
          fc.pre(count >= 1);
          fc.pre(target < start || target > start + count - 1); // target not inside the block itself

          const original = Array.from({ length: ROW_COUNT }, (_unused, i) => String.fromCharCode(65 + i));
          const { store } = createStore(original);
          const command = new MoveRowsCommand(store, start, count, target);

          command.execute(CONTEXT);
          command.undo(CONTEXT);

          expect(readNames(store, ROW_COUNT)).toEqual(original);
        }
      )
    );
  });
});

describe("row commands — beforeRowsChange cancellation leaves no partial commit", () => {
  function hooksThatCancel(): HookManager {
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    hooks.tap("beforeRowsChange", { name: "block" }, (context: BeforeRowsChangeContext) =>
      context.cancel("blocked in test")
    );
    return hooks;
  }

  it("InsertRowsCommand: a cancelled insert writes nothing and records no history", () => {
    const { store } = createStore();
    const hooks = hooksThatCancel();
    const manager = new CommandManager(store, { hooks });

    expect(() => manager.execute(new InsertRowsCommand(store, 1, [["X"]], hooks))).toThrow(CyExcelError);

    expect(readNames(store, 3)).toEqual(["A", "B", "C"]);
    expect(manager.historySize).toBe(0);
  });

  it("DeleteRowsCommand: a cancelled delete writes nothing and records no history", () => {
    const { store } = createStore();
    const hooks = hooksThatCancel();
    const manager = new CommandManager(store, { hooks });

    expect(() => manager.execute(new DeleteRowsCommand(store, 0, 1, hooks))).toThrow(CyExcelError);

    expect(readNames(store, 3)).toEqual(["A", "B", "C"]);
    expect(manager.historySize).toBe(0);
  });

  it("MoveRowsCommand: a cancelled move writes nothing and records no history", () => {
    const { store } = createStore();
    const hooks = hooksThatCancel();
    const manager = new CommandManager(store, { hooks });

    expect(() => manager.execute(new MoveRowsCommand(store, 0, 1, 2, hooks))).toThrow(CyExcelError);

    expect(readNames(store, 3)).toEqual(["A", "B", "C"]);
    expect(manager.historySize).toBe(0);
  });

  it("a mid-transaction cancellation rolls back an already-executed row command in the same transaction", () => {
    const { store } = createStore(["A", "B", "C"]);
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    let calls = 0;
    hooks.tap("beforeRowsChange", { name: "cancel-second" }, (context: BeforeRowsChangeContext) => {
      calls += 1;
      if (calls === 2) {
        context.cancel();
      }
    });
    const manager = new CommandManager(store, { hooks });

    expect(() => {
      manager.transaction(() => {
        manager.execute(new InsertRowsCommand(store, 0, [["X"]], hooks));
        manager.execute(new DeleteRowsCommand(store, 0, 1, hooks));
      });
    }).toThrow(CyExcelError);

    expect(readNames(store, 3)).toEqual(["A", "B", "C"]);
    expect(manager.historySize).toBe(0);
  });
});
