import fc from "fast-check";
import { describe, expect, it } from "vitest";

import {
  DeleteColumnsCommand,
  InsertColumnsCommand,
  MoveColumnsCommand,
  SetCellCommand
} from "../../../../src/core/commands";
import type { BeforeColumnsChangeContext, BeforeTransactionCommitContext } from "../../../../src/core/commands";
import { ArrayDataSource, TableStore } from "../../../../src/core/data";
import { CyExcelError } from "../../../../src/core/errors";
import { HookManager, PRODUCTION_HOOK_REGISTRY } from "../../../../src/core/hooks";
import { createTableCommandManager } from "../../../../src/core/integration";

/** One row whose cell value in each column is that column's own uppercased field name — lets a test read back "which logical column is currently at position N", the column-axis analog of rowCommands.test.ts's readNames() trick. */
function createStore(fieldNames: readonly string[] = ["a", "b", "c"]) {
  const rows = [Object.fromEntries(fieldNames.map((field) => [field, field.toUpperCase()]))];
  const columns = fieldNames.map((field) => ({ id: field, field }));
  const source = new ArrayDataSource(rows, columns);
  return { store: new TableStore(source) };
}

function readColumnLabels(store: TableStore, count: number): string[] {
  return Array.from({ length: count }, (_unused, column) => store.getCell(0, column) as string);
}

describe("InsertColumnsCommand", () => {
  it("inserts the given columns and produces a column-structure insert patch", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);
    const command = new InsertColumnsCommand(1, [{ id: "x", field: "x" }]);

    const patches = manager.execute(command);

    expect(patches).toEqual([{ type: "column-structure", operation: "insert", start: 1, count: 1 }]);
    expect(store.getColumnCount()).toBe(4);
    expect(readColumnLabels(store, 4)).toEqual(["A", undefined, "B", "C"]);
  });

  it("undo removes exactly the columns it inserted", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);
    const command = new InsertColumnsCommand(1, [
      { id: "x", field: "x" },
      { id: "y", field: "y" }
    ]);
    manager.execute(command);

    manager.undo();

    expect(store.getColumnCount()).toBe(3);
    expect(readColumnLabels(store, 3)).toEqual(["A", "B", "C"]);
  });

  it("redo (execute again) reinserts the same columns", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);
    const command = new InsertColumnsCommand(1, [{ id: "x", field: "x" }]);
    manager.execute(command);
    manager.undo();

    manager.redo();

    expect(readColumnLabels(store, 4)).toEqual(["A", undefined, "B", "C"]);
  });
});

describe("DeleteColumnsCommand", () => {
  it("deletes the given columns and produces a column-structure delete patch", () => {
    const { store } = createStore(["a", "b", "c", "d"]);
    const manager = createTableCommandManager(store);
    const command = new DeleteColumnsCommand(1, 2);

    const patches = manager.execute(command);

    expect(patches).toEqual([{ type: "column-structure", operation: "delete", start: 1, count: 2 }]);
    expect(readColumnLabels(store, 2)).toEqual(["A", "D"]);
  });

  it("undo restores exactly the deleted columns' original data, with no captured data snapshot needed", () => {
    const { store } = createStore(["a", "b", "c", "d"]);
    const manager = createTableCommandManager(store);
    const command = new DeleteColumnsCommand(1, 2);
    manager.execute(command);

    manager.undo();

    expect(readColumnLabels(store, 4)).toEqual(["A", "B", "C", "D"]);
  });

  it("redo (execute again) deletes the same columns again", () => {
    const { store } = createStore(["a", "b", "c", "d"]);
    const manager = createTableCommandManager(store);
    const command = new DeleteColumnsCommand(1, 2);
    manager.execute(command);
    manager.undo();

    manager.redo();

    expect(readColumnLabels(store, 2)).toEqual(["A", "D"]);
  });
});

describe("MoveColumnsCommand", () => {
  it("moves the given block and produces a column-structure move patch", () => {
    const { store } = createStore(["a", "b", "c", "d", "e"]);
    const manager = createTableCommandManager(store);
    const command = new MoveColumnsCommand(1, 2, 4);

    const patches = manager.execute(command);

    expect(patches).toEqual([{ type: "column-structure", operation: "move", start: 1, count: 2, target: 4 }]);
    expect(readColumnLabels(store, 5)).toEqual(["A", "D", "B", "C", "E"]);
  });

  it("undo moves the block back to its original position", () => {
    const { store } = createStore(["a", "b", "c", "d", "e"]);
    const manager = createTableCommandManager(store);
    const command = new MoveColumnsCommand(1, 2, 4);
    manager.execute(command);

    manager.undo();

    expect(readColumnLabels(store, 5)).toEqual(["A", "B", "C", "D", "E"]);
  });

  it("redo (execute again) moves the block forward again", () => {
    const { store } = createStore(["a", "b", "c", "d", "e"]);
    const manager = createTableCommandManager(store);
    const command = new MoveColumnsCommand(1, 2, 4);
    manager.execute(command);
    manager.undo();

    manager.redo();

    expect(readColumnLabels(store, 5)).toEqual(["A", "D", "B", "C", "E"]);
  });

  it("undoing a move-earlier round-trips back to the original order", () => {
    const { store } = createStore(["a", "b", "c", "d", "e"]);
    const manager = createTableCommandManager(store);
    const command = new MoveColumnsCommand(3, 1, 0);
    manager.execute(command);
    expect(readColumnLabels(store, 5)).toEqual(["D", "A", "B", "C", "E"]);

    manager.undo();

    expect(readColumnLabels(store, 5)).toEqual(["A", "B", "C", "D", "E"]);
  });
});

describe("column commands — staged transaction view", () => {
  it("inserts a column and lets a later command address it before commit", () => {
    const { store } = createStore(["a", "b"]);
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    hooks.tap("beforeTransactionCommit", { name: "inspect-inserted-column" }, (context: BeforeTransactionCommitContext) => {
      expect(context.model.getColumnCount()).toBe(3);
      expect(context.model.getCell(0, 1)).toBe("X");
    });
    const manager = createTableCommandManager(store, { hooks });

    manager.transaction(() => {
      manager.execute(new InsertColumnsCommand(1, [{ id: "x", field: "x" }]));
      manager.execute(new SetCellCommand(0, 1, "X"));
      expect(store.getColumnCount()).toBe(2);
    });

    expect(Array.from({ length: 3 }, (_unused, column) => store.getCell(0, column))).toEqual([
      "A", "X", "B"
    ]);
    manager.undo();
    expect(Array.from({ length: 2 }, (_unused, column) => store.getCell(0, column))).toEqual(["A", "B"]);
    manager.redo();
    expect(Array.from({ length: 3 }, (_unused, column) => store.getCell(0, column))).toEqual([
      "A", "X", "B"
    ]);
  });

  it("moves columns and resolves a later set against the staged field mapping", () => {
    const { store } = createStore(["a", "b", "c"]);
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    hooks.tap("beforeTransactionCommit", { name: "inspect-moved-column" }, (context: BeforeTransactionCommitContext) => {
      expect(Array.from({ length: 3 }, (_unused, column) => context.model.getCell(0, column))).toEqual([
        "B", "C", "A2"
      ]);
    });
    const manager = createTableCommandManager(store, { hooks });

    manager.transaction(() => {
      manager.execute(new MoveColumnsCommand(0, 1, 3));
      manager.execute(new SetCellCommand(0, 2, "A2"));
    });

    expect(Array.from({ length: 3 }, (_unused, column) => store.getCell(0, column))).toEqual([
      "B", "C", "A2"
    ]);
    manager.undo();
    expect(Array.from({ length: 3 }, (_unused, column) => store.getCell(0, column))).toEqual([
      "A", "B", "C"
    ]);
  });
});

describe("MoveColumnsCommand — property: execute then undo always restores the original order", () => {
  it("holds for random (start, count, target) triples on a fixed-size sheet", () => {
    const COLUMN_COUNT = 7;
    fc.assert(
      fc.property(
        fc.integer({ min: 0, max: COLUMN_COUNT - 1 }),
        fc.integer({ min: 1, max: COLUMN_COUNT }),
        fc.integer({ min: 0, max: COLUMN_COUNT }),
        (start, rawCount, target) => {
          const count = Math.min(rawCount, COLUMN_COUNT - start);
          fc.pre(count >= 1);
          fc.pre(target < start || target > start + count - 1);

          const fieldNames = Array.from({ length: COLUMN_COUNT }, (_unused, i) => String.fromCharCode(97 + i));
          const { store } = createStore(fieldNames);
          const manager = createTableCommandManager(store);
          const command = new MoveColumnsCommand(start, count, target);

          manager.execute(command);
          manager.undo();

          expect(readColumnLabels(store, COLUMN_COUNT)).toEqual(
            fieldNames.map((field) => field.toUpperCase())
          );
        }
      )
    );
  });
});

describe("column commands — beforeColumnsChange cancellation leaves no partial commit", () => {
  function hooksThatCancel(): HookManager {
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    hooks.tap("beforeColumnsChange", { name: "block" }, (context: BeforeColumnsChangeContext) =>
      context.cancel("blocked in test")
    );
    return hooks;
  }

  it("InsertColumnsCommand: a cancelled insert writes nothing and records no history", () => {
    const { store } = createStore();
    const hooks = hooksThatCancel();
    const manager = createTableCommandManager(store, { hooks });

    expect(() =>
      manager.execute(new InsertColumnsCommand(1, [{ id: "x", field: "x" }]))
    ).toThrow(CyExcelError);

    expect(readColumnLabels(store, 3)).toEqual(["A", "B", "C"]);
    expect(manager.historySize).toBe(0);
  });

  it("DeleteColumnsCommand: a cancelled delete writes nothing and records no history", () => {
    const { store } = createStore();
    const hooks = hooksThatCancel();
    const manager = createTableCommandManager(store, { hooks });

    expect(() =>
      manager.execute(new DeleteColumnsCommand(0, 1))
    ).toThrow(CyExcelError);

    expect(readColumnLabels(store, 3)).toEqual(["A", "B", "C"]);
    expect(manager.historySize).toBe(0);
  });

  it("MoveColumnsCommand: a cancelled move writes nothing and records no history", () => {
    const { store } = createStore();
    const hooks = hooksThatCancel();
    const manager = createTableCommandManager(store, { hooks });

    expect(() => manager.execute(new MoveColumnsCommand(0, 1, 2))).toThrow(CyExcelError);

    expect(readColumnLabels(store, 3)).toEqual(["A", "B", "C"]);
    expect(manager.historySize).toBe(0);
  });

  it("a mid-transaction cancellation rolls back an already-executed column command in the same transaction", () => {
    const { store } = createStore(["a", "b", "c"]);
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    let calls = 0;
    hooks.tap("beforeColumnsChange", { name: "cancel-second" }, (context: BeforeColumnsChangeContext) => {
      calls += 1;
      if (calls === 2) {
        context.cancel();
      }
    });
    const manager = createTableCommandManager(store, { hooks });

    expect(() => {
      manager.transaction(() => {
        manager.execute(new InsertColumnsCommand(0, [{ id: "x", field: "x" }]));
        manager.execute(new DeleteColumnsCommand(0, 1));
      });
    }).toThrow(CyExcelError);

    expect(readColumnLabels(store, 3)).toEqual(["A", "B", "C"]);
    expect(manager.historySize).toBe(0);
  });
});
