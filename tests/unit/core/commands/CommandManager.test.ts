import { describe, expect, it } from "vitest";

import {
  ArrayDataSource,
  CommandManager,
  CyExcelError,
  CyExcelErrorCode,
  HookManager,
  PRODUCTION_HOOK_REGISTRY,
  SetCellCommand,
  SetRangeCommand,
  TableStore
} from "../../../../src/core";
import type { BeforeCommandContext, BeforeTransactionCommitContext, Command } from "../../../../src/core";

function createStore() {
  const rows = [
    { name: "A", price: 1 },
    { name: "B", price: 2 }
  ];
  const source = new ArrayDataSource(rows, [
    { id: "name", field: "name" },
    { id: "price", field: "price" }
  ]);
  return new TableStore(source);
}

describe("CommandManager basics", () => {
  it("exposes the store it was constructed with", () => {
    const store = createStore();
    const manager = new CommandManager(store);

    expect(manager.store).toBe(store);
  });

  it("clearHistory empties both the undo and redo stacks", () => {
    const store = createStore();
    const manager = new CommandManager(store);

    manager.execute(new SetCellCommand(store, 0, 0, "X"));
    manager.undo();
    expect(manager.canRedo()).toBe(true);

    manager.clearHistory();

    expect(manager.canUndo()).toBe(false);
    expect(manager.canRedo()).toBe(false);
  });
});

describe("CommandManager.execute", () => {
  it("execute/undo/redo keep the cell value and store.version correct", () => {
    const store = createStore();
    const manager = new CommandManager(store);

    manager.execute(new SetCellCommand(store, 0, 1, 99));
    expect(store.getCell(0, 1)).toBe(99);
    expect(store.version).toBe(1);

    manager.undo();
    expect(store.getCell(0, 1)).toBe(1);
    expect(store.version).toBe(2);

    manager.redo();
    expect(store.getCell(0, 1)).toBe(99);
    expect(store.version).toBe(3);
  });

  it("a single SetRangeCommand forms exactly one History entry", () => {
    const store = createStore();
    const manager = new CommandManager(store);
    const range = { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 };

    manager.execute(new SetRangeCommand(store, range, [["X", 10], ["Y", 20]]));

    expect(manager.historySize).toBe(1);
  });
});

describe("CommandManager.transaction", () => {
  it("merges nested transactions into the outer one", () => {
    const store = createStore();
    const manager = new CommandManager(store);

    manager.transaction(() => {
      manager.transaction(() => {
        manager.execute(new SetCellCommand(store, 0, 0, "X"));
      });
      manager.transaction(() => {
        manager.execute(new SetCellCommand(store, 0, 1, 42));
      });
    });

    expect(manager.historySize).toBe(1);
    expect(store.getCell(0, 0)).toBe("X");
    expect(store.getCell(0, 1)).toBe(42);

    manager.undo();
    expect(store.getCell(0, 0)).toBe("A");
    expect(store.getCell(0, 1)).toBe(1);
    expect(manager.historySize).toBe(0);
  });

  it("rolls back already-executed commands when the body throws", () => {
    const store = createStore();
    const manager = new CommandManager(store);

    expect(() => {
      manager.transaction(() => {
        manager.execute(new SetCellCommand(store, 0, 0, "X"));
        throw new Error("boom");
      });
    }).toThrow("boom");

    expect(store.getCell(0, 0)).toBe("A");
    expect(manager.historySize).toBe(0);
  });

  it("empty transactions do not create a History entry", () => {
    const store = createStore();
    const manager = new CommandManager(store);

    manager.transaction(() => {
      // no commands executed
    });

    expect(manager.historySize).toBe(0);
  });
});

describe("CommandManager beforeCommit", () => {
  it("rejecting the first command leaves the store and history untouched", () => {
    const store = createStore();
    const manager = new CommandManager(store, {
      beforeCommit: command => command.type !== "set-cell"
    });

    expect(() => manager.execute(new SetCellCommand(store, 0, 0, "X"))).toThrow(
      CyExcelError
    );
    expect(store.version).toBe(0);
    expect(manager.historySize).toBe(0);

    try {
      manager.execute(new SetCellCommand(store, 0, 0, "X"));
      expect.unreachable("expected execute to throw");
    } catch (error) {
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.CommandCancelled);
    }
  });

  it("rejecting a later command rolls back earlier ones in the same transaction", () => {
    const store = createStore();
    let calls = 0;
    const manager = new CommandManager(store, {
      beforeCommit: () => {
        calls += 1;
        return calls < 2;
      }
    });

    expect(() => {
      manager.transaction(() => {
        manager.execute(new SetCellCommand(store, 0, 0, "X"));
        manager.execute(new SetCellCommand(store, 0, 1, 42));
      });
    }).toThrow(CyExcelError);

    expect(store.getCell(0, 0)).toBe("A");
    expect(store.getCell(0, 1)).toBe(1);
    expect(manager.historySize).toBe(0);
  });
});

describe("CommandManager — shared HookManager", () => {
  it("hooks.tap('beforeCommand', ...) is equivalent to the beforeCommit option", () => {
    const store = createStore();
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    hooks.tap("beforeCommand", { name: "reject-set-cell" }, (context: BeforeCommandContext) => {
      if (context.command.type === "set-cell") {
        context.cancel("no set-cell allowed");
      }
    });
    const manager = new CommandManager(store, { hooks });

    expect(() => manager.execute(new SetCellCommand(store, 0, 0, "X"))).toThrow(CyExcelError);
    expect(store.version).toBe(0);
  });

  it("beforeTransactionCommit cancels and rolls back the whole transaction after every command already ran", () => {
    const store = createStore();
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    const seenCommandCount: number[] = [];
    hooks.tap("beforeTransactionCommit", { name: "cross-cell-check" }, (context: BeforeTransactionCommitContext) => {
      seenCommandCount.push(context.commands.length);
      context.cancel("cross-cell validation failed");
    });
    const manager = new CommandManager(store, { hooks });

    expect(() => {
      manager.transaction(() => {
        manager.execute(new SetCellCommand(store, 0, 0, "X"));
        manager.execute(new SetCellCommand(store, 0, 1, 42));
      });
    }).toThrow(CyExcelError);

    // Both commands had already executed by the time beforeTransactionCommit ran.
    expect(seenCommandCount).toEqual([2]);
    // ...but the cancellation rolled both of them back.
    expect(store.getCell(0, 0)).toBe("A");
    expect(store.getCell(0, 1)).toBe(1);
    expect(manager.historySize).toBe(0);
  });

  it("a transaction that isn't cancelled commits normally through beforeTransactionCommit", () => {
    const store = createStore();
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    let calls = 0;
    hooks.tap("beforeTransactionCommit", { name: "observe-only" }, () => {
      calls += 1;
    });
    const manager = new CommandManager(store, { hooks });

    manager.execute(new SetCellCommand(store, 0, 0, "X"));

    expect(calls).toBe(1);
    expect(store.getCell(0, 0)).toBe("X");
    expect(manager.historySize).toBe(1);
  });
});

describe("CommandManager redo stack", () => {
  it("clears the redo stack once a new command is executed after undo", () => {
    const store = createStore();
    const manager = new CommandManager(store);

    manager.execute(new SetCellCommand(store, 0, 0, "X"));
    manager.undo();
    expect(manager.canRedo()).toBe(true);

    manager.execute(new SetCellCommand(store, 0, 1, 42));
    expect(manager.canRedo()).toBe(false);
  });
});

describe("CommandManager undo/redo edge cases", () => {
  it("undo/redo return undefined when there is nothing to undo/redo", () => {
    const store = createStore();
    const manager = new CommandManager(store);

    expect(manager.undo()).toBeUndefined();
    expect(manager.redo()).toBeUndefined();
  });

  it("accepts a command whose execute/undo return a single Patch instead of an array", () => {
    const store = createStore();
    const manager = new CommandManager(store);
    let value = "before";
    const command: Command = {
      type: "single-patch",
      execute: () => {
        value = "after";
        return { type: "cell-value", row: 0, column: 0 };
      },
      undo: () => {
        value = "before";
        return { type: "cell-value", row: 0, column: 0 };
      }
    };

    const executePatches = manager.execute(command);
    expect(executePatches).toEqual([{ type: "cell-value", row: 0, column: 0 }]);
    expect(value).toBe("after");

    const undoPatches = manager.undo();
    expect(undoPatches).toEqual([{ type: "cell-value", row: 0, column: 0 }]);
    expect(value).toBe("before");

    const redoPatches = manager.redo();
    expect(redoPatches).toEqual([{ type: "cell-value", row: 0, column: 0 }]);
    expect(value).toBe("after");
  });

  it("skips commands without an undo() when undoing a transaction", () => {
    const store = createStore();
    const manager = new CommandManager(store);
    const command: Command = {
      type: "no-undo",
      execute: () => ({ type: "full-reset" })
    };

    manager.execute(command);

    expect(() => manager.undo()).not.toThrow();
  });
});

describe("CommandManager History budgets", () => {
  it("evicts the oldest entries once historyLimit is exceeded", () => {
    const store = createStore();
    const manager = new CommandManager(store, { historyLimit: 3 });

    for (let i = 0; i < 5; i += 1) {
      manager.execute(new SetCellCommand(store, 0, 0, i));
    }

    expect(manager.historySize).toBe(3);

    while (manager.canUndo()) {
      manager.undo();
    }

    // Entries for values 0 and 1 were evicted; undoing the retained three
    // (2, 3, 4) lands back on the value written right before 2 — i.e. 1.
    expect(store.getCell(0, 0)).toBe(1);
  });

  it("evicts the oldest entries once the byte budget is exceeded", () => {
    const store = createStore();
    const manager = new CommandManager(store, {
      historyLimit: 100,
      historyByteLimit: 90
    });

    for (let i = 0; i < 10; i += 1) {
      manager.execute(new SetCellCommand(store, 0, 0, i));
    }

    expect(manager.historySize).toBeGreaterThan(0);
    expect(manager.historySize).toBeLessThan(10);
  });
});
