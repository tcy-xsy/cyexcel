import { describe, expect, it } from "vitest";

import {
  SetCellCommand,
  SetRangeCommand
} from "../../../../src/core/commands";
import type {
  BeforeCommandContext,
  BeforeTransactionCommitContext
} from "../../../../src/core/commands";
import { ArrayDataSource, TableStore } from "../../../../src/core/data";
import { CyExcelError, CyExcelErrorCode } from "../../../../src/core/errors";
import {
  TableCommandManager,
  type TableCommandManagerOptions
} from "../../../../src/core/integration";
import { HookManager, PRODUCTION_HOOK_REGISTRY } from "../../../../src/core/hooks";
import { createStateMutation, type Mutation } from "../../../../src/core/mutations";
import type {
  Command,
  CommandContext,
  Patch
} from "../../../../src/core/types";

type TestPlanningCommand = Command<"test-plan", { readonly plan: (context: CommandContext) => void }>;

class CommandManager extends TableCommandManager {
  constructor(store: TableStore, options: TableCommandManagerOptions = {}) {
    super(store, {
      ...options,
      configureHandlers: (handlers) => {
        handlers.register<TestPlanningCommand>({
          type: "test-plan",
          plan: (params, context) => params.plan(context)
        });
      }
    });
  }
}

function createPlanningCommand(plan: (context: CommandContext) => void): TestPlanningCommand {
  return Object.freeze({ type: "test-plan", params: Object.freeze({ plan }) });
}

function createStateCommand(
  type: string,
  changes: readonly Patch[],
  apply: () => void,
  revert: () => void
): Command {
  return createPlanningCommand((context) => {
      context.stage(createStateMutation({ type, changes, apply, revert }));
    });
}

function createNoopCommand(type: string, onCommit: () => void, onDispose: () => void): Command {
  return createPlanningCommand((context) => {
    const mutation: Mutation = {
      type,
      changes: Object.freeze([{ type: "full-reset" }]),
      prepare: () => ({
        hasEffect: false,
        forward: mutation,
        inverse: mutation,
        estimatedBytes: 999,
        commit: onCommit,
        rollback: () => {},
        dispose: onDispose
      })
    };
    context.stage(mutation);
  });
}

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
  it("clearHistory empties both the undo and redo stacks", () => {
    const store = createStore();
    const manager = new CommandManager(store);

    manager.execute(new SetCellCommand(0, 0, "X"));
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

    manager.execute(new SetCellCommand(0, 1, 99));
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

    manager.execute(new SetRangeCommand(range, [["X", 10], ["Y", 20]]));

    expect(manager.historySize).toBe(1);
  });

  it("skip commits normally without recording a History entry", () => {
    const store = createStore();
    const manager = new CommandManager(store);

    manager.execute(new SetCellCommand(0, 0, "X"), { history: "skip" });

    expect(store.getCell(0, 0)).toBe("X");
    expect(store.version).toBe(1);
    expect(manager.lastChangeSet?.modelVersion).toBe(1);
    expect(manager.historySize).toBe(0);
    expect(manager.canUndo()).toBe(false);
  });

  it("skip preserves the redo stack after an undo", () => {
    const store = createStore();
    const manager = new CommandManager(store);

    manager.execute(new SetCellCommand(0, 0, "recorded"));
    manager.undo();
    expect(manager.canRedo()).toBe(true);

    manager.execute(new SetCellCommand(0, 1, 99), { history: "skip" });

    expect(store.getCell(0, 1)).toBe(99);
    expect(store.version).toBe(3);
    expect(manager.historySize).toBe(0);
    expect(manager.canRedo()).toBe(true);

    const redoPatches = manager.redo();

    expect(redoPatches).toEqual([{ type: "cell-value", row: 0, column: 0 }]);
    expect(store.getCell(0, 0)).toBe("recorded");
    expect(store.getCell(0, 1)).toBe(99);
    expect(store.version).toBe(4);
    expect(manager.historySize).toBe(1);
    expect(manager.canRedo()).toBe(false);
  });

  it("runs hooks for a no-op without advancing revision, recording history, or clearing redo", () => {
    const store = createStore();
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    const manager = new CommandManager(store, { hooks });
    manager.execute(new SetCellCommand(0, 0, "recorded"));
    manager.undo();
    expect(manager.canRedo()).toBe(true);
    expect(store.version).toBe(2);

    let beforeCommandCalls = 0;
    let beforeTransactionCommitCalls = 0;
    let commitCalls = 0;
    let disposeCalls = 0;
    hooks.tap("beforeCommand", { name: "observe-no-op-command" }, () => {
      beforeCommandCalls += 1;
    });
    hooks.tap("beforeTransactionCommit", { name: "observe-no-op-transaction" }, (context: BeforeTransactionCommitContext) => {
      beforeTransactionCommitCalls += 1;
      expect(context.patches).toEqual([]);
    });

    const patches = manager.execute(createNoopCommand(
      "explicit-no-op",
      () => {
        commitCalls += 1;
      },
      () => {
        disposeCalls += 1;
      }
    ));

    expect(patches).toEqual([]);
    expect(beforeCommandCalls).toBe(1);
    expect(beforeTransactionCommitCalls).toBe(1);
    expect(commitCalls).toBe(0);
    expect(disposeCalls).toBe(1);
    expect(store.version).toBe(2);
    expect(manager.lastChangeSet).toEqual({ modelVersion: 2, changes: [] });
    expect(manager.historySize).toBe(0);
    expect(manager.canUndo()).toBe(false);
    expect(manager.canRedo()).toBe(true);

    manager.redo();
    expect(store.getCell(0, 0)).toBe("recorded");
    expect(store.version).toBe(3);
  });
});

describe("CommandManager.transaction", () => {
  it("applies one skip policy to the whole transaction and lets nested calls inherit it", () => {
    const store = createStore();
    const manager = new CommandManager(store);

    manager.transaction(() => {
      manager.execute(new SetCellCommand(0, 0, "X"));
      manager.transaction(() => {
        manager.execute(new SetCellCommand(0, 1, 42));
      });
    }, { history: "skip" });

    expect(store.getCell(0, 0)).toBe("X");
    expect(store.getCell(0, 1)).toBe(42);
    expect(store.version).toBe(1);
    expect(manager.historySize).toBe(0);
    expect(manager.canUndo()).toBe(false);
  });

  it("rejects an execute policy that conflicts with the active transaction", () => {
    const store = createStore();
    const manager = new CommandManager(store);

    expect(() => manager.transaction(() => {
      manager.execute(new SetCellCommand(0, 0, "X"), { history: "record" });
    }, { history: "skip" })).toThrow(
      'Nested command history policy "record" conflicts with active transaction policy "skip"'
    );

    expect(store.getCell(0, 0)).toBe("A");
    expect(store.version).toBe(0);
    expect(manager.historySize).toBe(0);
  });

  it("rejects a nested transaction policy that conflicts with its outer transaction", () => {
    const store = createStore();
    const manager = new CommandManager(store);

    expect(() => manager.transaction(() => {
      manager.transaction(() => {}, { history: "skip" });
    }, { history: "record" })).toThrow(
      'Nested command history policy "skip" conflicts with active transaction policy "record"'
    );

    expect(store.version).toBe(0);
    expect(manager.historySize).toBe(0);
  });

  it("keeps duplicate structural facts in lastChangeSet even when the legacy patch adapter dedupes", () => {
    const store = createStore();
    const manager = new CommandManager(store);
    const structuralCommand = createPlanningCommand((context) => {
        context.stage(createStateMutation({
          type: "repeat-structure",
          changes: [
            { type: "row-structure", operation: "insert", start: 0, count: 1 },
            { type: "row-structure", operation: "insert", start: 0, count: 1 }
          ],
          apply: () => {},
          revert: () => {}
        }));
      });

    const patches = manager.execute(structuralCommand);

    expect(patches).toEqual([{ type: "row-structure", operation: "insert", start: 0, count: 1 }]);
    expect(manager.lastChangeSet?.changes).toEqual([
      { type: "row-structure", operation: "insert", start: 0, count: 1 },
      { type: "row-structure", operation: "insert", start: 0, count: 1 }
    ]);
  });

  it("merges nested transactions into the outer one", () => {
    const store = createStore();
    const manager = new CommandManager(store);

    manager.transaction(() => {
      manager.transaction(() => {
        manager.execute(new SetCellCommand(0, 0, "X"));
      });
      manager.transaction(() => {
        manager.execute(new SetCellCommand(0, 1, 42));
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

  it("lets later commands read mutations staged earlier in the same transaction", () => {
    const store = createStore();
    const manager = new CommandManager(store);
    const deriveFromStagedValue = createPlanningCommand((context) => {
        expect(context.model.baseRevision).toBe(0);
        expect(context.model.getCell(0, 0)).toBe("X");
        context.stage(store.createSetCellMutation(0, 1, `${context.model.getCell(0, 0)}!`));
      });

    manager.transaction(() => {
      manager.execute(new SetCellCommand(0, 0, "X"));
      expect(store.getCell(0, 0)).toBe("A");
      manager.execute(deriveFromStagedValue);
      expect(store.getCell(0, 1)).toBe(1);
    });

    expect(store.getCell(0, 0)).toBe("X");
    expect(store.getCell(0, 1)).toBe("X!");
    expect(store.version).toBe(1);

    manager.undo();
    expect(store.getCell(0, 0)).toBe("A");
    expect(store.getCell(0, 1)).toBe(1);

    manager.redo();
    expect(store.getCell(0, 0)).toBe("X");
    expect(store.getCell(0, 1)).toBe("X!");
  });

  it("rejects synchronous command re-entry from beforeCommand hooks", () => {
    const store = createStore();
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    const manager = new CommandManager(store, { hooks });
    hooks.tap("beforeCommand", { name: "reenter" }, () => {
      manager.execute(new SetCellCommand(0, 1, 99));
    });

    expect(() => manager.execute(new SetCellCommand(0, 0, "X"))).toThrow(CyExcelError);
    try {
      manager.execute(new SetCellCommand(0, 0, "X"));
    } catch (error) {
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.ReentrantMutation);
    }
    expect(store.getCell(0, 0)).toBe("A");
    expect(store.getCell(0, 1)).toBe(1);
    expect(store.version).toBe(0);
  });

  it("discards a prepared command when a read lease appears before commit", () => {
    const store = createStore();
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    const manager = new CommandManager(store, { hooks });
    let lease: ReturnType<typeof store.coordinator.revisionGate.acquireReadLease> | undefined;
    hooks.tap("beforeTransactionCommit", { name: "lease-before-commit" }, () => {
      lease = store.coordinator.revisionGate.acquireReadLease();
    });

    expect(() => manager.execute(new SetCellCommand(0, 0, "X"))).toThrow(CyExcelError);
    expect(store.coordinator.state).toBe("idle");
    expect(store.getCell(0, 0)).toBe("A");
    expect(store.version).toBe(0);
    expect(manager.historySize).toBe(0);
    lease?.release();
  });

  it("rolls back already-executed commands when the body throws", () => {
    const store = createStore();
    const manager = new CommandManager(store);

    expect(() => {
      manager.transaction(() => {
        manager.execute(new SetCellCommand(0, 0, "X"));
        throw new Error("boom");
      });
    }).toThrow("boom");

    expect(store.getCell(0, 0)).toBe("A");
    expect(store.version).toBe(0);
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

describe("CommandManager beforeCommand hook", () => {
  it("rejecting the first command leaves the store and history untouched", () => {
    const store = createStore();
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    hooks.tap("beforeCommand", { name: "reject-set-cell" }, (context: BeforeCommandContext) => {
      if (context.command.type === "set-cell") {
        context.cancel();
      }
    });
    const manager = new CommandManager(store, { hooks });

    expect(() => manager.execute(new SetCellCommand(0, 0, "X"))).toThrow(
      CyExcelError
    );
    expect(store.version).toBe(0);
    expect(manager.historySize).toBe(0);

    try {
      manager.execute(new SetCellCommand(0, 0, "X"));
      expect.unreachable("expected execute to throw");
    } catch (error) {
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.CommandCancelled);
    }
  });

  it("rejecting a later command rolls back earlier ones in the same transaction", () => {
    const store = createStore();
    let calls = 0;
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    hooks.tap("beforeCommand", { name: "reject-second-command" }, (context: BeforeCommandContext) => {
      calls += 1;
      if (calls === 2) {
        context.cancel();
      }
    });
    const manager = new CommandManager(store, { hooks });

    expect(() => {
      manager.transaction(() => {
        manager.execute(new SetCellCommand(0, 0, "X"));
        manager.execute(new SetCellCommand(0, 1, 42));
      });
    }).toThrow(CyExcelError);

    expect(store.getCell(0, 0)).toBe("A");
    expect(store.getCell(0, 1)).toBe(1);
    expect(manager.historySize).toBe(0);
  });
});

describe("CommandManager — shared HookManager", () => {
  it("runs hooks for skip and atomically discards a cancelled commit", () => {
    const store = createStore();
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    let beforeCommandCalls = 0;
    let beforeTransactionCommitCalls = 0;
    hooks.tap("beforeCommand", { name: "observe-skip-command" }, () => {
      beforeCommandCalls += 1;
    });
    hooks.tap("beforeTransactionCommit", { name: "cancel-skip-commit" }, (context: BeforeTransactionCommitContext) => {
      beforeTransactionCommitCalls += 1;
      expect(context.model.getCell(0, 0)).toBe("X");
      context.cancel("skip commit rejected");
    });
    const manager = new CommandManager(store, { hooks });

    expect(() => manager.execute(
      new SetCellCommand(0, 0, "X"),
      { history: "skip" }
    )).toThrow("skip commit rejected");

    expect(beforeCommandCalls).toBe(1);
    expect(beforeTransactionCommitCalls).toBe(1);
    expect(store.getCell(0, 0)).toBe("A");
    expect(store.version).toBe(0);
    expect(manager.lastChangeSet).toBeUndefined();
    expect(manager.historySize).toBe(0);
    expect(manager.canUndo()).toBe(false);
    expect(manager.canRedo()).toBe(false);
  });

  it("discards a staged transaction when a commit hook throws", () => {
    const store = createStore();
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    hooks.tap("beforeTransactionCommit", { name: "throwing-hook" }, () => {
      throw new Error("hook failed");
    });
    const manager = new CommandManager(store, { hooks });

    expect(() => manager.execute(new SetCellCommand(0, 0, "X"))).toThrow("hook failed");
    expect(store.getCell(0, 0)).toBe("A");
    expect(store.version).toBe(0);
    expect(manager.historySize).toBe(0);
    expect(manager.lastChangeSet).toBeUndefined();
  });

  it("exposes the complete staged preview to beforeTransactionCommit", () => {
    const store = createStore();
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    hooks.tap("beforeTransactionCommit", { name: "read-preview" }, (context: BeforeTransactionCommitContext) => {
      expect(context.model.baseRevision).toBe(0);
      expect(context.model.getCell(0, 0)).toBe("X");
      expect(store.getCell(0, 0)).toBe("A");
    });
    const manager = new CommandManager(store, { hooks });

    manager.execute(new SetCellCommand(0, 0, "X"));

    expect(store.getCell(0, 0)).toBe("X");
  });

  it("beforeCommand can reject a command before mutation planning", () => {
    const store = createStore();
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    hooks.tap("beforeCommand", { name: "reject-set-cell" }, (context: BeforeCommandContext) => {
      if (context.command.type === "set-cell") {
        context.cancel("no set-cell allowed");
      }
    });
    const manager = new CommandManager(store, { hooks });

    expect(() => manager.execute(new SetCellCommand(0, 0, "X"))).toThrow(CyExcelError);
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
        manager.execute(new SetCellCommand(0, 0, "X"));
        manager.execute(new SetCellCommand(0, 1, 42));
      });
    }).toThrow(CyExcelError);

    // Both commands had already been planned by the time beforeTransactionCommit ran.
    expect(seenCommandCount).toEqual([2]);
    // ...but the cancellation rolled both of them back.
    expect(store.getCell(0, 0)).toBe("A");
    expect(store.getCell(0, 1)).toBe(1);
    expect(store.version).toBe(0);
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

    manager.execute(new SetCellCommand(0, 0, "X"));

    expect(calls).toBe(1);
    expect(store.getCell(0, 0)).toBe("X");
    expect(manager.historySize).toBe(1);
  });
});

describe("CommandManager redo stack", () => {
  it("clears the redo stack once a new command is executed after undo", () => {
    const store = createStore();
    const manager = new CommandManager(store);

    manager.execute(new SetCellCommand(0, 0, "X"));
    manager.undo();
    expect(manager.canRedo()).toBe(true);

    manager.execute(new SetCellCommand(0, 1, 42));
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
    const command = createStateCommand(
      "single-patch",
      [{ type: "cell-value", row: 0, column: 0 }],
      () => {
        value = "after";
      },
      () => {
        value = "before";
      }
    );

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

  it("uses a fresh receipt when the same Command instance is executed twice", () => {
    const store = createStore();
    const manager = new CommandManager(store);
    let value = 0;
    const command = createPlanningCommand((context) => {
        const previous = value;
        const next = previous + 1;
        context.stage(createStateMutation({
          type: "increment",
          changes: [{ type: "full-reset" }],
          apply: () => {
            value = next;
          },
          revert: () => {
            value = previous;
          }
        }));
      });

    manager.execute(command);
    manager.execute(command);
    expect(value).toBe(2);

    manager.undo();
    expect(value).toBe(1);
    manager.undo();
    expect(value).toBe(0);
  });
});

describe("CommandManager History budgets", () => {
  it("evicts the oldest entries once historyLimit is exceeded", () => {
    const store = createStore();
    const manager = new CommandManager(store, { historyLimit: 3 });

    for (let i = 0; i < 5; i += 1) {
      manager.execute(new SetCellCommand(0, 0, i));
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
      manager.execute(new SetCellCommand(0, 0, i));
    }

    expect(manager.historySize).toBeGreaterThan(0);
    expect(manager.historySize).toBeLessThan(10);
  });
});
