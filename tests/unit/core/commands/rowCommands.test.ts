import fc from "fast-check";
import { describe, expect, it } from "vitest";

import {
  DeleteRowsCommand,
  InsertRowsCommand,
  MoveRowsCommand,
  SetCellCommand
} from "../../../../src/core/commands";
import type { BeforeRowsChangeContext, BeforeTransactionCommitContext } from "../../../../src/core/commands";
import { TableStore } from "../../../../src/core/data";
import { CyExcelError } from "../../../../src/core/errors";
import { HookManager, PRODUCTION_HOOK_REGISTRY } from "../../../../src/core/hooks";
import { createTableCommandManager } from "../../../../src/core/integration";
import { createSheetFromRows } from "../../../helpers/createSheetModel";

function createStore(names: readonly string[] = ["A", "B", "C"]) {
  const rows = names.map((name) => ({ name }));
  return { rows, store: new TableStore(createSheetFromRows(rows, ["name"])) };
}

function readNames(store: TableStore, count: number): string[] {
  return Array.from({ length: count }, (_unused, row) => store.getCell(row, 0) as string);
}

describe("InsertRowsCommand", () => {
  it("inserts the given rows and produces a row-structure insert patch", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);
    const command = new InsertRowsCommand(1, [["X"]]);

    const patches = manager.execute(command);

    expect(patches).toEqual([{ type: "row-structure", operation: "insert", start: 1, count: 1 }]);
    expect(readNames(store, 4)).toEqual(["A", "X", "B", "C"]);
  });

  it("undo removes exactly the rows it inserted", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);
    const command = new InsertRowsCommand(1, [["X"], ["Y"]]);
    manager.execute(command);

    manager.undo();

    expect(readNames(store, 3)).toEqual(["A", "B", "C"]);
  });

  /**
   * 追加一行空行：一个格子都不用搬，坐标空间只能靠 `blankRows` 涨。
   *
   * 从前它原地不动——表还是三行——而布局那边已经是四行了；撤销于是拿
   * `delete-rows 3,1` 去撞一张三行的表，`validateAxisChange` 直接抛
   * `delete range is outside axis length`。
   */
  it("appends an empty row at the end and undoes it", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);

    manager.execute(new InsertRowsCommand(3, [[undefined]]));
    expect(store.getRowCount()).toBe(4);
    expect(readNames(store, 4)).toEqual(["A", "B", "C", undefined]);

    manager.undo();
    expect(store.getRowCount()).toBe(3);
    expect(readNames(store, 3)).toEqual(["A", "B", "C"]);
  });

  it("redo (execute again) reinserts the same rows", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);
    const command = new InsertRowsCommand(1, [["X"]]);
    manager.execute(command);
    manager.undo();

    manager.redo();

    expect(readNames(store, 4)).toEqual(["A", "X", "B", "C"]);
  });
});

describe("DeleteRowsCommand", () => {
  it("deletes the given rows and produces a row-structure delete patch", () => {
    const { store } = createStore(["A", "B", "C", "D"]);
    const manager = createTableCommandManager(store);
    const command = new DeleteRowsCommand(1, 2);

    const patches = manager.execute(command);

    expect(patches).toEqual([{ type: "row-structure", operation: "delete", start: 1, count: 2 }]);
    expect(readNames(store, 2)).toEqual(["A", "D"]);
  });

  it("undo restores exactly the deleted rows' original values", () => {
    const { store } = createStore(["A", "B", "C", "D"]);
    const manager = createTableCommandManager(store);
    const command = new DeleteRowsCommand(1, 2);
    manager.execute(command);

    manager.undo();

    expect(readNames(store, 4)).toEqual(["A", "B", "C", "D"]);
  });

  it("redo (execute again) deletes the same rows again", () => {
    const { store } = createStore(["A", "B", "C", "D"]);
    const manager = createTableCommandManager(store);
    const command = new DeleteRowsCommand(1, 2);
    manager.execute(command);
    manager.undo();

    manager.redo();

    expect(readNames(store, 2)).toEqual(["A", "D"]);
  });
});

describe("MoveRowsCommand", () => {
  it("moves the given block and produces a row-structure move patch", () => {
    const { store } = createStore(["A", "B", "C", "D", "E"]);
    const manager = createTableCommandManager(store);
    const command = new MoveRowsCommand(1, 2, 4);

    const patches = manager.execute(command);

    expect(patches).toEqual([{ type: "row-structure", operation: "move", start: 1, count: 2, target: 4 }]);
    expect(readNames(store, 5)).toEqual(["A", "D", "B", "C", "E"]);
  });

  it("undo moves the block back to its original position", () => {
    const { store } = createStore(["A", "B", "C", "D", "E"]);
    const manager = createTableCommandManager(store);
    const command = new MoveRowsCommand(1, 2, 4);
    manager.execute(command);

    manager.undo();

    expect(readNames(store, 5)).toEqual(["A", "B", "C", "D", "E"]);
  });

  it("redo (execute again) moves the block forward again", () => {
    const { store } = createStore(["A", "B", "C", "D", "E"]);
    const manager = createTableCommandManager(store);
    const command = new MoveRowsCommand(1, 2, 4);
    manager.execute(command);
    manager.undo();

    manager.redo();

    expect(readNames(store, 5)).toEqual(["A", "D", "B", "C", "E"]);
  });

  // 行对象不再是存储（ADR 0004 §3.1），所以"行的身份"这件事没有了；仍然成立
  // 而且仍然要钉的是**单元格值**的身份：复杂值是原子的（ADR 0001 §5），移动一段
  // 行不能把值复制一份。
  it("moves complex cell values by reference, not by copy", () => {
    const marker = { id: "row-b" };
    const store = new TableStore(
      createSheetFromRows(
        [{ name: "A" }, { name: "B", marker }, { name: "C" }],
        ["name", "marker"]
      )
    );
    const manager = createTableCommandManager(store);
    const command = new MoveRowsCommand(1, 1, 3);

    manager.execute(command);

    expect(store.getCell(2, 0)).toBe("B");
    expect(store.getCell(2, 1)).toBe(marker);
  });

  it("undoing a move-earlier round-trips back to the original order", () => {
    const { store } = createStore(["A", "B", "C", "D", "E"]);
    const manager = createTableCommandManager(store);
    const command = new MoveRowsCommand(3, 1, 0);
    manager.execute(command);
    expect(readNames(store, 5)).toEqual(["D", "A", "B", "C", "E"]);

    manager.undo();

    expect(readNames(store, 5)).toEqual(["A", "B", "C", "D", "E"]);
  });
});

describe("row commands — staged transaction view", () => {
  it("deletes a row, writes the shifted row, and restores exact values on undo", () => {
    // 导入用的行对象在这里被留着，钉的是**反过来**的一条：写回通道已经切断
    // （4B 设计 §3.3），表格改的是自己的存储，这三个对象一个字节都不会动。
    const rowA = { name: "A" };
    const rowB = { name: "B" };
    const rowC = { name: "C" };
    const store = new TableStore(createSheetFromRows([rowA, rowB, rowC], ["name"]));
    const manager = createTableCommandManager(store);

    manager.transaction(() => {
      manager.execute(new DeleteRowsCommand(0, 1));
      manager.execute(new SetCellCommand(0, 0, "B2"));
    });

    expect(readNames(store, 2)).toEqual(["B2", "C"]);
    expect(rowB.name).toBe("B");

    manager.undo();
    expect(readNames(store, 3)).toEqual(["A", "B", "C"]);

    manager.redo();
    expect(readNames(store, 2)).toEqual(["B2", "C"]);
    expect([rowA.name, rowB.name, rowC.name]).toEqual(["A", "B", "C"]);
  });

  it("can read and write an inserted row before the transaction commits", () => {
    const { store } = createStore(["A"]);
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    hooks.tap("beforeTransactionCommit", { name: "inspect-inserted-row" }, (context: BeforeTransactionCommitContext) => {
      expect(context.model.getRowCount()).toBe(2);
      expect(context.model.getCell(1, 0)).toBe("B2");
    });
    const manager = createTableCommandManager(store, { hooks });

    manager.transaction(() => {
      manager.execute(new InsertRowsCommand(1, [["B"]]));
      manager.execute(new SetCellCommand(1, 0, "B2"));
      expect(store.getRowCount()).toBe(1);
    });

    expect(readNames(store, 2)).toEqual(["A", "B2"]);
    manager.undo();
    expect(readNames(store, 1)).toEqual(["A"]);
    manager.redo();
    expect(readNames(store, 2)).toEqual(["A", "B2"]);
  });

  it("moves rows and resolves later reads and writes against the staged order", () => {
    const { store } = createStore(["A", "B", "C", "D"]);
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    hooks.tap("beforeTransactionCommit", { name: "inspect-moved-row" }, (context: BeforeTransactionCommitContext) => {
      expect(Array.from({ length: 4 }, (_unused, row) => context.model.getCell(row, 0))).toEqual([
        "A", "C", "B2", "D"
      ]);
    });
    const manager = createTableCommandManager(store, { hooks });

    manager.transaction(() => {
      manager.execute(new MoveRowsCommand(1, 1, 3));
      manager.execute(new SetCellCommand(2, 0, "B2"));
    });

    expect(readNames(store, 4)).toEqual(["A", "C", "B2", "D"]);
    manager.undo();
    expect(readNames(store, 4)).toEqual(["A", "B", "C", "D"]);
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
          const manager = createTableCommandManager(store);
          const command = new MoveRowsCommand(start, count, target);

          manager.execute(command);
          manager.undo();

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
    const manager = createTableCommandManager(store, { hooks });

    expect(() => manager.execute(new InsertRowsCommand(1, [["X"]]))).toThrow(CyExcelError);

    expect(readNames(store, 3)).toEqual(["A", "B", "C"]);
    expect(manager.historySize).toBe(0);
  });

  it("DeleteRowsCommand: a cancelled delete writes nothing and records no history", () => {
    const { store } = createStore();
    const hooks = hooksThatCancel();
    const manager = createTableCommandManager(store, { hooks });

    expect(() => manager.execute(new DeleteRowsCommand(0, 1))).toThrow(CyExcelError);

    expect(readNames(store, 3)).toEqual(["A", "B", "C"]);
    expect(manager.historySize).toBe(0);
  });

  it("MoveRowsCommand: a cancelled move writes nothing and records no history", () => {
    const { store } = createStore();
    const hooks = hooksThatCancel();
    const manager = createTableCommandManager(store, { hooks });

    expect(() => manager.execute(new MoveRowsCommand(0, 1, 2))).toThrow(CyExcelError);

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
    const manager = createTableCommandManager(store, { hooks });

    expect(() => {
      manager.transaction(() => {
        manager.execute(new InsertRowsCommand(0, [["X"]]));
        manager.execute(new DeleteRowsCommand(0, 1));
      });
    }).toThrow(CyExcelError);

    expect(readNames(store, 3)).toEqual(["A", "B", "C"]);
    expect(manager.historySize).toBe(0);
  });
});
