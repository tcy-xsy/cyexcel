import { describe, expect, it } from "vitest";

import {
  ClearRangeCommand,
  SetCellCommand,
  SetRangeCommand
} from "../../../../src/core/commands";
import { SetRowHeightsCommand } from "../../../../src/core/commands/SetRowHeightsCommand";
import { TableStore } from "../../../../src/core/data";
import { CyExcelError, CyExcelErrorCode } from "../../../../src/core/errors";
import { createTableCommandManager } from "../../../../src/core/integration";
import { createSheetFromRows } from "../../../helpers/createSheetModel";

function createStore() {
  const rows = [
    { name: "A", price: 1 },
    { name: "B", price: 2 }
  ];
  return new TableStore(createSheetFromRows(rows, ["name", "price"]));
}

describe("SetCellCommand", () => {
  it("writes the new value and produces one cell-value patch", () => {
    const store = createStore();
    const manager = createTableCommandManager(store);
    const command = new SetCellCommand(0, 1, 99);

    const patches = manager.execute(command);

    expect(patches).toEqual([{ type: "cell-value", row: 0, column: 1 }]);
    expect(store.getCell(0, 1)).toBe(99);
  });

  it("undo restores the value from before execute()", () => {
    const store = createStore();
    const manager = createTableCommandManager(store);
    const command = new SetCellCommand(0, 1, 99);

    manager.execute(command);
    manager.undo();

    expect(store.getCell(0, 1)).toBe(1);
  });
});

describe("SetRangeCommand", () => {
  it("writes the whole range and produces one range-value patch", () => {
    const store = createStore();
    const manager = createTableCommandManager(store);
    const range = { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 };
    const command = new SetRangeCommand(range, [
      ["A2", 10],
      ["B2", 20]
    ]);

    const patches = manager.execute(command);

    expect(patches).toEqual([{ type: "range-value", range }]);
    expect(store.getCell(0, 0)).toBe("A2");
    expect(store.getCell(1, 1)).toBe(20);
  });

  it("undo restores every cell in the range", () => {
    const store = createStore();
    const manager = createTableCommandManager(store);
    const range = { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 };
    const command = new SetRangeCommand(range, [
      ["A2", 10],
      ["B2", 20]
    ]);

    manager.execute(command);
    manager.undo();

    expect(store.getCell(0, 0)).toBe("A");
    expect(store.getCell(0, 1)).toBe(1);
    expect(store.getCell(1, 0)).toBe("B");
    expect(store.getCell(1, 1)).toBe(2);
  });
});

describe("ClearRangeCommand", () => {
  it("blanks every cell in the range", () => {
    const store = createStore();
    const manager = createTableCommandManager(store);
    const range = { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 };
    const command = new ClearRangeCommand(range);

    const patches = manager.execute(command);

    expect(patches).toEqual([{ type: "range-value", range }]);
    expect(store.getCell(0, 0)).toBeUndefined();
    expect(store.getCell(1, 1)).toBeUndefined();
  });

  it("undo restores the values that were cleared", () => {
    const store = createStore();
    const manager = createTableCommandManager(store);
    const range = { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 };
    const command = new ClearRangeCommand(range);

    manager.execute(command);
    manager.undo();

    expect(store.getCell(0, 0)).toBe("A");
    expect(store.getCell(1, 1)).toBe(2);
  });
});

describe("SetRowHeightsCommand", () => {
  it("snapshots and deeply freezes every update", () => {
    const updates = [{ row: 0, height: 28 }, { row: 1, height: 36 }];
    const command = new SetRowHeightsCommand(updates, "auto");

    updates[0]!.height = 99;
    updates.push({ row: 2, height: 48 });

    expect(command.params.updates).toEqual([
      { row: 0, height: 28 },
      { row: 1, height: 36 }
    ]);
    expect(Object.isFrozen(command)).toBe(true);
    expect(Object.isFrozen(command.params)).toBe(true);
    expect(Object.isFrozen(command.params.updates)).toBe(true);
    expect(command.params.updates.every((update) => Object.isFrozen(update))).toBe(true);
    expect(command.params.provenance).toBe("auto");
  });

  it.each([
    { name: "empty updates", updates: [] },
    { name: "a negative row", updates: [{ row: -1, height: 20 }] },
    { name: "a fractional row", updates: [{ row: 0.5, height: 20 }] },
    { name: "a zero height", updates: [{ row: 0, height: 0 }] },
    { name: "a non-finite height", updates: [{ row: 0, height: Number.POSITIVE_INFINITY }] },
    { name: "a duplicate row", updates: [{ row: 0, height: 20 }, { row: 0, height: 30 }] }
  ])("rejects $name", ({ updates }) => {
    try {
      new SetRowHeightsCommand(updates, "auto");
      throw new Error("Expected SetRowHeightsCommand construction to fail");
    } catch (error) {
      expect(error).toBeInstanceOf(CyExcelError);
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.InvalidOptions);
    }
  });
});
