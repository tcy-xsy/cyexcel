import { describe, expect, it } from "vitest";

import {
  ArrayDataSource,
  ClearRangeCommand,
  SetCellCommand,
  SetRangeCommand,
  TableStore
} from "../../../../src/core";

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

const CONTEXT = { transactionId: "t-test" };

describe("SetCellCommand", () => {
  it("writes the new value and produces one cell-value patch", () => {
    const store = createStore();
    const command = new SetCellCommand(store, 0, 1, 99);

    const patches = command.execute(CONTEXT);

    expect(patches).toEqual([{ type: "cell-value", row: 0, column: 1 }]);
    expect(store.getCell(0, 1)).toBe(99);
  });

  it("undo restores the value from before execute()", () => {
    const store = createStore();
    const command = new SetCellCommand(store, 0, 1, 99);

    command.execute(CONTEXT);
    command.undo(CONTEXT);

    expect(store.getCell(0, 1)).toBe(1);
  });
});

describe("SetRangeCommand", () => {
  it("writes the whole range and produces one range-value patch", () => {
    const store = createStore();
    const range = { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 };
    const command = new SetRangeCommand(store, range, [
      ["A2", 10],
      ["B2", 20]
    ]);

    const patches = command.execute(CONTEXT);

    expect(patches).toEqual([{ type: "range-value", range }]);
    expect(store.getCell(0, 0)).toBe("A2");
    expect(store.getCell(1, 1)).toBe(20);
  });

  it("undo restores every cell in the range", () => {
    const store = createStore();
    const range = { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 };
    const command = new SetRangeCommand(store, range, [
      ["A2", 10],
      ["B2", 20]
    ]);

    command.execute(CONTEXT);
    command.undo(CONTEXT);

    expect(store.getCell(0, 0)).toBe("A");
    expect(store.getCell(0, 1)).toBe(1);
    expect(store.getCell(1, 0)).toBe("B");
    expect(store.getCell(1, 1)).toBe(2);
  });
});

describe("ClearRangeCommand", () => {
  it("blanks every cell in the range", () => {
    const store = createStore();
    const range = { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 };
    const command = new ClearRangeCommand(store, range);

    const patches = command.execute(CONTEXT);

    expect(patches).toEqual([{ type: "range-value", range }]);
    expect(store.getCell(0, 0)).toBeUndefined();
    expect(store.getCell(1, 1)).toBeUndefined();
  });

  it("undo restores the values that were cleared", () => {
    const store = createStore();
    const range = { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 };
    const command = new ClearRangeCommand(store, range);

    command.execute(CONTEXT);
    command.undo(CONTEXT);

    expect(store.getCell(0, 0)).toBe("A");
    expect(store.getCell(1, 1)).toBe(2);
  });
});
