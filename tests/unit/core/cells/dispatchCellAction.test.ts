import { describe, expect, it } from "vitest";

import { dispatchCellAction } from "../../../../src/core/commands";
import { ArrayDataSource, TableStore } from "../../../../src/core/data";
import { createTableCommandManager } from "../../../../src/core/integration";

function createStore(): TableStore {
  const dataSource = new ArrayDataSource(
    [{ name: "old" }],
    [{ id: "name", field: "name" }]
  );
  return new TableStore(dataSource);
}

describe("dispatchCellAction — hitTest CellAction to Command translation", () => {
  it("translates a set-value action into an executable SetCellCommand", () => {
    const store = createStore();
    const manager = createTableCommandManager(store);

    const dispatch = dispatchCellAction({ type: "set-value", value: "new" }, 0, 0);

    expect(dispatch.kind).toBe("command");
    if (dispatch.kind === "command") {
      manager.execute(dispatch.command);
    }
    expect(store.getCell(0, 0)).toBe("new");
  });

  it("undoing the resulting command restores the previous value", () => {
    const store = createStore();
    const manager = createTableCommandManager(store);
    const dispatch = dispatchCellAction({ type: "set-value", value: "new" }, 0, 0);
    if (dispatch.kind === "command") {
      manager.execute(dispatch.command);
    }

    manager.undo();

    expect(store.getCell(0, 0)).toBe("old");
  });

  it("translates a start-edit action without touching the store", () => {
    const store = createStore();

    const dispatch = dispatchCellAction({ type: "start-edit" }, 0, 0);

    expect(dispatch).toEqual({ kind: "start-edit" });
    expect(store.getCell(0, 0)).toBe("old");
  });
});
