import { describe, expect, it } from "vitest";
import { SetCellTypeCommand } from "../../../../src/core/commands";
import { TableStore } from "../../../../src/core/data";
import { createTableCommandManager } from "../../../../src/core/integration";
import type { CellRange } from "../../../../src/core/types";
import { createSheetFromRows } from "../../../helpers/createSheetModel";

const cell = (row: number, column: number): CellRange => ({
  startRow: row,
  startColumn: column,
  endRow: row,
  endColumn: column
});

function createStore(): TableStore {
  const rows = [{ a: "A0" }, { a: "A1" }, { a: "A2" }];
  return new TableStore(createSheetFromRows(rows, ["a"]));
}

describe("SetCellTypeCommand", () => {
  it("把类型名写进记录的 t 字段，不动值", () => {
    const store = createStore();
    const manager = createTableCommandManager(store);

    manager.execute(new SetCellTypeCommand(cell(0, 0), "date"));

    expect(store.sheet.getRecord(0, 0)?.t).toBe("date");
    expect(store.getCell(0, 0)).toBe("A0");
  });

  it("null 清掉格级类型，回到列默认", () => {
    const store = createStore();
    const manager = createTableCommandManager(store);

    manager.execute(new SetCellTypeCommand(cell(0, 0), "date"));
    manager.execute(new SetCellTypeCommand(cell(0, 0), null));

    expect(store.sheet.getRecord(0, 0)?.t).toBeUndefined();
    expect(store.getCell(0, 0)).toBe("A0");
  });

  it("撤销重做", () => {
    const store = createStore();
    const manager = createTableCommandManager(store);

    manager.execute(new SetCellTypeCommand(cell(0, 0), "date"));
    manager.undo();
    expect(store.sheet.getRecord(0, 0)?.t).toBeUndefined();

    manager.redo();
    expect(store.sheet.getRecord(0, 0)?.t).toBe("date");
  });

  it("一片区域是一次事务一次撤销", () => {
    const store = createStore();
    const manager = createTableCommandManager(store);

    manager.execute(new SetCellTypeCommand({ startRow: 0, startColumn: 0, endRow: 2, endColumn: 0 }, "date"));
    expect(store.sheet.getRecord(2, 0)?.t).toBe("date");

    manager.undo();
    for (let row = 0; row <= 2; row += 1) {
      expect(store.sheet.getRecord(row, 0)?.t).toBeUndefined();
    }
  });

  it("类型没变的格子不产生写入", () => {
    const store = createStore();
    const manager = createTableCommandManager(store);

    manager.execute(new SetCellTypeCommand(cell(0, 0), "date"));
    const before = store.version;
    manager.execute(new SetCellTypeCommand(cell(0, 0), "date"));

    expect(store.version).toBe(before);
  });
});
