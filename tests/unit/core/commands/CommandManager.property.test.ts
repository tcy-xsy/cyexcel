import fc from "fast-check";
import { describe, expect, it } from "vitest";

import { SetCellCommand } from "../../../../src/core/commands";
import { TableStore } from "../../../../src/core/data";
import { createTableCommandManager } from "../../../../src/core/integration";
import { createSheetFromRows } from "../../../helpers/createSheetModel";

const ROW_COUNT = 5;
const COLUMNS = ["a", "b", "c"];

function createGrid() {
  const rows = Array.from({ length: ROW_COUNT }, (_unused, row) => ({
    a: row * 10,
    b: row * 10 + 1,
    c: row * 10 + 2
  }));
  return new TableStore(createSheetFromRows(rows, COLUMNS));
}

function snapshot(store: TableStore): unknown[][] {
  const grid: unknown[][] = [];
  for (let row = 0; row < ROW_COUNT; row += 1) {
    const rowValues: unknown[] = [];
    for (let column = 0; column < COLUMNS.length; column += 1) {
      rowValues.push(store.getCell(row, column));
    }
    grid.push(rowValues);
  }
  return grid;
}

describe("CommandManager — random command sequences", () => {
  it("undoing every executed command restores the initial grid", () => {
    fc.assert(
      fc.property(
        fc.array(
          fc.record({
            row: fc.integer({ min: 0, max: ROW_COUNT - 1 }),
            column: fc.integer({ min: 0, max: COLUMNS.length - 1 }),
            value: fc.integer()
          }),
          { maxLength: 30 }
        ),
        operations => {
          const store = createGrid();
          // historyLimit is sized to the batch so eviction never kicks in —
          // this property is about undo correctness, not budget eviction
          // (that has its own tests in CommandManager.test.ts).
          const manager = createTableCommandManager(store, {
            historyLimit: operations.length + 1
          });
          const initial = snapshot(store);

          for (const operation of operations) {
            manager.execute(
              new SetCellCommand(operation.row, operation.column, operation.value)
            );
          }

          while (manager.canUndo()) {
            manager.undo();
          }

          expect(snapshot(store)).toEqual(initial);
        }
      )
    );
  });
});
