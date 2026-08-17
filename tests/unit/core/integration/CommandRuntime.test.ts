import { describe, expect, it } from "vitest";

import {
  CellTypeRegistry,
  ColumnManager,
  registerBuiltinCellTypes,
  type ColumnDefinition
} from "../../../../src/core/cells";
import {
  InsertRowsCommand,
  SetRangeCommand,
  type BeforeCommandContext,
  type BeforeTransactionCommitContext
} from "../../../../src/core/commands";
import { ArrayDataSource, TableStore } from "../../../../src/core/data";
import { HeaderManager } from "../../../../src/core/headers";
import { HookManager, PRODUCTION_HOOK_REGISTRY } from "../../../../src/core/hooks";
import { GridStructureOwner, createGridCommandManager } from "../../../../src/core/integration";
import { SelectionModel } from "../../../../src/core/interaction";
import { LayoutEngine } from "../../../../src/core/layout";

function createFixture() {
  let columns: readonly ColumnDefinition[] = Object.freeze([{ id: "a" }, { id: "b" }]);
  const source = new ArrayDataSource(
    [{ a: "A", b: 1 }, { a: "B", b: 2 }],
    [{ id: "a", field: "a" }, { id: "b", field: "b" }]
  );
  const store = new TableStore(source);
  const registry = new CellTypeRegistry();
  registerBuiltinCellTypes(registry);
  const columnManager = new ColumnManager({ columns, registry });
  const headerManager = new HeaderManager({ columns });
  const layout = new LayoutEngine({
    rowCount: 2,
    columnCount: 2,
    rowHeight: 20,
    columnWidths: [80, 90]
  });
  const selection = new SelectionModel({ rowCount: 2, columnCount: 2 });
  const structureOwner = new GridStructureOwner({
    getColumns: () => columns,
    replaceColumns: (next) => {
      columns = next;
    },
    columnManager,
    headerManager,
    layout,
    selection
  });
  return { store, layout, selection, structureOwner };
}

describe("grid command runtime", () => {
  it("keeps commands as frozen data snapshots", () => {
    const range = { startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 };
    const values = [["A", 1]];
    const command = new SetRangeCommand(range, values);

    range.endColumn = 0;
    values[0]![0] = "changed";

    expect(Object.isFrozen(command)).toBe(true);
    expect(Object.isFrozen(command.params)).toBe(true);
    expect(Object.isFrozen(command.params.range)).toBe(true);
    expect(Object.isFrozen(command.params.values)).toBe(true);
    expect(command.params).toEqual({
      range: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 },
      values: [["A", 1]]
    });
  });

  it("plans one structural intent into one atomic data+structure transaction", () => {
    const fixture = createFixture();
    const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
    const commandTypes: string[] = [];
    const transactionCommandCounts: number[] = [];
    hooks.tap("beforeCommand", { name: "observe-command" }, (context: BeforeCommandContext) => {
      commandTypes.push(context.command.type);
    });
    hooks.tap(
      "beforeTransactionCommit",
      { name: "observe-transaction" },
      (context: BeforeTransactionCommitContext) => {
        transactionCommandCounts.push(context.commands.length);
      }
    );
    const manager = createGridCommandManager({ ...fixture, hooks });

    manager.execute(new InsertRowsCommand(1, [[undefined, undefined]]));

    expect(commandTypes).toEqual(["insert-rows"]);
    expect(transactionCommandCounts).toEqual([1]);
    expect(fixture.store.getRowCount()).toBe(3);
    expect(fixture.layout.rowCount).toBe(3);
    expect(fixture.selection.bounds.rowCount).toBe(3);
    expect(fixture.store.version).toBe(1);
    expect(manager.historySize).toBe(1);

    manager.undo();

    expect(fixture.store.getRowCount()).toBe(2);
    expect(fixture.layout.rowCount).toBe(2);
    expect(fixture.selection.bounds.rowCount).toBe(2);
    expect(fixture.store.version).toBe(2);
  });
});
