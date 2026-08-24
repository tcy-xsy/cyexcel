import { describe, expect, it } from "vitest";

import {
  CellTypeRegistry,
  ColumnManager,
  registerBuiltinCellTypes,
  type ColumnDefinition
} from "../../../../src/core/cells";
import {
  InsertRowsCommand,
  ReorderRowsCommand,
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

/** Three rows with three distinct heights, so no assertion can pass by accident. */
function createReorderFixture() {
  let columns: readonly ColumnDefinition[] = Object.freeze([{ id: "a" }]);
  const source = new ArrayDataSource(
    [{ a: "A" }, { a: "B" }, { a: "C" }],
    [{ id: "a", field: "a" }]
  );
  const store = new TableStore(source);
  const registry = new CellTypeRegistry();
  registerBuiltinCellTypes(registry);
  const columnManager = new ColumnManager({ columns, registry });
  const headerManager = new HeaderManager({ columns });
  const layout = new LayoutEngine({
    rowCount: 3,
    columnCount: 1,
    rowHeight: 20,
    columnWidths: [80]
  });
  layout.setRowHeight(0, 60);
  layout.setRowHeight(2, 44);
  const selection = new SelectionModel({ rowCount: 3, columnCount: 1 });
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
  return {
    store,
    layout,
    selection,
    structureOwner,
    readColumn: () => [0, 1, 2].map((row) => store.getCell(row, 0)),
    readHeights: () => [0, 1, 2].map((row) => layout.getRowHeight(row))
  };
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

  describe("reorder-rows", () => {
    // [2, 0, 1] is asymmetric on purpose: a swap would be its own inverse and
    // could not tell a direction bug from a correct one.
    const order = () => Int32Array.from([2, 0, 1]);

    it("commits the row data and the row geometry in one transaction", () => {
      const fixture = createReorderFixture();
      const hooks = new HookManager({ registry: PRODUCTION_HOOK_REGISTRY });
      const stateAtCommitBarrier: { column: unknown[]; heights: number[] }[] = [];
      const transactionCommandCounts: number[] = [];
      hooks.tap(
        "beforeTransactionCommit",
        { name: "observe-transaction" },
        (context: BeforeTransactionCommitContext) => {
          transactionCommandCounts.push(context.commands.length);
          stateAtCommitBarrier.push({
            column: fixture.readColumn(),
            heights: fixture.readHeights()
          });
        }
      );
      const manager = createGridCommandManager({ ...fixture, hooks });

      expect(fixture.readColumn()).toEqual(["A", "B", "C"]);
      expect(fixture.readHeights()).toEqual([60, 20, 44]);

      manager.execute(new ReorderRowsCommand(order()));

      // One transaction planned both sides: the whole reorder was still
      // unpublished at the commit barrier, data and geometry alike.
      expect(transactionCommandCounts).toEqual([1]);
      expect(stateAtCommitBarrier).toEqual([
        { column: ["A", "B", "C"], heights: [60, 20, 44] }
      ]);
      // ...and both landed under a single model revision, so no frame can
      // observe sorted data against unsorted geometry.
      expect(fixture.store.version).toBe(1);
      expect(fixture.readColumn()).toEqual(["C", "A", "B"]);
      expect(fixture.readHeights()).toEqual([44, 60, 20]);
      expect(manager.historySize).toBe(1);
    });

    it("undoes the data and the geometry together in one step", () => {
      const fixture = createReorderFixture();
      const manager = createGridCommandManager({ ...fixture });

      manager.execute(new ReorderRowsCommand(order()));
      manager.undo();

      expect(fixture.readColumn()).toEqual(["A", "B", "C"]);
      expect(fixture.readHeights()).toEqual([60, 20, 44]);
      expect(fixture.store.version).toBe(2);
      expect(manager.historySize).toBe(0);
    });

    it("emits a single row-reorder patch carrying the order", () => {
      const fixture = createReorderFixture();
      const manager = createGridCommandManager({ ...fixture });

      const patches = manager.execute(new ReorderRowsCommand(order()));

      expect(patches).toEqual([{ type: "row-reorder", order: Int32Array.from([2, 0, 1]) }]);
    });
  });
});
