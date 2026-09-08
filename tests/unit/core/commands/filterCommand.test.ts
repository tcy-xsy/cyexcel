import { describe, expect, it } from "vitest";

import { FilterCommand, ReorderRowsCommand } from "../../../../src/core/commands";
import { TableStore } from "../../../../src/core/data";
import { createTableCommandManager } from "../../../../src/core/integration";
import { createSeededLayout } from "../../../helpers/columnWidthAuthority";
import { createSheetFromRows } from "../../../helpers/createSheetModel";

function createFixture() {
  const sheet = createSheetFromRows(
    [{ a: "A1" }, { a: "A2" }, { a: "A3" }, { a: "A4" }],
    ["a"]
  );
  const store = new TableStore(sheet);
  const layout = createSeededLayout(sheet, {
    rowCount: 4,
    columnCount: 1,
    rowHeight: 20,
    columnWidths: [50]
  });
  const manager = createTableCommandManager(store, { layout });
  return { layout, manager, store, sheet };
}

describe("FilterCommand", () => {
  it("freezes a copy of the given set, independent of later mutation of the source", () => {
    const source = new Set([1, 2]);
    const command = new FilterCommand(source);
    source.add(3);

    expect([...command.params.filteredRows].sort()).toEqual([1, 2]);
  });

  it("stages exactly one layout mutation and no data mutation", () => {
    const { layout, manager, store } = createFixture();
    const before = [0, 1, 2, 3].map((row) => store.getCell(row, 0));

    const patches = manager.execute(new FilterCommand(new Set([1, 3])), { history: "skip" });

    // The commit still bumps the shared model version (same as ResizeRowCommand,
    // which also stages only a layout mutation) — what matters is that no
    // *data* value changed.
    expect(store.version).toBe(1);
    expect(layout.version).toBe(1);
    expect([0, 1, 2, 3].map((row) => store.getCell(row, 0))).toEqual(before);
    expect(patches).toEqual([{ type: "row-visibility" }]);
    expect(layout.isRowVisible(1)).toBe(false);
    expect(layout.isRowVisible(3)).toBe(false);
    expect(layout.isRowVisible(0)).toBe(true);
    expect(layout.isRowVisible(2)).toBe(true);
  });

  it("run with history: skip never enters the undo stack", () => {
    const { manager } = createFixture();

    manager.execute(new FilterCommand(new Set([1])), { history: "skip" });

    expect(manager.canUndo()).toBe(false);
  });

  it("run with history: skip does not clear an existing redo entry", () => {
    const { layout, manager } = createFixture();

    // Build a redo entry: reorder, then undo it.
    manager.execute(new ReorderRowsCommand(Int32Array.from([1, 0, 2, 3])));
    manager.undo();
    expect(manager.canRedo()).toBe(true);

    // Filtering must not touch that redo entry.
    manager.execute(new FilterCommand(new Set([2])), { history: "skip" });
    expect(manager.canRedo()).toBe(true);

    manager.redo();
    // The reorder replayed correctly even after the intervening filter.
    expect(layout.isRowVisible(2)).toBe(false);
  });

  it("clearing the filter (empty set) is also a skip-history commit", () => {
    const { layout, manager } = createFixture();
    manager.execute(new FilterCommand(new Set([1])), { history: "skip" });

    manager.execute(new FilterCommand(new Set()), { history: "skip" });

    expect(layout.isRowVisible(1)).toBe(true);
    expect(manager.canUndo()).toBe(false);
  });
});
