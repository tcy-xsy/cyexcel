import { describe, expect, it } from "vitest";

import { HideColumnCommand, ShowColumnCommand } from "../../../../src/core/commands";
import { ArrayDataSource, TableStore } from "../../../../src/core/data";
import { createTableCommandManager } from "../../../../src/core/integration";
import { LayoutEngine } from "../../../../src/core/layout";


function createLayout(): LayoutEngine {
  return new LayoutEngine({ rowCount: 5, columnCount: 3, rowHeight: 20, columnWidths: [50, 60, 70] });
}

function createManager(layout: LayoutEngine) {
  const source = new ArrayDataSource([], [
    { id: "a", field: "a" },
    { id: "b", field: "b" },
    { id: "c", field: "c" }
  ]);
  return createTableCommandManager(new TableStore(source), { layout });
}

describe("HideColumnCommand", () => {
  it("hides the column and produces a column-layout patch", () => {
    const layout = createLayout();
    const manager = createManager(layout);
    const command = new HideColumnCommand(1);

    const patches = manager.execute(command);

    expect(patches).toEqual([{ type: "column-layout", column: 1 }]);
    expect(layout.getColumnWidth(1)).toBe(0);
    expect(layout.isColumnHidden(1)).toBe(true);
  });

  it("undo restores the column's exact original width", () => {
    const layout = createLayout();
    const manager = createManager(layout);
    const command = new HideColumnCommand(1);
    manager.execute(command);

    manager.undo();

    expect(layout.getColumnWidth(1)).toBe(60);
    expect(layout.isColumnHidden(1)).toBe(false);
  });

  it("redo (execute again) hides it again", () => {
    const layout = createLayout();
    const manager = createManager(layout);
    const command = new HideColumnCommand(1);
    manager.execute(command);
    manager.undo();

    manager.redo();

    expect(layout.getColumnWidth(1)).toBe(0);
  });

  it("undoing a hide no-op keeps an already hidden column hidden", () => {
    const layout = createLayout();
    const manager = createManager(layout);
    layout.hideColumn(1);

    manager.execute(new HideColumnCommand(1));
    manager.undo();

    expect(layout.getColumnWidth(1)).toBe(0);
    expect(layout.isColumnHidden(1)).toBe(true);
  });
});

describe("ShowColumnCommand", () => {
  function createHiddenLayout(): LayoutEngine {
    const layout = createLayout();
    layout.hideColumn(1);
    return layout;
  }

  it("shows a hidden column and produces a column-layout patch", () => {
    const layout = createHiddenLayout();
    const manager = createManager(layout);
    const command = new ShowColumnCommand(1);

    const patches = manager.execute(command);

    expect(patches).toEqual([{ type: "column-layout", column: 1 }]);
    expect(layout.getColumnWidth(1)).toBe(60);
    expect(layout.isColumnHidden(1)).toBe(false);
  });

  it("undo re-hides the column", () => {
    const layout = createHiddenLayout();
    const manager = createManager(layout);
    const command = new ShowColumnCommand(1);
    manager.execute(command);

    manager.undo();

    expect(layout.getColumnWidth(1)).toBe(0);
    expect(layout.isColumnHidden(1)).toBe(true);
  });

  it("redo (execute again) shows it again", () => {
    const layout = createHiddenLayout();
    const manager = createManager(layout);
    const command = new ShowColumnCommand(1);
    manager.execute(command);
    manager.undo();

    manager.redo();

    expect(layout.getColumnWidth(1)).toBe(60);
  });

  it("undoing a show no-op keeps an already visible column visible", () => {
    const layout = createLayout();
    const manager = createManager(layout);

    manager.execute(new ShowColumnCommand(1));
    manager.undo();

    expect(layout.getColumnWidth(1)).toBe(60);
    expect(layout.isColumnHidden(1)).toBe(false);
  });
});
