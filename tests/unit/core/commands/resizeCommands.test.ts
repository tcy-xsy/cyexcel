import { describe, expect, it } from "vitest";

import {
  ResizeColumnCommand,
  ResizeRowCommand
} from "../../../../src/core/commands";
import { SetRowHeightsCommand } from "../../../../src/core/commands/SetRowHeightsCommand";
import { TableStore } from "../../../../src/core/data";
import { createTableCommandManager } from "../../../../src/core/integration";
import { LayoutEngine } from "../../../../src/core/layout";
import { createSheetFromRows } from "../../../helpers/createSheetModel";

function createFixture() {
  const store = new TableStore(
    createSheetFromRows(
      [
        { a: "A1", b: "B1", c: "C1" },
        { a: "A2", b: "B2", c: "C2" },
        { a: "A3", b: "B3", c: "C3" }
      ],
      ["a", "b", "c"]
    )
  );
  const layout = new LayoutEngine({
    rowCount: 3,
    columnCount: 3,
    rowHeight: 20,
    columnWidths: [50, 60, 70]
  });
  const manager = createTableCommandManager(store, { layout });
  return { layout, manager, store };
}

describe("ResizeRowCommand", () => {
  it("keeps layout and model versions aligned through execute, undo, and redo", () => {
    const { layout, manager, store } = createFixture();
    const patch = {
      type: "row-layout" as const,
      rows: [1],
      range: { startRow: 1, startColumn: 0, endRow: 2, endColumn: 2 }
    };

    expect(layout.version).toBe(0);
    expect(store.version).toBe(0);

    expect(manager.execute(new ResizeRowCommand(1, 36))).toEqual([patch]);
    expect(layout.getRowHeight(1)).toBe(36);
    expect(layout.version).toBe(1);
    expect(store.version).toBe(1);
    expect(manager.lastChangeSet).toEqual({ modelVersion: 1, changes: [patch] });

    expect(manager.undo()).toEqual([patch]);
    expect(layout.getRowHeight(1)).toBe(20);
    expect(layout.version).toBe(2);
    expect(store.version).toBe(2);
    expect(manager.lastChangeSet).toEqual({ modelVersion: 2, changes: [patch] });

    expect(manager.redo()).toEqual([patch]);
    expect(layout.getRowHeight(1)).toBe(36);
    expect(layout.version).toBe(3);
    expect(store.version).toBe(3);
    expect(manager.lastChangeSet).toEqual({ modelVersion: 3, changes: [patch] });
  });

  it("commits a row-height batch as one revision and one undo step", () => {
    const { layout, manager, store } = createFixture();
    const patches = [{
      type: "row-layout" as const,
      rows: [0, 2],
      range: { startRow: 0, startColumn: 0, endRow: 2, endColumn: 2 }
    }];

    expect(manager.execute(new SetRowHeightsCommand([
      { row: 0, height: 28 },
      { row: 2, height: 44 }
    ], "auto"))).toEqual(patches);
    expect(layout.getRowHeight(0)).toBe(28);
    expect(layout.getRowHeight(2)).toBe(44);
    expect(layout.version).toBe(1);
    expect(store.version).toBe(1);
    expect(manager.historySize).toBe(1);

    expect(manager.undo()).toEqual(patches);
    expect(layout.getRowHeight(0)).toBe(20);
    expect(layout.getRowHeight(2)).toBe(20);
    expect(layout.version).toBe(2);
    expect(store.version).toBe(2);
    expect(manager.historySize).toBe(0);

    expect(manager.redo()).toEqual(patches);
    expect(layout.getRowHeight(0)).toBe(28);
    expect(layout.getRowHeight(2)).toBe(44);
    expect(layout.version).toBe(3);
    expect(store.version).toBe(3);
    expect(manager.historySize).toBe(1);
  });

  it("does not let an auto-height batch overwrite a manually sized row", () => {
    const { layout, manager } = createFixture();

    manager.execute(new ResizeRowCommand(1, 36));
    const manualRoot = layout.root;
    const patches = manager.execute(
      new SetRowHeightsCommand([{ row: 1, height: 72 }], "auto")
    );

    expect(patches).toEqual([]);
    expect(layout.root).toBe(manualRoot);
    expect(layout.getRowHeight(1)).toBe(36);
    expect(layout.getRowHeightProvenance(1)).toBe("manual");
  });

  it("restores auto provenance when a manual resize is undone", () => {
    const { layout, manager } = createFixture();

    manager.execute(new SetRowHeightsCommand([{ row: 1, height: 36 }], "auto"));
    expect(layout.getRowHeightProvenance(1)).toBe("auto");

    manager.execute(new ResizeRowCommand(1, 52));
    expect(layout.getRowHeightProvenance(1)).toBe("manual");

    manager.undo();
    expect(layout.getRowHeight(1)).toBe(36);
    expect(layout.getRowHeightProvenance(1)).toBe("auto");

    manager.redo();
    expect(layout.getRowHeight(1)).toBe(52);
    expect(layout.getRowHeightProvenance(1)).toBe("manual");
  });

  it("publishes a manual lock even when the numeric height is unchanged", () => {
    const { layout, manager } = createFixture();

    manager.execute(new SetRowHeightsCommand([{ row: 0, height: 28 }], "auto"));
    const patches = manager.execute(new ResizeRowCommand(0, 28));

    expect(patches).toHaveLength(1);
    expect(layout.getRowHeight(0)).toBe(28);
    expect(layout.getRowHeightProvenance(0)).toBe("manual");
  });
});

describe("ResizeColumnCommand", () => {
  it("keeps layout and model versions aligned through execute, undo, and redo", () => {
    const { layout, manager, store } = createFixture();
    const patch = { type: "column-layout" as const, column: 1 };

    expect(manager.execute(new ResizeColumnCommand(1, 95))).toEqual([patch]);
    expect(layout.getColumnWidth(1)).toBe(95);
    expect(layout.version).toBe(1);
    expect(store.version).toBe(1);
    expect(manager.lastChangeSet).toEqual({ modelVersion: 1, changes: [patch] });

    expect(manager.undo()).toEqual([patch]);
    expect(layout.getColumnWidth(1)).toBe(60);
    expect(layout.version).toBe(2);
    expect(store.version).toBe(2);
    expect(manager.lastChangeSet).toEqual({ modelVersion: 2, changes: [patch] });

    expect(manager.redo()).toEqual([patch]);
    expect(layout.getColumnWidth(1)).toBe(95);
    expect(layout.version).toBe(3);
    expect(store.version).toBe(3);
    expect(manager.lastChangeSet).toEqual({ modelVersion: 3, changes: [patch] });
  });

  it("resizes a hidden column's restore width without showing the column", () => {
    const { layout, manager, store } = createFixture();
    layout.hideColumn(1);

    expect(layout.getColumnWidth(1)).toBe(0);
    expect(layout.getColumnRestoreWidth(1)).toBe(60);
    expect(layout.isColumnHidden(1)).toBe(true);
    expect(layout.version).toBe(1);
    expect(store.version).toBe(0);

    manager.execute(new ResizeColumnCommand(1, 95));
    expect(layout.getColumnWidth(1)).toBe(0);
    expect(layout.getColumnRestoreWidth(1)).toBe(95);
    expect(layout.isColumnHidden(1)).toBe(true);
    expect(layout.version).toBe(2);
    expect(store.version).toBe(1);
    expect(manager.lastChangeSet?.modelVersion).toBe(1);

    manager.undo();
    expect(layout.getColumnWidth(1)).toBe(0);
    expect(layout.getColumnRestoreWidth(1)).toBe(60);
    expect(layout.isColumnHidden(1)).toBe(true);
    expect(layout.version).toBe(3);
    expect(store.version).toBe(2);
    expect(manager.lastChangeSet?.modelVersion).toBe(2);

    manager.redo();
    expect(layout.getColumnWidth(1)).toBe(0);
    expect(layout.getColumnRestoreWidth(1)).toBe(95);
    expect(layout.isColumnHidden(1)).toBe(true);
    expect(layout.version).toBe(4);
    expect(store.version).toBe(3);
    expect(manager.lastChangeSet?.modelVersion).toBe(3);
  });
});
