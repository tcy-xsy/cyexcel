import { describe, expect, it } from "vitest";

import {
  HideColumnCommand,
  ResizeColumnCommand,
  ResizeRowCommand,
  ShowColumnCommand
} from "../../../../src/core/commands";
import { SetRowHeightsCommand } from "../../../../src/core/commands/SetRowHeightsCommand";
import { TableStore } from "../../../../src/core/data";
import { createTableCommandManager } from "../../../../src/core/integration";
import {
  columnWidthOf,
  createSeededLayout,
  isColumnHidden
} from "../../../helpers/columnWidthAuthority";
import { createSheetFromRows } from "../../../helpers/createSheetModel";
import { rowHeightProvenanceOf } from "../../../helpers/rowHeightAuthority";

function createFixture() {
  const sheet = createSheetFromRows(
    [
      { a: "A1", b: "B1", c: "C1" },
      { a: "A2", b: "B2", c: "C2" },
      { a: "A3", b: "B3", c: "C3" }
    ],
    ["a", "b", "c"]
  );
  const store = new TableStore(sheet);
  // 走 `createSeededLayout` 而不是直接 `new LayoutEngine`：列宽的权威在 `columnData`，
  // 只喂给布局的话这个夹具一出生就违反 §4.2 的列轴不变量（撤销会退到缺省宽 120 而不是
  // 夹具写的那几个数）。
  const layout = createSeededLayout(sheet, {
    rowCount: 3,
    columnCount: 3,
    rowHeight: 20,
    columnWidths: [50, 60, 70]
  });
  const manager = createTableCommandManager(store, { layout });
  return { layout, manager, store, sheet };
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
    const { layout, manager, store } = createFixture();

    manager.execute(new ResizeRowCommand(1, 36));
    const manualRoot = layout.root;
    const patches = manager.execute(
      new SetRowHeightsCommand([{ row: 1, height: 72 }], "auto")
    );

    expect(patches).toEqual([]);
    expect(layout.root).toBe(manualRoot);
    expect(layout.getRowHeight(1)).toBe(36);
    expect(rowHeightProvenanceOf(store.sheet, 1)).toBe("manual");
  });

  it("restores auto provenance when a manual resize is undone", () => {
    const { layout, manager, store } = createFixture();

    manager.execute(new SetRowHeightsCommand([{ row: 1, height: 36 }], "auto"));
    expect(rowHeightProvenanceOf(store.sheet, 1)).toBe("auto");

    manager.execute(new ResizeRowCommand(1, 52));
    expect(rowHeightProvenanceOf(store.sheet, 1)).toBe("manual");

    manager.undo();
    expect(layout.getRowHeight(1)).toBe(36);
    expect(rowHeightProvenanceOf(store.sheet, 1)).toBe("auto");

    manager.redo();
    expect(layout.getRowHeight(1)).toBe(52);
    expect(rowHeightProvenanceOf(store.sheet, 1)).toBe("manual");
  });

  it("publishes a manual lock even when the numeric height is unchanged", () => {
    const { layout, manager, store } = createFixture();

    manager.execute(new SetRowHeightsCommand([{ row: 0, height: 28 }], "auto"));
    const patches = manager.execute(new ResizeRowCommand(0, 28));

    expect(patches).toHaveLength(1);
    expect(layout.getRowHeight(0)).toBe(28);
    expect(rowHeightProvenanceOf(store.sheet, 0)).toBe("manual");
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
    const { layout, manager, store, sheet } = createFixture();
    // 隐藏走命令：它要同时写权威（`columnData[1].hd`）和索引（几何宽度 0），
    // 两半必须在同一个事务里（设计稿 §4.3）。因此版本号从这一条之后开始数。
    manager.execute(new HideColumnCommand(1));
    const layoutVersion = layout.version;
    const storeVersion = store.version;

    expect(layout.getColumnWidth(1)).toBe(0);
    expect(columnWidthOf(sheet, 1, layout.defaultColumnWidth)).toBe(60);
    expect(isColumnHidden(sheet, 1)).toBe(true);

    // 三次操作里**只有权威那半边在动**：几何宽度全程是 0，所以布局版本一次都不该走，
    // 补丁也一条都不该发——屏幕上一个像素没变。会走的是模型版本，因为 `columnData[1].w`
    // 确实改了，取消隐藏时就是按它逐像素回去的。
    expect(manager.execute(new ResizeColumnCommand(1, 95))).toEqual([]);
    expect(layout.getColumnWidth(1)).toBe(0);
    expect(columnWidthOf(sheet, 1, layout.defaultColumnWidth)).toBe(95);
    expect(isColumnHidden(sheet, 1)).toBe(true);
    expect(layout.version).toBe(layoutVersion);
    expect(store.version).toBe(storeVersion + 1);
    expect(manager.lastChangeSet).toEqual({ modelVersion: storeVersion + 1, changes: [] });

    manager.undo();
    expect(layout.getColumnWidth(1)).toBe(0);
    expect(columnWidthOf(sheet, 1, layout.defaultColumnWidth)).toBe(60);
    expect(isColumnHidden(sheet, 1)).toBe(true);
    expect(layout.version).toBe(layoutVersion);
    expect(store.version).toBe(storeVersion + 2);
    expect(manager.lastChangeSet?.modelVersion).toBe(storeVersion + 2);

    manager.redo();
    expect(layout.getColumnWidth(1)).toBe(0);
    expect(columnWidthOf(sheet, 1, layout.defaultColumnWidth)).toBe(95);
    expect(isColumnHidden(sheet, 1)).toBe(true);
    expect(layout.version).toBe(layoutVersion);
    expect(store.version).toBe(storeVersion + 3);
    expect(manager.lastChangeSet?.modelVersion).toBe(storeVersion + 3);

    // 而这一条要看得见：取消隐藏时回到的正是上面那个 95，不是夹具原来的 60。
    manager.execute(new ShowColumnCommand(1));
    expect(layout.getColumnWidth(1)).toBe(95);
    expect(layout.version).toBe(layoutVersion + 1);
  });
});
