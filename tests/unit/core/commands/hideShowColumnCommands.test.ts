import { describe, expect, it } from "vitest";

import { HideColumnCommand, ShowColumnCommand } from "../../../../src/core/commands";
import { TableStore } from "../../../../src/core/data";
import { createTableCommandManager } from "../../../../src/core/integration";
import type { LayoutEngine } from "../../../../src/core/layout";
import type { SheetModel } from "../../../../src/core/sheet/SheetModel";
import {
  columnWidthOf,
  createSeededLayout,
  isColumnHidden
} from "../../../helpers/columnWidthAuthority";
import { createSheetFromRows } from "../../../helpers/createSheetModel";

/**
 * 隐藏/取消隐藏在 T8b 之后是**两半**：权威 `columnData[c].hd`（外加记着「恢复成多宽」
 * 的 `w`）和索引里的几何宽度（隐藏即 0）。两半都要断言——只看几何宽度的话，一条把
 * 索引写对、把权威写漏的实现能全绿通过（设计稿 §4.4）。
 */
function createFixture(): { layout: LayoutEngine; sheet: SheetModel; manager: ReturnType<typeof createTableCommandManager> } {
  const sheet = createSheetFromRows([], ["a", "b", "c"]);
  const layout = createSeededLayout(sheet, {
    rowCount: 5,
    columnCount: 3,
    rowHeight: 20,
    columnWidths: [50, 60, 70]
  });
  return { layout, sheet, manager: createTableCommandManager(new TableStore(sheet), { layout }) };
}

describe("HideColumnCommand", () => {
  it("hides the column and produces a column-layout patch", () => {
    const { layout, sheet, manager } = createFixture();

    const patches = manager.execute(new HideColumnCommand(1));

    expect(patches).toEqual([{ type: "column-layout", column: 1 }]);
    expect(layout.getColumnWidth(1)).toBe(0);
    expect(isColumnHidden(sheet, 1)).toBe(true);
    // 恢复宽度留在权威里，这正是取消隐藏时逐像素回到 60 的依据。
    expect(columnWidthOf(sheet, 1, layout.defaultColumnWidth)).toBe(60);
  });

  it("undo restores the column's exact original width", () => {
    const { layout, sheet, manager } = createFixture();
    manager.execute(new HideColumnCommand(1));

    manager.undo();

    expect(layout.getColumnWidth(1)).toBe(60);
    expect(isColumnHidden(sheet, 1)).toBe(false);
  });

  it("redo (execute again) hides it again", () => {
    const { layout, manager } = createFixture();
    manager.execute(new HideColumnCommand(1));
    manager.undo();

    manager.redo();

    expect(layout.getColumnWidth(1)).toBe(0);
  });

  it("hiding an already hidden column does nothing and takes no undo slot", () => {
    const { layout, sheet, manager } = createFixture();
    manager.execute(new HideColumnCommand(1));
    const historySize = manager.historySize;

    // 两半都判得出「没变」：权威那边整条 `ColumnMeta` 不动（`hd` 已经是 `true`），
    // 索引那边几何宽度已经是 0。于是这一条没有效果，不进历史。
    expect(manager.execute(new HideColumnCommand(1))).toEqual([]);
    expect(manager.historySize).toBe(historySize);
    // 重复隐藏尤其不该把 0 记成恢复宽度——它写的是 `hd`，`w` 一个字节没碰。
    expect(columnWidthOf(sheet, 1, layout.defaultColumnWidth)).toBe(60);

    // 既然它没进历史，撤销撤的就是**前一条**——那条真的隐藏了这一列的。
    manager.undo();

    expect(layout.getColumnWidth(1)).toBe(60);
    expect(isColumnHidden(sheet, 1)).toBe(false);
  });
});

describe("ShowColumnCommand", () => {
  function createHiddenFixture(): ReturnType<typeof createFixture> {
    const fixture = createFixture();
    fixture.manager.execute(new HideColumnCommand(1));
    fixture.manager.clearHistory();
    return fixture;
  }

  it("shows a hidden column and produces a column-layout patch", () => {
    const { layout, sheet, manager } = createHiddenFixture();

    const patches = manager.execute(new ShowColumnCommand(1));

    expect(patches).toEqual([{ type: "column-layout", column: 1 }]);
    expect(layout.getColumnWidth(1)).toBe(60);
    expect(isColumnHidden(sheet, 1)).toBe(false);
  });

  it("undo re-hides the column", () => {
    const { layout, sheet, manager } = createHiddenFixture();
    manager.execute(new ShowColumnCommand(1));

    manager.undo();

    expect(layout.getColumnWidth(1)).toBe(0);
    expect(isColumnHidden(sheet, 1)).toBe(true);
  });

  it("redo (execute again) shows it again", () => {
    const { layout, manager } = createHiddenFixture();
    manager.execute(new ShowColumnCommand(1));
    manager.undo();

    manager.redo();

    expect(layout.getColumnWidth(1)).toBe(60);
  });

  it("showing an already visible column does nothing and takes no undo slot", () => {
    const { layout, sheet, manager } = createFixture();

    expect(manager.execute(new ShowColumnCommand(1))).toEqual([]);
    expect(manager.historySize).toBe(0);

    manager.undo();

    expect(layout.getColumnWidth(1)).toBe(60);
    expect(isColumnHidden(sheet, 1)).toBe(false);
  });
});
