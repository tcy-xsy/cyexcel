import { describe, expect, it } from "vitest";

import { SetCellCommand, SetStyleCommand } from "../../../../src/core/commands";
import { TableStore } from "../../../../src/core/data";
import { CyExcelError } from "../../../../src/core/errors";
import { createTableCommandManager } from "../../../../src/core/integration";
import { StyleRegistry } from "../../../../src/core/style/StyleRegistry";
import type { CellRange } from "../../../../src/core/types";
import { createSheetFromRows } from "../../../helpers/createSheetModel";
import { MutationCoordinator } from "../../../../src/core/mutations";

function range(startRow: number, startColumn: number, endRow: number, endColumn: number): CellRange {
  return { startRow, startColumn, endRow, endColumn };
}

const cell = (row: number, column: number): CellRange => range(row, column, row, column);

/**
 * 3 行 x 3 列，但**第三列整列没有值**——坐标空间里存在、却未占用的格子，
 * "给空格子刷底色"那几条要用它（`rowCount === rows.length`，所以没有空行可用）。
 */
function createStore(): { store: TableStore; styles: StyleRegistry } {
  const fields = ["c0", "c1", "c2"];
  const rows = Array.from({ length: 3 }, (_unused, r) =>
    Object.fromEntries(["c0", "c1"].map((field, c) => [field, `R${r}C${c}`]))
  );
  const styles = new StyleRegistry();
  const store = new TableStore(createSheetFromRows(rows, fields), new MutationCoordinator(), styles);
  return { store, styles };
}

/** 一格当前解析出来的样式对象（走 id 查表），没有样式时是 undefined。 */
function styleAt(store: TableStore, styles: StyleRegistry, row: number, column: number) {
  const id = store.sheet.getRecord(row, column)?.s;
  return id === undefined ? undefined : styles.get(id);
}

describe("SetStyleCommand", () => {
  it("写单格样式，产出 cell-style 补丁而不是 cell-value", () => {
    const { store, styles } = createStore();
    const manager = createTableCommandManager(store);

    const patches = manager.execute(new SetStyleCommand(cell(0, 0), { bold: true }));

    expect(patches).toContainEqual({ type: "cell-style", row: 0, column: 0 });
    expect(patches).not.toContainEqual({ type: "cell-value", row: 0, column: 0 });
    expect(styleAt(store, styles, 0, 0)).toEqual({ bold: true });
  });

  it("区域样式产出 range-style 补丁", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);

    const patches = manager.execute(new SetStyleCommand(range(0, 0, 1, 1), { bold: true }));

    expect(patches).toContainEqual({ type: "range-style", range: range(0, 0, 1, 1) });
  });

  it("不碰这一格的值——样式和值是同一条记录里的两个字段", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);

    manager.execute(new SetStyleCommand(cell(0, 0), { bold: true }));

    expect(store.getCell(0, 0)).toBe("R0C0");
  });

  it("合并语义：再设一个键不会把先前的键冲掉", () => {
    const { store, styles } = createStore();
    const manager = createTableCommandManager(store);

    manager.execute(new SetStyleCommand(cell(0, 0), { bold: true }));
    manager.execute(new SetStyleCommand(cell(0, 0), { color: "#f00" }));

    expect(styleAt(store, styles, 0, 0)).toEqual({ bold: true, color: "#f00" });
  });

  it("null 清一个键，其余保留", () => {
    const { store, styles } = createStore();
    const manager = createTableCommandManager(store);

    manager.execute(new SetStyleCommand(cell(0, 0), { bold: true, color: "#f00" }));
    manager.execute(new SetStyleCommand(cell(0, 0), { bold: null }));

    expect(styleAt(store, styles, 0, 0)).toEqual({ color: "#f00" });
  });

  it("清掉最后一个键后这一格不再有 s 字段，值还在", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);

    manager.execute(new SetStyleCommand(cell(0, 0), { bold: true }));
    manager.execute(new SetStyleCommand(cell(0, 0), { bold: null }));

    expect(store.sheet.getRecord(0, 0)?.s).toBeUndefined();
    expect(store.getCell(0, 0)).toBe("R0C0");
  });

  it("给空格子设样式会造出一条只有样式的记录——Excel 里给空格刷底色就是这样", () => {
    const { store, styles } = createStore();
    const manager = createTableCommandManager(store);

    // 第三列在坐标空间里但没有值
    manager.execute(new SetStyleCommand(cell(2, 2), { background: "#eee" }));
    manager.execute(new SetStyleCommand(cell(0, 0), { background: "#eee" }));

    expect(store.sheet.getRecord(2, 2)?.v).toBeUndefined();
    // 两格共用同一条样式——intern 的意义就在这里
    expect(store.sheet.getRecord(2, 2)?.s).toBe(store.sheet.getRecord(0, 0)?.s);
    expect(styles.size).toBe(1);
  });

  it("空格子的样式被清光时整条记录消失，回到未占用", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);

    manager.execute(new SetStyleCommand(cell(2, 2), { background: "#eee" }));
    manager.execute(new SetStyleCommand(cell(2, 2), { background: null }));

    expect(store.sheet.getRecord(2, 2)).toBeUndefined();
  });

  it("撤销回到原来的样式，重做再回去", () => {
    const { store, styles } = createStore();
    const manager = createTableCommandManager(store);

    manager.execute(new SetStyleCommand(cell(0, 0), { bold: true }));
    manager.execute(new SetStyleCommand(cell(0, 0), { color: "#f00" }));

    manager.undo();
    expect(styleAt(store, styles, 0, 0)).toEqual({ bold: true });

    manager.undo();
    expect(styleAt(store, styles, 0, 0)).toBeUndefined();

    manager.redo();
    expect(styleAt(store, styles, 0, 0)).toEqual({ bold: true });
  });

  it("撤销一片区域是一步，不是每格一步——一次事务一次 History 记录", () => {
    const { store, styles } = createStore();
    const manager = createTableCommandManager(store);

    manager.execute(new SetStyleCommand(range(0, 0, 2, 2), { bold: true }));
    expect(styleAt(store, styles, 2, 2)).toEqual({ bold: true });

    manager.undo();

    // 一次 undo 之后九格全部回到无样式
    for (let row = 0; row <= 2; row += 1) {
      for (let column = 0; column <= 2; column += 1) {
        expect(styleAt(store, styles, row, column)).toBeUndefined();
      }
    }
  });

  it("先写值再改样式，样式不会把值盖掉——读的是影子模型", () => {
    const { store, styles } = createStore();
    const manager = createTableCommandManager(store);

    manager.transaction(() => {
      manager.execute(new SetCellCommand(0, 0, "新值"));
      manager.execute(new SetStyleCommand(cell(0, 0), { bold: true }));
    });

    expect(store.getCell(0, 0)).toBe("新值");
    expect(styleAt(store, styles, 0, 0)).toEqual({ bold: true });
  });

  it("样式没变的格子不产生写入——重复设同一个样式是无操作", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);

    manager.execute(new SetStyleCommand(cell(0, 0), { bold: true }));
    const before = store.version;
    manager.execute(new SetStyleCommand(cell(0, 0), { bold: true }));

    expect(store.version).toBe(before);
  });

  it("非法 numFmt 在构造命令时就抛，事务根本不会开始", () => {
    expect(() => new SetStyleCommand(cell(0, 0), { numFmt: "[[[" })).not.toThrow();
    // 校验发生在 mutation 构造期（plan 阶段），不是命令构造期
    const { store } = createStore();
    const manager = createTableCommandManager(store);
    let thrown: unknown;
    try {
      manager.execute(new SetStyleCommand(cell(0, 0), { numFmt: "[[[" }));
    } catch (error) {
      thrown = error;
    }
    expect(thrown).toBeInstanceOf(CyExcelError);
    expect((thrown as CyExcelError).code).toBe("invalid-number-format");
    expect(store.sheet.getRecord(0, 0)?.s).toBeUndefined();
  });

  it("合法 numFmt 存进样式", () => {
    const { store, styles } = createStore();
    const manager = createTableCommandManager(store);

    manager.execute(new SetStyleCommand(cell(0, 0), { numFmt: "#,##0.00" }));

    expect(styleAt(store, styles, 0, 0)).toEqual({ numFmt: "#,##0.00" });
  });
});
