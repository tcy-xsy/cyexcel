import { describe, expect, it } from "vitest";

import { SetCellCommand, SetCellImageCommand, SetStyleCommand } from "../../../../src/core/commands";
import { TableStore } from "../../../../src/core/data";
import { createTableCommandManager } from "../../../../src/core/integration";
import { StyleRegistry } from "../../../../src/core/style/StyleRegistry";
import { createSheetFromRows } from "../../../helpers/createSheetModel";

/**
 * 3 行 x 3 列，但**第三列整列没有值**——坐标空间里存在、却未占用的格子，
 * 用来验证"空格子也能只挂图片"这类场景（照抄 styleCommands.test.ts 的形状）。
 */
function createStore(): { store: TableStore; styles: StyleRegistry } {
  const fields = ["c0", "c1", "c2"];
  const rows = Array.from({ length: 3 }, (_unused, r) =>
    Object.fromEntries(["c0", "c1"].map((field, c) => [field, `R${r}C${c}`]))
  );
  const styles = new StyleRegistry();
  const store = new TableStore(createSheetFromRows(rows, fields), { styles });
  return { store, styles };
}

describe("SetCellImageCommand", () => {
  it("写单格图片，产出 cell-image 补丁而不是 cell-value/cell-style", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);

    const patches = manager.execute(new SetCellImageCommand(0, 0, "https://example.com/a.png"));

    expect(patches).toContainEqual({ type: "cell-image", row: 0, column: 0 });
    expect(patches).not.toContainEqual({ type: "cell-value", row: 0, column: 0 });
    expect(patches).not.toContainEqual({ type: "cell-style", row: 0, column: 0 });
    expect(store.sheet.getRecord(0, 0)?.img).toBe("https://example.com/a.png");
  });

  it("不碰这一格的值——图片和值是同一条记录里的两个字段", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);

    manager.execute(new SetCellImageCommand(0, 0, "https://example.com/a.png"));

    expect(store.getCell(0, 0)).toBe("R0C0");
    expect(store.sheet.getRecord(0, 0)?.img).toBe("https://example.com/a.png");
  });

  it("给空格子设图片会造出一条只有图片的记录，没有 v", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);

    // 第三列在坐标空间里但没有值
    manager.execute(new SetCellImageCommand(2, 2, "https://example.com/a.png"));

    expect(store.sheet.getRecord(2, 2)?.v).toBeUndefined();
    expect(store.sheet.getRecord(2, 2)?.img).toBe("https://example.com/a.png");
  });

  it("清空图片（url 传 undefined）后这一格不再有 img 字段，值还在", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);

    manager.execute(new SetCellImageCommand(0, 0, "https://example.com/a.png"));
    manager.execute(new SetCellImageCommand(0, 0, undefined));

    expect(store.sheet.getRecord(0, 0)?.img).toBeUndefined();
    expect(store.getCell(0, 0)).toBe("R0C0");
  });

  it("空格子的图片被清空时整条记录消失，回到未占用", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);

    manager.execute(new SetCellImageCommand(2, 2, "https://example.com/a.png"));
    manager.execute(new SetCellImageCommand(2, 2, undefined));

    expect(store.sheet.getRecord(2, 2)).toBeUndefined();
  });

  it("撤销回到原来的图片，重做再回去", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);

    manager.execute(new SetCellImageCommand(0, 0, "https://example.com/a.png"));
    manager.execute(new SetCellImageCommand(0, 0, "https://example.com/b.png"));

    manager.undo();
    expect(store.sheet.getRecord(0, 0)?.img).toBe("https://example.com/a.png");

    manager.undo();
    expect(store.sheet.getRecord(0, 0)?.img).toBeUndefined();

    manager.redo();
    expect(store.sheet.getRecord(0, 0)?.img).toBe("https://example.com/a.png");
  });

  it("图片没变的格子不产生写入——重复设同一个 URL 是无操作", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);

    manager.execute(new SetCellImageCommand(0, 0, "https://example.com/a.png"));
    const before = store.version;
    manager.execute(new SetCellImageCommand(0, 0, "https://example.com/a.png"));

    expect(store.version).toBe(before);
  });

  it("先写值再设图片，图片不会把值盖掉——读的是影子模型", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);

    manager.transaction(() => {
      manager.execute(new SetCellCommand(0, 0, "新值"));
      manager.execute(new SetCellImageCommand(0, 0, "https://example.com/a.png"));
    });

    expect(store.getCell(0, 0)).toBe("新值");
    expect(store.sheet.getRecord(0, 0)?.img).toBe("https://example.com/a.png");
  });

  it("先设样式再设图片，两者共存，互不覆盖", () => {
    const { store, styles } = createStore();
    const manager = createTableCommandManager(store);

    manager.transaction(() => {
      manager.execute(new SetStyleCommand({ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }, { bold: true }));
      manager.execute(new SetCellImageCommand(0, 0, "https://example.com/a.png"));
    });

    const styleId = store.sheet.getRecord(0, 0)?.s;
    expect(styleId === undefined ? undefined : styles.get(styleId)).toEqual({ bold: true });
    expect(store.sheet.getRecord(0, 0)?.img).toBe("https://example.com/a.png");
  });

  it("清空样式不会把只剩图片的记录误判成空记录删掉", () => {
    const { store } = createStore();
    const manager = createTableCommandManager(store);

    manager.transaction(() => {
      manager.execute(new SetStyleCommand({ startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 }, { bold: true }));
      manager.execute(new SetCellImageCommand(2, 2, "https://example.com/a.png"));
    });
    manager.execute(new SetStyleCommand({ startRow: 2, startColumn: 2, endRow: 2, endColumn: 2 }, { bold: null }));

    expect(store.sheet.getRecord(2, 2)?.s).toBeUndefined();
    expect(store.sheet.getRecord(2, 2)?.img).toBe("https://example.com/a.png");
  });
});
