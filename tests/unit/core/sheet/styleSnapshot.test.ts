import { describe, expect, it } from "vitest";
import { SetStyleCommand } from "../../../../src/core/commands";
import { TableStore } from "../../../../src/core/data";
import { createTableCommandManager } from "../../../../src/core/integration";
import { WorkbookModel } from "../../../../src/core/sheet/WorkbookModel";
import { rowsToWorkbookSnapshot } from "../../../../src/core/sheet/import";

function createWorkbook(): WorkbookModel {
  return WorkbookModel.fromSnapshot(
    rowsToWorkbookSnapshot([{ a: "A0" }, { a: "A1" }, { a: "A2" }], ["a"])
  );
}

function managerFor(workbook: WorkbookModel) {
  const store = new TableStore(workbook.sheet(), { styles: workbook.styles() });
  return { store, manager: createTableCommandManager(store) };
}

const cell = (row: number, column: number) => ({
  startRow: row,
  startColumn: column,
  endRow: row,
  endColumn: column
});

describe("样式的快照往返", () => {
  it("写进去的样式出现在快照的 styles 表里，格子只存 id", () => {
    const workbook = createWorkbook();
    const { manager } = managerFor(workbook);

    manager.execute(new SetStyleCommand(cell(0, 0), { bold: true }));
    const snapshot = workbook.toSnapshot();

    const entries = Object.entries(snapshot.styles);
    expect(entries).toHaveLength(1);
    expect(entries[0]?.[1]).toEqual({ bold: true });

    const sheetSnapshot = snapshot.sheets[snapshot.sheetOrder[0] as string];
    const record = sheetSnapshot?.cellData[0]?.[0];
    // 格子里是 id，不是内联的样式对象（ADR 0004 §3.4 否掉的那个联合类型）
    expect(record?.s).toBe(entries[0]?.[0]);
  });

  it("加载快照后样式还在，并且能继续 intern 命中它", () => {
    const workbook = createWorkbook();
    const { manager } = managerFor(workbook);
    manager.execute(new SetStyleCommand(cell(0, 0), { bold: true }));

    const reloaded = WorkbookModel.fromSnapshot(workbook.toSnapshot());
    const styleId = reloaded.sheet().getRecord(0, 0)?.s;

    expect(styleId).toBeDefined();
    expect(reloaded.styles().get(styleId as string)).toEqual({ bold: true });
    // 同一个样式在加载后的表里 intern，命中已有条目而不是造出第二条
    expect(reloaded.styles().intern({ bold: true })).toBe(styleId);
    expect(reloaded.styles().size).toBe(1);
  });

  it("没人引用的样式不写进快照——这是唯一的回收时机", () => {
    const workbook = createWorkbook();
    const { manager } = managerFor(workbook);

    manager.execute(new SetStyleCommand(cell(0, 0), { bold: true }));
    manager.execute(new SetStyleCommand(cell(0, 0), { bold: null, italic: true }));

    // 内存里两条都还在（undo 回去还要用）
    expect(workbook.styles().size).toBe(2);
    // 但快照里只剩仍被引用的那条
    const styles = Object.values(workbook.toSnapshot().styles);
    expect(styles).toEqual([{ italic: true }]);
  });

  it("多格共用一个样式时快照里也只有一条", () => {
    const workbook = createWorkbook();
    const { manager } = managerFor(workbook);

    manager.execute(
      new SetStyleCommand({ startRow: 0, startColumn: 0, endRow: 2, endColumn: 0 }, { bold: true })
    );

    expect(Object.keys(workbook.toSnapshot().styles)).toHaveLength(1);
  });

  it("样式的内存开销只跟设过样式的格子走，与总格数无关", () => {
    // 坐标空间很大、占用格很少：styles 表的条目数只反映"有几种样式"，
    // 与行数、列数、未设样式的格子数都没有关系（退出条件第 8 条）。
    const workbook = WorkbookModel.createEmpty({
      sheet: { blankRows: 1_000_000, blankColumns: 50 }
    });
    const { manager } = managerFor(workbook);

    manager.execute(new SetStyleCommand(cell(0, 0), { bold: true }));
    manager.execute(new SetStyleCommand(cell(500_000, 10), { bold: true }));

    expect(workbook.styles().size).toBe(1);
    expect(Object.keys(workbook.toSnapshot().styles)).toHaveLength(1);
  });
});
