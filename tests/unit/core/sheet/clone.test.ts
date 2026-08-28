import { describe, expect, it } from "vitest";
import { invertSheetMutation } from "../../../../src/core/sheet/invert";
import type { SheetMutation } from "../../../../src/core/sheet/mutation";
import type { SheetModel } from "../../../../src/core/sheet/SheetModel";
import { WorkbookModel } from "../../../../src/core/sheet/WorkbookModel";

/**
 * 一张够小、但每一类状态都占着的表：格子、行元数据、列元数据、两张插件侧表（行轴与
 * 列轴各一）。第 4 行**只有元数据、没有格子**，与 `invert.test.ts` 用的是同一个刻意
 * 安排——只按格子算占用范围的实现会把它整条漏掉。
 */
function workbookFixture(): WorkbookModel {
  const workbook = WorkbookModel.createEmpty({ sheet: { blankRows: 2, blankColumns: 2 } });
  const model = workbook.sheet();
  model.setRecord(0, 0, { v: "a0" });
  model.setRecord(0, 1, { v: "b0" });
  model.setRecord(1, 0, { v: "a1" });
  model.setRecord(3, 0, { v: "a3" });
  model.setRowMeta(1, { h: 40 });
  model.setRowMeta(4, { hd: true });
  model.setColumnMeta(0, { w: 100 });
  workbook.addResource({ name: "com.acme.rowKeys", indexedBy: "row", data: { 1: "k1", 3: "k3" } });
  workbook.addResource({ name: "com.acme.colTags", indexedBy: "column", data: { 0: "t0" } });
  return workbook;
}

function fixture(): SheetModel {
  return workbookFixture().sheet();
}

/** 影子与活模型的全部格子逐一比对——`SheetModel` 没有公开的稀疏迭代，按占用范围扫。 */
function cellsOf(model: SheetModel): Record<string, unknown> {
  const cells: Record<string, unknown> = {};
  for (let row = 0; row < model.occupiedRowCount(); row += 1) {
    for (let column = 0; column < model.occupiedColumnCount(); column += 1) {
      const record = model.getRecord(row, column);
      if (record !== undefined) {
        cells[`${row},${column}`] = record;
      }
    }
  }
  return cells;
}

describe("SheetModel.clone — 克隆出来的是同一份状态", () => {
  it("快照逐字段相等", () => {
    const live = fixture();
    expect(live.clone().toSnapshot()).toStrictEqual(live.toSnapshot());
  });

  it("记录按引用共享，不是深拷贝——浅共享规则", () => {
    // 深拷贝正是这个数据模型要省掉的分配开销（`types.ts` 顶部那条规则）。这条钉住
    // 它，免得有人"顺手"把 clone 改成深拷贝。
    const live = fixture();
    expect(live.clone().getRecord(0, 0)).toBe(live.getRecord(0, 0));
  });

  it("插件侧表跟着克隆——这一条失效是静默的", () => {
    // 影子的登记处若是空的，`capturePluginTables` 返回 `{}`：不报错、测试全绿、
    // 插件侧表的撤销数据静默丢失（spec §11 不变量 8 被整条架空）。所以这里不能只
    // 断言"表名对得上"，要一路断言到求逆的产物上。
    const shadow = fixture().clone();
    expect(shadow.tableNames()).toContain("com.acme.rowKeys");
    expect(shadow.table("com.acme.rowKeys")?.get(1)).toBe("k1");

    const inverse = invertSheetMutation(shadow, { type: "delete-rows", start: 1, count: 1 });
    const tables = (inverse as { tables: Record<string, readonly unknown[]> }).tables;
    expect(Object.keys(tables)).toContain("com.acme.rowKeys");
    expect(tables["com.acme.rowKeys"]).toStrictEqual(["k1"]);
  });
});

/**
 * 隔离性：写任一边，另一边不动。
 *
 * **每条都要写到一个"已经有格子的行"上**——那才是共享行桶的情形，也才是写时复制
 * 守卫唯一起作用的地方。写到空行上会新建一个桶，那条路本来就不共享，测不出守卫在
 * 不在。种缺陷验过：去掉 `CellMatrix.#privatize` 的拷贝，这一组必须变红。
 */
describe("SheetModel.clone — 隔离性", () => {
  it("往影子已有的行里写格子，活模型不动", () => {
    const live = fixture();
    const before = live.toSnapshot();
    const shadow = live.clone();
    shadow.setRecord(0, 1, { v: "changed" });
    shadow.setRecord(0, 5, { v: "new column in an existing bucket" });
    expect(live.toSnapshot()).toStrictEqual(before);
    expect(live.getRecord(0, 1)).toStrictEqual({ v: "b0" });
    expect(live.getRecord(0, 5)).toBeUndefined();
  });

  it("**反过来也要成立**：往活模型已有的行里写，影子不动", () => {
    // 共享是对称的。只给克隆那一边打标记的实现能过上一条，过不了这一条。
    const live = fixture();
    const shadow = live.clone();
    const before = shadow.toSnapshot();
    live.setRecord(0, 1, { v: "changed" });
    expect(shadow.toSnapshot()).toStrictEqual(before);
    expect(shadow.getRecord(0, 1)).toStrictEqual({ v: "b0" });
  });

  it("在影子上删格子，活模型不动", () => {
    const live = fixture();
    const before = live.toSnapshot();
    const shadow = live.clone();
    shadow.deleteCell(0, 0);
    expect(shadow.getRecord(0, 0)).toBeUndefined();
    expect(live.toSnapshot()).toStrictEqual(before);
  });

  it("行/列元数据与插件侧表同样隔离", () => {
    const live = fixture();
    const before = live.toSnapshot();
    const liveTable = live.table("com.acme.rowKeys");
    const shadow = live.clone();
    shadow.setRowMeta(1, { h: 999 });
    shadow.setColumnMeta(0, { w: 999 });
    shadow.table("com.acme.rowKeys")?.set(1, "changed");
    expect(live.toSnapshot()).toStrictEqual(before);
    expect(liveTable?.get(1)).toBe("k1");
  });

  it("影子上的结构操作不动活模型", () => {
    const live = fixture();
    const before = live.toSnapshot();
    const shadow = live.clone();
    shadow.deleteRows(0, 2);
    shadow.insertColumns(0, 1);
    expect(live.toSnapshot()).toStrictEqual(before);
  });

  /**
   * 共享标记是**按行号**记的，而结构操作正在换行号。不跟着搬的话，一个共享桶换个
   * 位置就成了"没标记"，紧接着的那次写会直接改在活模型看得见的那个桶上——现象是
   * 撤销之后活模型里凭空多出几个格子。
   */
  it("结构操作把桶搬到新行号之后，共享标记要跟着搬", () => {
    const live = fixture();
    const before = live.toSnapshot();
    const shadow = live.clone();
    // 第 0 行的桶被搬到第 1 行；它仍然是与活模型共享的那一个
    shadow.insertRows(0, 1);
    shadow.setRecord(1, 1, { v: "written after the bucket moved" });
    expect(live.toSnapshot()).toStrictEqual(before);
    expect(live.getRecord(0, 1)).toStrictEqual({ v: "b0" });
  });

  it("列操作之后同理", () => {
    const live = fixture();
    const before = live.toSnapshot();
    const shadow = live.clone();
    shadow.insertColumns(0, 1);
    shadow.setRecord(0, 1, { v: "written after the columns shifted" });
    expect(live.toSnapshot()).toStrictEqual(before);
  });

  it("连着克隆两次，三份互不影响", () => {
    // 第二次克隆时，第一份影子里已经有被私有化过的桶了——标记要按当前状态重打，
    // 而不是并进上一次那份残留。
    const live = fixture();
    const first = live.clone();
    first.setRecord(0, 0, { v: "first" });
    const second = live.clone();
    second.setRecord(0, 0, { v: "second" });
    expect(live.getRecord(0, 0)).toStrictEqual({ v: "a0" });
    expect(first.getRecord(0, 0)).toStrictEqual({ v: "first" });
    expect(second.getRecord(0, 0)).toStrictEqual({ v: "second" });
    expect(cellsOf(live)).toStrictEqual(cellsOf(fixture()));
  });

  it("影子跑完一整条「求逆、施加、再施加逆」也不碰活模型", () => {
    const live = fixture();
    const before = live.toSnapshot();
    const shadow = live.clone();
    const mutation: SheetMutation = { type: "delete-rows", start: 0, count: 2 };
    const inverse = invertSheetMutation(shadow, mutation);
    shadow.deleteRows(0, 2);
    expect(live.toSnapshot()).toStrictEqual(before);
    // 反空洞：这条逆确实带着插件侧表的条目，不然下面是在给一条空 `tables` 背书
    expect(Object.keys((inverse as { tables: Record<string, unknown> }).tables)).toContain(
      "com.acme.rowKeys"
    );
  });
});
