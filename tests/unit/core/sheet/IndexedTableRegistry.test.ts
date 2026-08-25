import { describe, expect, it } from "vitest";
import { IndexedTable } from "../../../../src/core/sheet/IndexedTable";
import { IndexedTableRegistry } from "../../../../src/core/sheet/IndexedTableRegistry";
import { buildInsertColumnsShift } from "../../../../src/core/sheet/columnShift";
import { ROW_PLAN_HOLE } from "../../../../src/core/sheet/rowPlan";
import { SheetModel } from "../../../../src/core/sheet/SheetModel";

const H = ROW_PLAN_HOLE;

describe("IndexedTableRegistry", () => {
  it("登记后可以按名字取回", () => {
    const registry = new IndexedTableRegistry();
    const table = registry.register(new IndexedTable<number>("a", "row"));
    expect(registry.table("a")).toBe(table);
    expect(registry.has("a")).toBe(true);
  });

  it("重名必须拒绝——两张表共用一个名字会让快照往返对不上", () => {
    const registry = new IndexedTableRegistry();
    registry.register(new IndexedTable<number>("a", "row"));
    expect(() => registry.register(new IndexedTable<number>("a", "column"))).toThrow(/a/);
  });

  it("一次行排列施加到**所有**行表上", () => {
    const registry = new IndexedTableRegistry();
    const first = registry.register(new IndexedTable<string>("first", "row", { 0: "x" }));
    const second = registry.register(new IndexedTable<string>("second", "row", { 0: "y" }));
    registry.applyRowPlan(Int32Array.from([H, 0]));
    expect(first.read()).toStrictEqual({ 1: "x" });
    expect(second.read()).toStrictEqual({ 1: "y" });
  });

  it("行排列不碰列表，列移位不碰行表", () => {
    const registry = new IndexedTableRegistry();
    const rowTable = registry.register(new IndexedTable<string>("r", "row", { 0: "x" }));
    const columnTable = registry.register(new IndexedTable<string>("c", "column", { 0: "y" }));
    registry.applyRowPlan(Int32Array.from([H, 0]));
    expect(columnTable.read()).toStrictEqual({ 0: "y" });
    registry.applyColumnShift(buildInsertColumnsShift(0, 1));
    expect(rowTable.read()).toStrictEqual({ 1: "x" });
  });

  it("rowExtent 是所有行表的并集，不是某一张的", () => {
    const registry = new IndexedTableRegistry();
    registry.register(new IndexedTable<string>("a", "row", { 2: "x" }));
    registry.register(new IndexedTable<string>("b", "row", { 9: "y" }));
    expect(registry.rowExtent()).toBe(10);
  });

  it("rowExtent 不把列表算进去", () => {
    const registry = new IndexedTableRegistry();
    registry.register(new IndexedTable<string>("c", "column", { 50: "y" }));
    expect(registry.rowExtent()).toBe(0);
  });
});

describe("SheetModel — 侧表登记处", () => {
  it("rowData / columnData 已经是登记处的成员", () => {
    const model = SheetModel.createEmpty();
    expect(model.tableNames()).toContain("rowData");
    expect(model.tableNames()).toContain("columnData");
  });

  it("外部登记的行表跟着结构操作一起搬——这就是 spec §7.2 的重键契约", () => {
    const model = SheetModel.createEmpty();
    const keys = model.registerTable(new IndexedTable<string>("rowKeys", "row", { 0: "k0", 1: "k1" }));
    model.insertRows(0, 1);
    expect(keys.read()).toStrictEqual({ 1: "k0", 2: "k1" });
  });

  it("外部登记的行表参与占用范围——只有格子的话它会被静默丢掉", () => {
    const model = SheetModel.createEmpty();
    model.setCell(0, 0, "a");
    const keys = model.registerTable(new IndexedTable<string>("rowKeys", "row", { 5: "k5" }));
    // 占用范围必须是 6（第 5 行有资源），而不是 1（只有第 0 行有格子）
    model.insertRows(0, 1);
    expect(keys.read()).toStrictEqual({ 6: "k5" });
  });

  it("外部登记的列表跟着列操作一起搬", () => {
    const model = SheetModel.createEmpty();
    const widths = model.registerTable(new IndexedTable<string>("colTag", "column", { 0: "c0" }));
    model.insertColumns(0, 2);
    expect(widths.read()).toStrictEqual({ 2: "c0" });
  });

  it("外部登记的行表也跟着 reorderRows 走", () => {
    const model = SheetModel.createEmpty();
    const keys = model.registerTable(new IndexedTable<string>("rowKeys", "row", { 0: "k0", 1: "k1" }));
    model.reorderRows(Int32Array.from([1, 0]));
    expect(keys.read()).toStrictEqual({ 0: "k1", 1: "k0" });
  });

  it("重名登记被拒绝——rowData 是占用的名字", () => {
    const model = SheetModel.createEmpty();
    expect(() => model.registerTable(new IndexedTable<string>("rowData", "row"))).toThrow(/rowData/);
  });
});
