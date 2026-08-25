import { describe, expect, it } from "vitest";
import { IndexedTable } from "../../../../src/core/sheet/IndexedTable";
import { buildDeleteColumnsShift, buildInsertColumnsShift } from "../../../../src/core/sheet/columnShift";
import { ROW_PLAN_HOLE } from "../../../../src/core/sheet/rowPlan";

const H = ROW_PLAN_HOLE;

describe("IndexedTable — 存取", () => {
  it("是稀疏的：没设过的下标不出现在 read() 里", () => {
    const table = new IndexedTable<number>("t", "row");
    table.set(3, 30);
    expect(table.read()).toStrictEqual({ 3: 30 });
    expect(0 in table.read()).toBe(false);
  });

  it("delete 真的把键删掉，不是留一个 undefined", () => {
    const table = new IndexedTable<number>("t", "row", { 1: 10, 2: 20 });
    table.delete(1);
    expect(table.read()).toStrictEqual({ 2: 20 });
    expect(1 in table.read()).toBe(false);
  });

  it("构造时拷贝初始数据，之后改动不回写调用方的对象", () => {
    const initial = { 0: 1 };
    const table = new IndexedTable<number>("t", "row", initial);
    table.set(1, 2);
    expect(initial).toStrictEqual({ 0: 1 });
  });

  it("get 取单个下标，没设过的返回 undefined", () => {
    // 任务二的 SheetModel.getRowMeta / getColumnMeta 直接委托到这里，
    // 是这个类被调用得最频繁的读路径，得有自己的用例而不是靠 read() 顺带覆盖。
    const table = new IndexedTable<number>("t", "row", { 4: 40 });
    expect(table.get(4)).toBe(40);
    expect(table.get(0)).toBeUndefined();
  });
});

describe("IndexedTable — extent", () => {
  it("空表的 extent 是 0", () => {
    expect(new IndexedTable<number>("t", "row").extent()).toBe(0);
  });

  it("extent 是最大占用下标加一，不是键的个数", () => {
    const table = new IndexedTable<number>("t", "row", { 7: 1 });
    expect(table.extent()).toBe(8);
  });

  it("删掉最大的那个之后 extent 回退", () => {
    const table = new IndexedTable<number>("t", "row", { 2: 1, 9: 1 });
    table.delete(9);
    expect(table.extent()).toBe(3);
  });
});

describe("IndexedTable — 重键", () => {
  it("按行排列搬家，洞不产出键", () => {
    const table = new IndexedTable<string>("t", "row", { 0: "a", 1: "b" });
    // 在最前面插一行：新 0 是洞，旧 0 → 新 1，旧 1 → 新 2
    table.applyRowPlan(Int32Array.from([H, 0, 1]));
    expect(table.read()).toStrictEqual({ 1: "a", 2: "b" });
    expect(0 in table.read()).toBe(false);
  });

  it("排列没提到的旧下标被丢弃（那就是删除）", () => {
    const table = new IndexedTable<string>("t", "row", { 0: "a", 1: "b", 2: "c" });
    table.applyRowPlan(Int32Array.from([0, 2]));
    expect(table.read()).toStrictEqual({ 0: "a", 1: "c" });
  });

  it("按列移位搬家", () => {
    const table = new IndexedTable<string>("t", "column", { 0: "a", 2: "c" });
    table.applyColumnShift(buildInsertColumnsShift(1, 2));
    expect(table.read()).toStrictEqual({ 0: "a", 4: "c" });
  });

  it("列移位返回 undefined 的项被删掉", () => {
    const table = new IndexedTable<string>("t", "column", { 0: "a", 1: "b", 2: "c" });
    table.applyColumnShift(buildDeleteColumnsShift(1, 1));
    expect(table.read()).toStrictEqual({ 0: "a", 1: "c" });
  });

  it("重键之后 extent 跟着变", () => {
    const table = new IndexedTable<string>("t", "row", { 0: "a" });
    table.applyRowPlan(Int32Array.from([H, H, 0]));
    expect(table.extent()).toBe(3);
  });
});

describe("IndexedTable — 轴不匹配必须炸", () => {
  it("行表拒绝列移位", () => {
    const table = new IndexedTable<number>("t", "row");
    expect(() => table.applyColumnShift(buildInsertColumnsShift(0, 1))).toThrow(/row/);
  });

  it("列表拒绝行排列", () => {
    const table = new IndexedTable<number>("t", "column");
    expect(() => table.applyRowPlan(Int32Array.from([0]))).toThrow(/column/);
  });
});
