import { describe, expect, it } from "vitest";
import { CyExcelError } from "../../../../src/core/errors";
import { applySheetMutation, estimateSheetMutationBytes } from "../../../../src/core/sheet/apply";
import type { SheetMutation } from "../../../../src/core/sheet/mutation";
import { SheetModel } from "../../../../src/core/sheet/SheetModel";

/**
 * `applySheetMutation` 的**逐成员正确性**不在这个文件里测，在 `invert.test.ts`：那边
 * 每一条"施加 m、再施加逆、比 `WorkbookModel.toSnapshot()`"的往返用例施加的都是这份
 * 生产代码，13 个成员一条不落。在这里再抄一遍等于两份测试测同一件事。
 *
 * 这个文件只装三样它测不到的东西：`writeTables` 那道守卫、坐标空间的钉法，和
 * `estimateSheetMutationBytes`。坐标空间那一组之所以在这里而不在往返用例里，是因为
 * 往返是**对称**的：轴多涨了 1 再多缩 1，快照仍然逐字段一致，往返查不出来。
 */

function modelWithoutTables(): SheetModel {
  const model = SheetModel.createEmpty({ id: "s1", name: "S1" });
  model.setRecord(0, 0, { v: "a" });
  return model;
}

describe("applySheetMutation — writeTables 的守卫", () => {
  it("点名一张没登记的表：抛，而不是静默跳过", () => {
    // 静默跳过丢掉的正是 `tables` 这个字段存在的理由（spec §11 不变量 8）。这条路
    // 在正常流程上到不了——`invertSheetMutation` 的 prepare 期已经查过——所以它是
    // 一道守卫，不是主路径；守卫也得有用例，否则改坏了没人知道。
    const mutation: SheetMutation = {
      type: "restore-rows",
      start: 0,
      rows: [{}],
      meta: [undefined],
      tables: { "com.acme.nope": ["x"] },
      mergesBefore: []
    };
    expect(() => {
      applySheetMutation(modelWithoutTables(), mutation);
    }).toThrow(CyExcelError);
  });
});

/**
 * 一条变更把轴改动多少，只由它是**哪一条**决定，不由它恰好搬动了多少占用范围决定。
 * 见 `withCoordinateSpace` 的注释。
 */
describe("applySheetMutation — 坐标空间", () => {
  function twoByTwo(): SheetModel {
    const model = SheetModel.createEmpty({ blankRows: 0, blankColumns: 0 });
    model.setRecord(0, 0, { v: "a" });
    model.setRecord(1, 1, { v: "b" });
    return model;
  }

  it("插一行带内容的：轴只涨 1，不是插入与写值各涨一次", () => {
    const model = twoByTwo();
    // 追加在末尾且带内容：`insertRows` 把轴涨到 3，随后的写值把占用范围也推到 3,
    // 不钉住的话 `blank` 留在原地，轴变成 4
    applySheetMutation(model, { type: "insert-rows", start: 2, rows: [{ 0: { v: "c" } }] });

    expect(model.rowCount()).toBe(3);
    expect(model.getCell(2, 0)).toBe("c");
  });

  it("插一列：轴涨 1，行轴不动", () => {
    const model = twoByTwo();
    applySheetMutation(model, { type: "insert-columns", start: 2, count: 1 });

    expect([model.rowCount(), model.columnCount()]).toEqual([2, 3]);
  });

  it("写元数据不改轴——哪怕写的是空白区里的那一行", () => {
    const model = SheetModel.createEmpty({ blankRows: 4, blankColumns: 4 });
    model.setRecord(0, 0, { v: "a" });
    expect([model.rowCount(), model.columnCount()]).toEqual([5, 5]);

    // `rowData` / `columnData` 是登记在册的侧表，占用范围算它们一份——不钉的话
    // 给第 3 行设个行高就能把轴推到 8
    applySheetMutation(model, { type: "set-row-meta", row: 3, meta: { h: 44 } });
    applySheetMutation(model, { type: "set-column-meta", column: 3, meta: { w: 90 } });

    expect([model.rowCount(), model.columnCount()]).toEqual([5, 5]);
    expect(model.getRowMeta(3)).toEqual({ h: 44 });
    expect(model.getColumnMeta(3)).toEqual({ w: 90 });
  });
});

/**
 * `estimateSheetMutationBytes` 是个**粗估**，唯一的消费者是历史栈的内存上限。所以这里
 * 钉的不是具体数值——钉住常数等于把一次调参变成一次改测试——而是三条它必须满足的性质：
 *
 * 1. 每个成员都得给一个正的有限数（少写一个 arm，TypeScript 会拦；但一个返回 `NaN`
 *    的算术它拦不住，而 `NaN` 会让历史栈的上限判断永远为假）；
 * 2. **不随表长增长的成员，估值必须是常数**——`delete-rows(0, 10000)` 与
 *    `delete-rows(0, 1)` 一样便宜，因为它就只有两个下标；
 * 3. **随内容增长的成员，估值必须跟着长**，而且要同阶：带回一百行记录的
 *    `restore-rows` 得比 `delete-rows` 贵两个数量级，否则这个闸门形同虚设。
 */
describe("estimateSheetMutationBytes", () => {
  const everyMember: readonly SheetMutation[] = [
    { type: "set-cell", row: 0, column: 0, record: { v: 1 } },
    { type: "set-range", range: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 }, records: [[{ v: 1 }]] },
    { type: "insert-rows", start: 0, rows: [{ 0: { v: 1 } }] },
    { type: "delete-rows", start: 0, count: 1 },
    {
      type: "restore-rows",
      start: 0,
      rows: [{ 0: { v: 1 } }],
      meta: [{ h: 20 }],
      tables: {},
      mergesBefore: []
    },
    { type: "move-rows", start: 0, count: 1, target: 2 },
    { type: "reorder-rows", order: Int32Array.from([1, 0]) },
    { type: "insert-columns", start: 0, count: 1 },
    { type: "delete-columns", start: 0, count: 1 },
    {
      type: "restore-columns",
      start: 0,
      count: 1,
      cells: [{ 0: { v: 1 } }],
      meta: [{ w: 10 }],
      tables: {},
      mergesBefore: []
    },
    { type: "move-columns", start: 0, count: 1, target: 2 },
    { type: "set-row-meta", row: 0, meta: { h: 20 } },
    { type: "set-column-meta", column: 0, meta: { w: 10 } },
    { type: "set-merges", merges: [{ startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 }] }
  ];

  it("每个成员都给一个正的有限数", () => {
    // 反空洞：这个数组必须真的覆盖了整个联合，少一个成员这条就是在给 13 个背书
    expect(new Set(everyMember.map((m) => m.type)).size).toBe(14);
    for (const mutation of everyMember) {
      const bytes = estimateSheetMutationBytes(mutation);
      expect(Number.isFinite(bytes)).toBe(true);
      expect(bytes).toBeGreaterThan(0);
    }
  });

  it("只带下标的成员，估值不随 count 变", () => {
    expect(estimateSheetMutationBytes({ type: "delete-rows", start: 0, count: 10000 })).toBe(
      estimateSheetMutationBytes({ type: "delete-rows", start: 0, count: 1 })
    );
  });

  it("带内容的成员，估值跟着内容长，而且同阶", () => {
    const rows = Array.from({ length: 100 }, () => ({ 0: { v: "x" }, 1: { v: "y" } }));
    const restore = estimateSheetMutationBytes({
      type: "restore-rows",
      start: 0,
      rows,
      meta: [],
      tables: {},
      mergesBefore: []
    });
    const del = estimateSheetMutationBytes({ type: "delete-rows", start: 0, count: 100 });
    expect(restore).toBeGreaterThan(del * 100);
  });

  it("`undefined` 的载荷不计入", () => {
    const withRecord = estimateSheetMutationBytes({ type: "set-cell", row: 0, column: 0, record: { v: 1 } });
    const cleared = estimateSheetMutationBytes({ type: "set-cell", row: 0, column: 0, record: undefined });
    expect(cleared).toBeLessThan(withRecord);
    const withMeta = estimateSheetMutationBytes({ type: "set-row-meta", row: 0, meta: { h: 20 } });
    const clearedMeta = estimateSheetMutationBytes({ type: "set-row-meta", row: 0, meta: undefined });
    expect(clearedMeta).toBeLessThan(withMeta);
  });

  it("reorder-rows 跟着 order 的长度长——那是它里面唯一随表长增长的东西", () => {
    const short = estimateSheetMutationBytes({ type: "reorder-rows", order: Int32Array.from([1, 0]) });
    const long = estimateSheetMutationBytes({
      type: "reorder-rows",
      order: Int32Array.from({ length: 1000 }, (_, index) => index)
    });
    expect(long).toBeGreaterThan(short);
  });

  it("插件侧表的条目计入 restore-* 的估值", () => {
    const base: SheetMutation = {
      type: "restore-rows",
      start: 0,
      rows: [{}],
      meta: [undefined],
      tables: {},
      mergesBefore: []
    };
    const withEntries: SheetMutation = {
      type: "restore-rows",
      start: 0,
      rows: [{}],
      meta: [undefined],
      tables: { "com.acme.rowKeys": ["k"] },
      mergesBefore: []
    };
    expect(estimateSheetMutationBytes(withEntries)).toBeGreaterThan(estimateSheetMutationBytes(base));
  });
});
