import { describe, expect, it } from "vitest";
import { IndexedTable } from "../../../../src/core/sheet/IndexedTable";
import { invertSheetMutation } from "../../../../src/core/sheet/invert";
import type { RowCells, SheetMutation } from "../../../../src/core/sheet/mutation";
import { SheetModel } from "../../../../src/core/sheet/SheetModel";
import type { RowMeta } from "../../../../src/core/sheet/types";

/**
 * 本地的 `apply`——**4A 没有 `applySheetMutation`，那是 4B**。
 *
 * 它只用 `SheetModel` 已有的公开方法，逐条把 11 个成员施加下去。两条要点：
 *
 * 1. **`restore-rows` / `restore-columns` 是"插空行/列 + 逐格写回 + 写回元数据"
 *    的组合。** `SheetModel.insertRows` 确实只插空行、不带记录，但把记录与元数据
 *    补写回去用的也都是既有的公开方法（`setRecord` / `setRowMeta` /
 *    `setColumnMeta`），组合出来的效果与"带记录的插入"在快照上逐字段一致。
 * 2. **这个组合不会让精确性测试失去检出能力。** 它是按 mutation 的字段结构写的，
 *    不知道自己施加的是不是一条逆变更：`invertSheetMutation` 若抓错了行、抓漏了
 *    元数据、或把 `RowCells` / `ColumnCells` 写转置了，这里就会把格子放到错误的
 *    位置上，`toSnapshot()` 的逐字段比较随即失败。
 */
function apply(model: SheetModel, mutation: SheetMutation): void {
  switch (mutation.type) {
    case "set-cell":
      writeCell(model, mutation.row, mutation.column, mutation.record);
      return;
    case "set-range": {
      const { startRow, startColumn, endRow, endColumn } = mutation.range;
      for (let row = startRow; row <= endRow; row += 1) {
        for (let column = startColumn; column <= endColumn; column += 1) {
          writeCell(model, row, column, mutation.records[row - startRow]?.[column - startColumn]);
        }
      }
      return;
    }
    case "insert-rows":
      model.insertRows(mutation.start, mutation.rows.length);
      writeRows(model, mutation.start, mutation.rows, undefined);
      return;
    case "restore-rows":
      model.insertRows(mutation.start, mutation.rows.length);
      writeRows(model, mutation.start, mutation.rows, mutation.meta);
      return;
    case "delete-rows":
      model.deleteRows(mutation.start, mutation.count);
      return;
    case "move-rows":
      model.moveRows(mutation.start, mutation.count, mutation.target);
      return;
    case "reorder-rows":
      model.reorderRows(mutation.order);
      return;
    case "insert-columns":
      model.insertColumns(mutation.start, mutation.count);
      return;
    case "delete-columns":
      model.deleteColumns(mutation.start, mutation.count);
      return;
    case "restore-columns": {
      model.insertColumns(mutation.start, mutation.count);
      for (let offset = 0; offset < mutation.count; offset += 1) {
        const column = mutation.cells[offset];
        if (column) {
          // `ColumnCells` 的键是**行号**——写成列号就是一次静默转置
          for (const rowKey of Object.keys(column)) {
            writeCell(model, Number(rowKey), mutation.start + offset, column[Number(rowKey)]);
          }
        }
        const meta = mutation.meta[offset];
        if (meta !== undefined) {
          model.setColumnMeta(mutation.start + offset, meta);
        }
      }
      return;
    }
    case "move-columns":
      model.moveColumns(mutation.start, mutation.count, mutation.target);
      return;
  }
}

function writeCell(model: SheetModel, row: number, column: number, record: unknown): void {
  if (record === undefined) {
    // spec §11 不变量 1：未占用的格子在 `cellData` 中不存在，不是存一条空记录
    model.deleteCell(row, column);
  } else {
    model.setRecord(row, column, record as Parameters<SheetModel["setRecord"]>[2]);
  }
}

function writeRows(
  model: SheetModel,
  start: number,
  rows: readonly RowCells[],
  meta: readonly (RowMeta | undefined)[] | undefined
): void {
  for (let offset = 0; offset < rows.length; offset += 1) {
    const cells = rows[offset];
    if (cells) {
      // `RowCells` 的键是**列号**
      for (const columnKey of Object.keys(cells)) {
        writeCell(model, start + offset, Number(columnKey), cells[Number(columnKey)]);
      }
    }
    const rowMeta = meta?.[offset];
    if (rowMeta !== undefined) {
      model.setRowMeta(start + offset, rowMeta);
    }
  }
}

/**
 * 第 4 行**只有元数据、没有任何格子**，是故意的：只按格子算占用范围的实现会把它
 * 整条漏掉（`SheetModel.#rowExtentIncludingTables` 的注释记着这个洞在阶段二出现过
 * 四次）。`blankRows` / `blankColumns` 取小值，好让"删除范围伸进空白区"这类边界
 * 用得上的数字都小。
 */
function fixture(): SheetModel {
  const model = SheetModel.createEmpty({ blankRows: 2, blankColumns: 2 });
  model.setRecord(0, 0, { v: "a0" });
  model.setRecord(0, 1, { v: "b0", t: "text" });
  model.setRecord(1, 0, { v: "a1" });
  model.setRecord(1, 2, { v: "c1", s: "s1" });
  model.setRecord(2, 1, { v: "b2", f: "=1+1" });
  model.setRecord(3, 0, { v: "a3" });
  model.setRowMeta(1, { h: 40 });
  model.setRowMeta(4, { hd: true });
  model.setColumnMeta(0, { w: 100 });
  model.setColumnMeta(2, { w: 200, t: "number" });
  return model;
}

/** 施加 `m`，再施加在施加之前求出的逆，逐字段比对 `toSnapshot()`。 */
function expectExactRoundTrip(mutation: SheetMutation): void {
  const model = fixture();
  const before = model.toSnapshot();
  const inverse = invertSheetMutation(model, mutation);
  apply(model, mutation);
  apply(model, inverse);
  expect(model.toSnapshot()).toStrictEqual(before);
}

function expectNonInvasive(mutation: SheetMutation): void {
  const model = fixture();
  const before = model.toSnapshot();
  invertSheetMutation(model, mutation);
  expect(model.toSnapshot()).toStrictEqual(before);
}

describe("invertSheetMutation — 精确性", () => {
  it("格子：set-cell 与 set-range 往返后逐字段相等", () => {
    expectExactRoundTrip({ type: "set-cell", row: 0, column: 0, record: { v: "changed", t: "money" } });
    // 区域里混着"原本有记录"与"原本是空格"两种格子，逆必须两种都还原
    expectExactRoundTrip({
      type: "set-range",
      range: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 2 },
      records: [
        [{ v: 1 }, undefined, { v: 3 }],
        [undefined, { v: 5 }, undefined]
      ]
    });
  });

  it("删行：含只有元数据、没有格子的那一行", () => {
    expectExactRoundTrip({ type: "delete-rows", start: 1, count: 2 });
    // 3..4 里第 4 行只有 rowData——逆若只抓格子，这一行的元数据回不来
    expectExactRoundTrip({ type: "delete-rows", start: 3, count: 2 });
  });

  it("删列：格子按行号回位，列元数据一并回来", () => {
    expectExactRoundTrip({ type: "delete-columns", start: 0, count: 2 });
    expectExactRoundTrip({ type: "delete-columns", start: 1, count: 1 });
  });

  it("reorder：逆排列把行与行元数据一起搬回", () => {
    expectExactRoundTrip({ type: "reorder-rows", order: Int32Array.from([2, 0, 1, 4, 3]) });
  });
});

describe("invertSheetMutation — 非侵入", () => {
  it("格子类的求逆不碰模型", () => {
    expectNonInvasive({ type: "set-cell", row: 0, column: 0, record: { v: "changed" } });
    expectNonInvasive({
      type: "set-range",
      range: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 2 },
      records: [[{ v: 1 }], []]
    });
  });

  it("删行的求逆不碰模型", () => {
    expectNonInvasive({ type: "delete-rows", start: 1, count: 2 });
  });

  it("删列的求逆不碰模型", () => {
    expectNonInvasive({ type: "delete-columns", start: 0, count: 2 });
  });

  it("move / reorder 的求逆不碰模型", () => {
    expectNonInvasive({ type: "move-rows", start: 1, count: 2, target: 4 });
    expectNonInvasive({ type: "move-columns", start: 0, count: 1, target: 3 });
    expectNonInvasive({ type: "reorder-rows", order: Int32Array.from([2, 0, 1, 4, 3]) });
  });
});

describe("invertSheetMutation — 边界", () => {
  it("删除范围伸进空白区：逆只带回落在占用范围内的那部分，往返仍然精确", () => {
    // 占用到第 4 行（只有元数据的那行），坐标空间到第 6 行——3..6 有两行在空白区
    expectExactRoundTrip({ type: "delete-rows", start: 3, count: 4 });
    // 列轴同理：占用到第 2 列，坐标空间到第 4 列
    expectExactRoundTrip({ type: "delete-columns", start: 2, count: 3 });
  });

  it("set-cell 写进空格时，逆的 record 是 undefined（施加下去是删格，不是写空记录）", () => {
    const model = fixture();
    const mutation: SheetMutation = { type: "set-cell", row: 2, column: 2, record: { v: "new" } };
    expect(invertSheetMutation(model, mutation)).toStrictEqual({
      type: "set-cell",
      row: 2,
      column: 2,
      record: undefined
    });
    expectExactRoundTrip(mutation);
  });

  it("reorder 的逆排列再逆一次等于原排列，且原排列没被就地改动", () => {
    const model = fixture();
    const order = Int32Array.from([2, 0, 1, 4, 3]);
    const once = invertSheetMutation(model, { type: "reorder-rows", order });
    expect(once).toStrictEqual({ type: "reorder-rows", order: Int32Array.from([1, 2, 0, 4, 3]) });
    expect(invertSheetMutation(model, once)).toStrictEqual({ type: "reorder-rows", order });
    expect(order).toStrictEqual(Int32Array.from([2, 0, 1, 4, 3]));
  });

  it("move 的 target 恰在块尾是恒等，两个方向的真移动各自回到原位", () => {
    // target === start + count：块摘出去又插回原处，正变更与逆变更都是恒等
    expect(invertSheetMutation(fixture(), { type: "move-rows", start: 1, count: 2, target: 3 })).toStrictEqual({
      type: "move-rows",
      start: 1,
      count: 2,
      target: 1
    });
    expectExactRoundTrip({ type: "move-rows", start: 1, count: 2, target: 3 });

    // 后移。ArrayDataSource.test.ts:207 那条 `moveRows(1,2,4)`：A,B,C,D,E → A,D,B,C,E，
    // 块落到下标 2，逆把它从 2 搬回"原下标 1 之前"
    expect(invertSheetMutation(fixture(), { type: "move-rows", start: 1, count: 2, target: 4 })).toStrictEqual({
      type: "move-rows",
      start: 2,
      count: 2,
      target: 1
    });
    expectExactRoundTrip({ type: "move-rows", start: 1, count: 2, target: 4 });

    // 前移。move(3,1,0)：A,B,C,D,E → D,A,B,C,E，逆的 target 要 +count 才落回原处
    expect(invertSheetMutation(fixture(), { type: "move-rows", start: 3, count: 1, target: 0 })).toStrictEqual({
      type: "move-rows",
      start: 0,
      count: 1,
      target: 4
    });
    expectExactRoundTrip({ type: "move-rows", start: 3, count: 1, target: 0 });

    // 列轴共用同一行算术（buildMoveColumnsShift），逆的口径因此也共用
    expectExactRoundTrip({ type: "move-columns", start: 0, count: 1, target: 3 });
    expectExactRoundTrip({ type: "move-columns", start: 2, count: 1, target: 0 });
  });
});

describe("invertSheetMutation — 表达不出来的逆一律拒绝", () => {
  it("删除的逆带不回插件登记的侧表，宁可拒绝也不给会丢数据的近似", () => {
    const model = fixture();
    model.registerTable(new IndexedTable<string>("com.acme.rowKeys", "row", { 1: "k1" }));
    model.registerTable(new IndexedTable<string>("com.acme.colTags", "column", { 0: "t0" }));

    // restore-rows / restore-columns 没有装侧表条目的字段（任务二的类型），
    // 而 SheetModel.toSnapshot() 又不序列化这些表，精确性测试对这个缺口是瞎的
    expect(() => invertSheetMutation(model, { type: "delete-rows", start: 1, count: 1 })).toThrow(
      /com\.acme\.rowKeys/
    );
    expect(() => invertSheetMutation(model, { type: "delete-columns", start: 0, count: 1 })).toThrow(
      /com\.acme\.colTags/
    );

    // 只挡对应轴：行表不该挡住列删除的求逆，反之亦然
    const rowOnly = fixture();
    rowOnly.registerTable(new IndexedTable<string>("com.acme.rowKeys", "row", { 1: "k1" }));
    expect(() => invertSheetMutation(rowOnly, { type: "delete-columns", start: 0, count: 1 })).not.toThrow();

    // 拒绝也必须是非侵入的——抛错之前一个字都不许写
    const before = model.toSnapshot();
    expect(() => invertSheetMutation(model, { type: "delete-rows", start: 1, count: 1 })).toThrow();
    expect(model.toSnapshot()).toStrictEqual(before);
  });
});
