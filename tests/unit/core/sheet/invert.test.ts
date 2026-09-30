import { describe, expect, it } from "vitest";
import { CyExcelError, CyExcelErrorCode } from "../../../../src/core/errors";
import { applySheetMutation as apply } from "../../../../src/core/sheet/apply";
import { invertSheetMutation } from "../../../../src/core/sheet/invert";
import type { SheetMutation } from "../../../../src/core/sheet/mutation";
import { SheetModel } from "../../../../src/core/sheet/SheetModel";
import { WorkbookModel } from "../../../../src/core/sheet/WorkbookModel";
import type { SheetSnapshot, WorkbookSnapshot } from "../../../../src/core/sheet/types";
import type { CellRange } from "../../../../src/core/types";
import { CellValueType } from "../../../../src/core/sheet/types";

/**
 * `applySheetMutation` 在 4A 是本文件里的一个本地 helper（当时的注释写着"4A 没有
 * `applySheetMutation`，那是 4B"）。4B 把它提升成了生产代码，见
 * `src/core/sheet/apply.ts`：暂存层的三条路——预览、commit、回滚——必须跑同一份施加
 * 逻辑，留在测试里就意味着生产另写一份。
 *
 * 对本文件的意义是纯收益：**下面每一条"施加 m、再施加 inv、比快照"的往返用例，从此
 * 测的是生产代码**。它仍然不会让精确性测试失去检出能力——那份施加逻辑是按 mutation
 * 的字段结构写的，不知道自己施加的是不是一条逆变更：`invertSheetMutation` 若抓错了
 * 行、抓漏了元数据、或把 `RowCells` / `ColumnCells` 写转置了，格子就会落到错误的位置
 * 上，`toSnapshot()` 的逐字段比较随即失败。
 */

/**
 * 第 4 行**只有元数据、没有任何格子**，是故意的：只按格子算占用范围的实现会把它
 * 整条漏掉（`SheetModel.#rowExtentIncludingTables` 的注释记着这个洞在阶段二出现过
 * 四次）。`blankRows` / `blankColumns` 取小值，好让"删除范围伸进空白区"这类边界
 * 用得上的数字都小。
 *
 * **这张表还挂着两张插件侧表**（spec §7.2 的 `indexedBy: "row" / "column"` 资源），
 * 行轴与列轴各一张。它们在这里，是为了让两个对照物真正看得见插件表——只有
 * `WorkbookModel.toSnapshot()` 序列化 `resources`，`SheetModel.toSnapshot()` 不。
 *
 * 条目的下标是挑过的：行表落在第 1、4 行，列表落在第 0、2 列，都没有越出格子与
 * `rowData` / `columnData` 已有的占用范围，所以 `rowExtent()` 仍是 5、`rowCount()`
 * 仍是 7、`columnCount()` 仍是 5——既有用例钉住的那几个数一个都没动。
 */
function workbookFixture(): WorkbookModel {
  const workbook = WorkbookModel.createEmpty({ sheet: { blankRows: 2, blankColumns: 2 } });
  const model = workbook.sheet();
  model.setRecord(0, 0, { v: "a0" });
  model.setRecord(0, 1, { v: "b0", t: CellValueType.STRING });
  model.setRecord(1, 0, { v: "a1" });
  model.setRecord(1, 2, { v: "c1", s: "s1" });
  model.setRecord(2, 1, { v: "b2", f: "=1+1" });
  model.setRecord(3, 0, { v: "a3" });
  model.setRowMeta(1, { h: 40 });
  model.setRowMeta(4, { hd: true });
  model.setColumnMeta(0, { w: 100 });
  model.setColumnMeta(2, { w: 200, t: "number" });
  workbook.addResource({ name: "com.acme.rowKeys", indexedBy: "row", data: { 1: "k1", 4: "k4" } });
  workbook.addResource({ name: "com.acme.colTags", indexedBy: "column", data: { 0: "t0", 2: "t2" } });
  return workbook;
}

function fixture(): SheetModel {
  return workbookFixture().sheet();
}

/**
 * 对照物必须**看得见**插件侧表。
 *
 * 任务三的四条精确性用例比的是 `SheetModel.toSnapshot()`，而那个方法**不序列化
 * 插件侧表**（它自己的注释写明理由：那些表的持久化形态是 workbook 级的
 * `resources`）。于是"删除的逆有没有把插件表的条目带回来"这件事，它们**看不见**
 * ——任务三那道拒绝守卫之所以必要，正是因为测试看不见它。spec §11 不变量 8 把
 * **逆操作**也算进核心的义务，所以"逐字段相等"里的"逐字段"必须含 `resources`。
 *
 * 这条前置断言不是装饰：`resources` 为空的 workbook 比出来照样相等，于是**一个
 * 空的守卫与一个有效的守卫长得一模一样**。任务三的评审在
 * `expectPrepareMatchesCommit` 上抓到过同一个形状（两边都不抛时
 * `undefined === undefined` 空过），所以这里写成断言，不写成注释。
 */
function assertOracleSeesPluginTables(snapshot: WorkbookSnapshot): void {
  const indexed = snapshot.resources.filter((entry) => entry.indexedBy !== "none");
  expect(indexed.length).toBeGreaterThan(0);
  for (const entry of indexed) {
    expect(Object.keys(entry.data as Record<string, unknown>).length).toBeGreaterThan(0);
  }
}

/**
 * 前置断言：被删区间**之内**与**之外**都真的有插件侧表的条目。
 *
 * 区间内的那条靠 `restore-*` 的 `tables` 字段带回来，区间外的那条靠
 * `insertRows` / `insertColumns` 自己的重键移回去——两条路是两套代码，都要走到。
 * 缺了这条断言，一条区间里没有任何条目的用例照样全绿，而它什么也没验证。
 */
function assertBandHasPluginEntries(
  model: SheetModel,
  axis: "row" | "column",
  start: number,
  count: number
): void {
  const coreTable = axis === "row" ? "rowData" : "columnData";
  let inside = 0;
  let outside = 0;
  for (const name of model.tableNames()) {
    if (name === coreTable) {
      continue;
    }
    const table = model.table(name);
    if (table?.indexedBy !== axis) {
      continue;
    }
    for (const key of Object.keys(table.read())) {
      const index = Number(key);
      if (index >= start && index < start + count) {
        inside += 1;
      } else {
        outside += 1;
      }
    }
  }
  expect(inside, `[${start}, ${start + count}) 区间内必须有插件${axis}表的条目`).toBeGreaterThan(0);
  expect(outside, `区间外也必须有插件${axis}表的条目`).toBeGreaterThan(0);
}

/**
 * 造一个与 `capturePluginTables` 产出**同形**的期望值：原型是 `null`。
 *
 * `toStrictEqual` 连原型一起比，而产出侧用的是 `Object.create(null)`——`invert.ts` 里
 * 那段注释写明了为什么必须是它（`"__proto__"` 是一个能一路走到那里的合法表名）。
 * 所以期望值这边不能图省事写对象字面量。
 *
 * **不要为了让它们过而把 `toStrictEqual` 降成 `toEqual`。** 那不是等价的写法：
 * `writeTables` 里"跳过 `undefined`"那条稀疏性守卫，唯一的检出途径就是
 * `toStrictEqual`（`tableToResource` 会把凭空造出来的键当作 `{ 2: undefined }` 一路
 * 带进快照，只有它分得清"值是 undefined"和"根本没这个键"）。量过：同一个 bug，
 * `toStrictEqual` 下 6 条红，换成 `toEqual` 只剩 1 条。
 */
function sideTables(entries: Record<string, readonly unknown[]>): Record<string, readonly unknown[]> {
  return Object.assign(Object.create(null) as Record<string, readonly unknown[]>, entries);
}

/** 从一条 `restore-*` 里取出 `tables` 的表名，供只关心"有谁/没谁"的形状断言使用。 */
function tableNamesOf(mutation: SheetMutation): string[] {
  return Object.keys((mutation as { tables: Record<string, unknown> }).tables);
}

interface RoundTripOptions {
  /** 在取 `before` 之前往 workbook 里再加点东西（例如第二张插件行表） */
  readonly seed?: (workbook: WorkbookModel) => void;
  /** 施加之前的前置断言，用来挡掉"比两张空表"那种空过 */
  readonly precondition?: (model: SheetModel) => void;
}

/** 施加 `m`，再施加在施加之前求出的逆，逐字段比对 `WorkbookModel.toSnapshot()`。 */
function expectExactRoundTrip(mutation: SheetMutation, options: RoundTripOptions = {}): void {
  const workbook = workbookFixture();
  options.seed?.(workbook);
  const model = workbook.sheet();
  options.precondition?.(model);
  const before = workbook.toSnapshot();
  assertOracleSeesPluginTables(before);
  const inverse = invertSheetMutation(model, mutation);
  apply(model, mutation);
  apply(model, inverse);
  expect(workbook.toSnapshot()).toStrictEqual(before);
}

/**
 * 同一张表，外加一个合并区。
 *
 * 只能走 `fromSnapshot`：`mergeData` 没有公开的写入口，`createEmpty` 建出来的表
 * 永远是空的合并区列表。
 */
function mergedFixture(): SheetModel {
  return SheetModel.fromSnapshot({
    ...fixture().toSnapshot(),
    mergeData: [{ startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 }]
  });
}

/**
 * `mergedFixture` 的 workbook 版本——需要 `WorkbookModel.toSnapshot()` 才能看见
 * 插件侧表（见 `assertOracleSeesPluginTables`），往返用例因此要走这一个而不是
 * `mergedFixture`。合并区与 `mergedFixture` 同一个：`[0,0]`-`[1,1]`。
 */
function mergedWorkbookFixture(): WorkbookModel {
  const snapshot = workbookFixture().toSnapshot();
  const sheetId = snapshot.sheetOrder[0] as string;
  const sheet: SheetSnapshot = snapshot.sheets[sheetId] as SheetSnapshot;
  const mergedSheet: SheetSnapshot = {
    ...sheet,
    mergeData: [{ startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 }]
  };
  const sheets: Record<string, SheetSnapshot> = { ...snapshot.sheets, [sheetId]: mergedSheet };
  return WorkbookModel.fromSnapshot({ ...snapshot, sheets });
}

/**
 * 一张比 `mergedFixture` 宽得多的空表，专供 Finding 1（final-review-fix）的
 * outside-block 场景使用——那个场景需要 `start=8, count=2` 这样的下标，
 * `mergedFixture` 的 7 行 5 列坐标空间放不下。`setMerges` 是唯一的公开写口，
 * 不必经 `fromSnapshot`。
 */
function wideMergedFixture(merges: CellRange[]): SheetModel {
  const model = SheetModel.createEmpty({ blankRows: 12, blankColumns: 12 });
  model.setMerges(merges);
  return model;
}

/** 抓住抛出的错误本身——要比对类型与消息，`toThrow` 的正则做不到。 */
function grabError(run: () => void): Error | undefined {
  try {
    run();
  } catch (error) {
    return error as Error;
  }
  return undefined;
}

/**
 * 断言 prepare 期与 commit 期对同一个非法输入抛出**同一个**错误。
 *
 * 两边都要先确认真的抛了：都不抛的话 `fromPrepare?.constructor` 与
 * `fromCommit?.constructor` 双双是 `undefined`，`toBe` 会空过。
 */
function expectPrepareMatchesCommit(mutation: SheetMutation, commit: (model: SheetModel) => void): void {
  const fromPrepare = grabError(() => invertSheetMutation(fixture(), mutation));
  const fromCommit = grabError(() => {
    commit(fixture());
  });
  expect(fromCommit).toBeInstanceOf(Error);
  expect(fromPrepare).toBeInstanceOf(Error);
  expect(fromPrepare?.constructor).toBe(fromCommit?.constructor);
  expect(fromPrepare?.message).toBe(fromCommit?.message);
}

/**
 * 求逆不得改动模型。对照物同样换成 `WorkbookModel.toSnapshot()`——理由见
 * `assertOracleSeesPluginTables`：拿 `SheetModel.toSnapshot()` 比的话，"求逆的
 * 过程写坏了插件表"这件事它也看不见。今天没有代码会写它，所以这道守卫是空的；
 * 但**一个空的守卫与一个有效的守卫长得一模一样**，这正是它要换的理由。
 */
function expectNonInvasive(mutation: SheetMutation): void {
  const workbook = workbookFixture();
  const before = workbook.toSnapshot();
  assertOracleSeesPluginTables(before);
  invertSheetMutation(workbook.sheet(), mutation);
  expect(workbook.toSnapshot()).toStrictEqual(before);
}

describe("invertSheetMutation — 精确性", () => {
  it("格子：set-cell 与 set-range 往返后逐字段相等", () => {
    expectExactRoundTrip({ type: "set-cell", row: 0, column: 0, record: { v: "changed", t: CellValueType.NUMBER } });
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

  it("set-range 的 records 不齐时，逆仍按 range 的尺寸铺满", () => {
    // 只给了 1×1 的 records，其余五格按空格处理（施加下去是删格）。逆若按 records
    // 的尺寸算，那五格就永远回不来了——这条正是为盯住那个尺寸来源而存在。
    expectExactRoundTrip({
      type: "set-range",
      range: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 2 },
      records: [[{ v: 1 }]]
    });
  });
});

/**
 * 这四条只断言**形状**，不做模型往返。
 *
 * 它们对应的四个分支是对正变更自身字段的纯算术，一个字都不读模型，所以任何"往
 * 施加/回滚里塞个错误再看测试红不红"的探针都碰不到它们——四个分支同时写错，
 * 精确性与非侵入两组测试可以全绿。四条分开写而不是并成一条，是为了红的时候
 * 失败信息直接点名是哪一条臂断了。
 */
describe("invertSheetMutation — 纯算术分支的形状", () => {
  it("insert-rows 的逆是删掉同样多的行", () => {
    expect(
      invertSheetMutation(fixture(), {
        type: "insert-rows",
        start: 1,
        rows: [{ 0: { v: "x" } }, {}]
      })
    ).toStrictEqual({ type: "delete-rows", start: 1, count: 2 });
  });

  it("restore-rows 的逆是删掉同样多的行", () => {
    expect(
      invertSheetMutation(fixture(), {
        type: "restore-rows",
        start: 2,
        rows: [{ 1: { v: "y" } }, {}],
        meta: [{ h: 30 }, undefined],
        tables: {},
        mergesBefore: [], validationsBefore: []
      })
    ).toStrictEqual({ type: "delete-rows", start: 2, count: 2 });
  });

  it("insert-columns 的逆是删掉同样多的列", () => {
    expect(invertSheetMutation(fixture(), { type: "insert-columns", start: 1, count: 2 })).toStrictEqual({
      type: "delete-columns",
      start: 1,
      count: 2
    });
  });

  it("restore-columns 的逆是删掉同样多的列", () => {
    expect(
      invertSheetMutation(fixture(), {
        type: "restore-columns",
        start: 1,
        count: 3,
        cells: [{ 0: { v: "z" } }, {}, {}],
        meta: [{ w: 80 }, undefined, undefined],
        tables: {},
        mergesBefore: [], validationsBefore: []
      })
    ).toStrictEqual({ type: "delete-columns", start: 1, count: 3 });
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

  it("块紧贴轴尾时，逆的 target 正好等于轴长——validateAxisChange 只挡 > 轴长", () => {
    const model = fixture();
    const axisLength = model.rowCount();
    expect(axisLength).toBe(7);
    // start + count === 轴长：块就在轴的最末尾，把它前移，逆要把它送回最末尾，
    // 于是逆的 target === start + count === 轴长。差一格就会被 validateAxisChange
    // 挡掉（它抛的条件是 target > axisLength），所以这条边界值得钉住。
    expect(invertSheetMutation(model, { type: "move-rows", start: 5, count: 2, target: 0 })).toStrictEqual({
      type: "move-rows",
      start: 0,
      count: 2,
      target: axisLength
    });
    expectExactRoundTrip({ type: "move-rows", start: 5, count: 2, target: 0 });
  });
});

describe("invertSheetMutation — 参数校验前置到 prepare", () => {
  it("行：target 落在块内的 move，在求逆时就抛，不拖到 commit", () => {
    // 不拦的话这里会返回一条形状合法、语义是垃圾的逆，等 SheetModel.moveRows 在
    // commit 期间才抛——而那时垃圾逆已经进了撤销记录，且 commit 本该不会失败
    expect(() => invertSheetMutation(fixture(), { type: "move-rows", start: 1, count: 2, target: 2 })).toThrow(
      /move target cannot be inside the moved block/
    );
  });

  it("列：target 落在块内的 move，同样在求逆时就抛", () => {
    expect(() => invertSheetMutation(fixture(), { type: "move-columns", start: 0, count: 2, target: 1 })).toThrow(
      /move target cannot be inside the moved block/
    );
  });

  /*
   * 下面四条各盯一条臂。
   *
   * 它们此前是空白：先前那两条校验测试用的是"target 落在块内"，而 target 只有
   * `move` 有——于是 insert / delete 四条臂拿到了校验却没有任何测试。把
   * `validateAxisChange` 从这四条臂上整个删掉，1508 条测试全绿。
   *
   * 后果不是崩溃，是更坏的一种：`delete-rows(100, 5)` 在只有 7 行的表上不会炸，
   * `captureRowCells` 与 `getRowMeta` 对越界下标平静地给出 `{}` 与 `undefined`，
   * 于是求逆返回一条**形状合法、语义是垃圾**的 `restore-rows`，把失败推到 commit。
   * 四条分开写，红的时候直接点名是哪条臂。
   */
  it("insert-rows：start 越过轴长，在求逆时就抛，且与 commit 期是同一个错", () => {
    // 轴长 7（占用到第 4 行 + blankRows 2）
    expectPrepareMatchesCommit({ type: "insert-rows", start: 8, rows: [{}] }, (m) => {
      m.insertRows(8, 1);
    });
  });

  it("delete-rows：start + count 越过轴长，在求逆时就抛，且与 commit 期是同一个错", () => {
    expectPrepareMatchesCommit({ type: "delete-rows", start: 5, count: 3 }, (m) => {
      m.deleteRows(5, 3);
    });
  });

  it("insert-columns：start 越过轴长，在求逆时就抛，且与 commit 期是同一个错", () => {
    // 列轴长 5（占用到第 2 列 + blankColumns 2）
    expectPrepareMatchesCommit({ type: "insert-columns", start: 6, count: 1 }, (m) => {
      m.insertColumns(6, 1);
    });
  });

  it("delete-columns：start + count 越过轴长，在求逆时就抛，且与 commit 期是同一个错", () => {
    expectPrepareMatchesCommit({ type: "delete-columns", start: 3, count: 3 }, (m) => {
      m.deleteColumns(3, 3);
    });
  });

  it("reorder：不合法的 order 在求逆时就抛，且与 commit 期抛的完全是同一个错", () => {
    // 断言的不只是"抛了"，而是 prepare 与 commit 对"什么是合法 order"给出**同一个
    // 答案**——这两处一旦分叉，撤销记录里就会躺着一条 commit 期才会炸的变更。
    const orders = [
      Int32Array.from([0, 0, 1, 2, 3]), // 长度够，但不是排列（0 出现两次）
      Int32Array.from([2, 0, 1]), // 是排列，但盖不住占用范围（含只有元数据的第 4 行）
      // 两种都不合法。**这一条才钉得住两个检查的先后**：前两条各只触发一个检查，
      // 把顺序调过来它们仍然全绿；这一条在"先 extent"下抛 CyExcelError、在
      // "先排列"下抛 RangeError，顺序一反就与 commit 期对不上。
      Int32Array.from([0, 0, 5])
    ];
    for (const order of orders) {
      const fromPrepare = grabError(() => invertSheetMutation(fixture(), { type: "reorder-rows", order }));
      const fromCommit = grabError(() => {
        fixture().reorderRows(order);
      });
      expect(fromCommit).toBeInstanceOf(Error);
      expect(fromPrepare?.constructor).toBe(fromCommit?.constructor);
      expect(fromPrepare?.message).toBe(fromCommit?.message);
    }
    // 合法的 order 一如既往地过
    expect(() =>
      invertSheetMutation(fixture(), { type: "reorder-rows", order: Int32Array.from([2, 0, 1, 4, 3]) })
    ).not.toThrow();
  });

  /**
   * 任务四之后，九个结构操作里**只剩 `reorder-rows`** 在存在合并区时无条件拒绝
   * （`SheetModel.#assertNoMerges` 仍然只服务它）。另外八个已经改用 `merges.ts`
   * 的 `relocateMergesFor*` 正确搬移合并区——插入/删除/`restore-*` 永远能搬，
   * 两个 `move` 成员只在跨越块边界时才拒绝，见下面"合并区"那一组用例。
   */
  it("reorder-rows：存在合并区时仍然无条件拒绝，且与 commit 期抛的是同一个错", () => {
    const order = Int32Array.from([2, 0, 1, 4, 3]);
    const fromPrepare = grabError(() => invertSheetMutation(mergedFixture(), { type: "reorder-rows", order }));
    const fromCommit = grabError(() => {
      mergedFixture().reorderRows(order);
    });
    expect(fromCommit).toBeInstanceOf(CyExcelError);
    expect(fromPrepare?.constructor).toBe(fromCommit?.constructor);
    expect(fromPrepare?.message).toBe(fromCommit?.message);

    // 格子变更不受合并区阻拦——commit 期的 setRecord / deleteCell 也不查
    expect(() =>
      invertSheetMutation(mergedFixture(), { type: "set-cell", row: 0, column: 0, record: { v: 1 } })
    ).not.toThrow();
  });

  it("先后顺序与 commit 期一致：move 先校验参数、reorder 先校验排列，合并区都排在后面", () => {
    // move-rows：target 落在块内是参数错，与合并区无关，必须先报这条
    const paramsFirst = grabError(() => invertSheetMutation(mergedFixture(), { type: "move-rows", start: 1, count: 2, target: 2 }));
    const paramsFirstCommit = grabError(() => mergedFixture().moveRows(1, 2, 2));
    expect(paramsFirst).toBeInstanceOf(RangeError);
    expect(paramsFirst?.message).toBe(paramsFirstCommit?.message);

    // reorder ①：order 盖不住占用范围，同时又有合并区——报的必须是 extent 那条
    const extentFirst = grabError(() =>
      invertSheetMutation(mergedFixture(), { type: "reorder-rows", order: Int32Array.from([0, 0, 5]) })
    );
    const extentFirstCommit = grabError(() => mergedFixture().reorderRows(Int32Array.from([0, 0, 5])));
    expect(extentFirst).toBeInstanceOf(CyExcelError);
    expect(extentFirst?.message).toMatch(/would be silently dropped/);
    expect(extentFirst?.message).toBe(extentFirstCommit?.message);

    // reorder ②：长度够但不是排列，同时又有合并区——报的必须是排列那条
    const permutationFirst = grabError(() =>
      invertSheetMutation(mergedFixture(), { type: "reorder-rows", order: Int32Array.from([0, 0, 1, 2, 3]) })
    );
    const permutationFirstCommit = grabError(() =>
      mergedFixture().reorderRows(Int32Array.from([0, 0, 1, 2, 3]))
    );
    expect(permutationFirst).toBeInstanceOf(RangeError);
    expect(permutationFirst?.message).toMatch(/not a permutation/);
    expect(permutationFirst?.message).toBe(permutationFirstCommit?.message);
  });
});

/**
 * 决策 4：插入/删除/`restore-*`（八个非 `reorder-rows` 的结构成员里的六个）永远
 * 能搬，不再因为合并区存在就拒绝——`relocateMergesForInsert` / `relocateMergesForDelete`
 * 是 `(现有合并区, 操作参数)` 的纯函数，不会失败。`mergedFixture` 的合并区是
 * `[0,0]`-`[1,1]`（rows 0..1 × columns 0..1）。
 */
describe("invertSheetMutation — 插入/删除/restore 不再因合并区被拒绝，而是正确搬移", () => {
  it("insert-rows：插入点落在合并区内部，合并区被撑大，往返后逐字段相等", () => {
    const workbook = mergedWorkbookFixture();
    const model = workbook.sheet();
    const before = workbook.toSnapshot();
    const mutation: SheetMutation = { type: "insert-rows", start: 1, rows: [{}] };
    const inverse = invertSheetMutation(model, mutation);
    apply(model, mutation);
    // 插入点 1 严格落在 [0,1] 内部——relocateIntervalForInsert 的撑大分支
    expect(model.merges()).toStrictEqual([{ startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 }]);
    apply(model, inverse);
    expect(workbook.toSnapshot()).toStrictEqual(before);
  });

  it("delete-rows：删除区间部分落在合并区内，合并区收缩，往返后逐字段相等", () => {
    const workbook = mergedWorkbookFixture();
    const model = workbook.sheet();
    const before = workbook.toSnapshot();
    const mutation: SheetMutation = { type: "delete-rows", start: 1, count: 1 };
    const inverse = invertSheetMutation(model, mutation);
    apply(model, mutation);
    // 删掉合并区的下沿那一行——relocateIntervalForDelete 的"部分重叠"分支，收缩为 [0,0]
    expect(model.merges()).toStrictEqual([{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }]);
    apply(model, inverse);
    expect(workbook.toSnapshot()).toStrictEqual(before);
  });

  it("insert-columns / delete-columns：列轴同理，不再因合并区被拒绝", () => {
    const model = mergedFixture();
    expect(() => invertSheetMutation(model, { type: "insert-columns", start: 1, count: 1 })).not.toThrow();
    expect(() => invertSheetMutation(model, { type: "delete-columns", start: 1, count: 1 })).not.toThrow();
  });

  it("restore-rows / restore-columns：走的也是 insertRows / insertColumns，同样不再被拒绝", () => {
    const model = mergedFixture();
    expect(() =>
      invertSheetMutation(model, {
        type: "restore-rows",
        start: 1,
        rows: [{}],
        meta: [undefined],
        tables: {},
        mergesBefore: [], validationsBefore: []
      })
    ).not.toThrow();
    expect(() =>
      invertSheetMutation(model, {
        type: "restore-columns",
        start: 1,
        count: 1,
        cells: [{}],
        meta: [undefined],
        tables: {},
        mergesBefore: [], validationsBefore: []
      })
    ).not.toThrow();
  });
});

/**
 * 决策 4：两个 `move` 成员可能因为合并区**跨越移动块边界**而失败——合并区既不
 * 完全落在块内、也不完全落在块外，`relocateMergesForMove` 返回 `ok: false`。
 * `assertMergesDoNotStraddleMove` 与 `assertNoMerges`（存在即拒）不同，只在这
 * 一种情形下才抛，且错误码与消息与 `SheetModel.moveRows` / `moveColumns` 在
 * commit 期抛出的逐字一致。
 */
describe("invertSheetMutation — move 合并区跨越移动块边界时抛 MergeSpansMoveBoundary", () => {
  it("move-rows：合并区 [0,1] 与移动块 [1,2] 只有第 1 行重叠，跨越边界", () => {
    const mutation: SheetMutation = { type: "move-rows", start: 1, count: 2, target: 5 };
    const fromPrepare = grabError(() => invertSheetMutation(mergedFixture(), mutation));
    const fromCommit = grabError(() => mergedFixture().moveRows(1, 2, 5));
    expect(fromCommit).toBeInstanceOf(CyExcelError);
    expect((fromCommit as CyExcelError).code).toBe(CyExcelErrorCode.MergeSpansMoveBoundary);
    expect(fromPrepare?.constructor).toBe(fromCommit?.constructor);
    expect(fromPrepare?.message).toBe(fromCommit?.message);
  });

  it("move-columns：合并区 [0,1] 与移动块 [1,2] 只有第 1 列重叠，跨越边界", () => {
    const mutation: SheetMutation = { type: "move-columns", start: 1, count: 2, target: 4 };
    const fromPrepare = grabError(() => invertSheetMutation(mergedFixture(), mutation));
    const fromCommit = grabError(() => mergedFixture().moveColumns(1, 2, 4));
    expect(fromCommit).toBeInstanceOf(CyExcelError);
    expect((fromCommit as CyExcelError).code).toBe(CyExcelErrorCode.MergeSpansMoveBoundary);
    expect(fromPrepare?.constructor).toBe(fromCommit?.constructor);
    expect(fromPrepare?.message).toBe(fromCommit?.message);
  });

  it("move-rows：合并区完全落在移动块内，跟着块走，不抛错", () => {
    // 块 [0,2] 完全覆盖合并区 [0,1]——containedInBlock 分支，不是跨越边界
    expect(() =>
      invertSheetMutation(mergedFixture(), { type: "move-rows", start: 0, count: 3, target: 6 })
    ).not.toThrow();
  });

  it("move-rows：合并区完全落在移动块外，也不抛错", () => {
    // 合并区 [0,1] 与块 [3,4] 不相交——outsideBlock 分支
    expect(() =>
      invertSheetMutation(mergedFixture(), { type: "move-rows", start: 3, count: 2, target: 0 })
    ).not.toThrow();
  });

  /**
   * final-review Finding 1：合并区**完全在移动块外**，但块摘除之后的插入点
   * 严格落在这条合并区内部——`relocateIntervalForInsert` 的撑大分支会把块
   * "塞进"这条合并区，逆向搬回去必然撞上"跨越移动块边界"，undo 因此永久卡死
   * 撤销栈（见 `merges.ts` 的 `relocateMergesForMove` 文档注释）。修复后，这一步
   * 在**正向** `SheetModel.moveRows`/`moveColumns` 就直接拒绝，不必等到 undo
   * 才炸——`invert.ts` 的 `assertMergesDoNotStraddleMove` 与 `SheetModel` 的
   * commit 期检查读的是同一个 `relocateMergesForMove`，两边必须给出同一个答案。
   */
  it("move-rows：outside-block 合并区、重新插入点严格落在其内部——正向 moveRows 直接拒绝，不是撑大（Finding 1）", () => {
    // merge rows [2,6]；block rows [8,9]（start=8,count=2），target=4 =>
    // insertAt=4，严格落在 [2,6] 内部——修复前会把它错误撑大成 [2,8]。
    const merges: CellRange[] = [{ startRow: 2, startColumn: 0, endRow: 6, endColumn: 0 }];
    const mutation: SheetMutation = { type: "move-rows", start: 8, count: 2, target: 4 };
    const fromPrepare = grabError(() => invertSheetMutation(wideMergedFixture(merges), mutation));
    const fromCommit = grabError(() => wideMergedFixture(merges).moveRows(8, 2, 4));
    expect(fromCommit).toBeInstanceOf(CyExcelError);
    expect((fromCommit as CyExcelError).code).toBe(CyExcelErrorCode.MergeSpansMoveBoundary);
    expect(fromPrepare?.constructor).toBe(fromCommit?.constructor);
    expect(fromPrepare?.message).toBe(fromCommit?.message);
  });

  it("move-columns：outside-block 合并区、重新插入点严格落在其内部——正向 moveColumns 直接拒绝（Finding 1，列轴同构）", () => {
    const merges: CellRange[] = [{ startRow: 0, startColumn: 2, endRow: 0, endColumn: 6 }];
    const mutation: SheetMutation = { type: "move-columns", start: 8, count: 2, target: 4 };
    const fromPrepare = grabError(() => invertSheetMutation(wideMergedFixture(merges), mutation));
    const fromCommit = grabError(() => wideMergedFixture(merges).moveColumns(8, 2, 4));
    expect(fromCommit).toBeInstanceOf(CyExcelError);
    expect((fromCommit as CyExcelError).code).toBe(CyExcelErrorCode.MergeSpansMoveBoundary);
    expect(fromPrepare?.constructor).toBe(fromCommit?.constructor);
    expect(fromPrepare?.message).toBe(fromCommit?.message);
  });

  /**
   * final-review Finding 1 补的往返覆盖：此前本文件的 move+merge 用例只断言
   * `invertSheetMutation` 本身"不抛错"（`.not.toThrow()`），从没有真的施加正向
   * mutation、施加算出的逆、再比对合并区数组——这正是 bug 能躲过 T1-T10 十次
   * 单任务评审的原因。这里补一条**真正的往返**：outside-block 合并区在一次仍然
   * 允许成功的移动下（插入点落在摘除后区间的起点，触发的是"平移"支而不是
   * "撑大"支——见 `merges.test.ts` 的对照用例），施加正向、施加逆，`merges()`
   * 精确回到搬移前的状态。
   */
  it("move-rows：outside-block 合并区在允许成功的搬移下，往返后合并区数组精确复原（Finding 1 round-trip）", () => {
    // merge rows [4,6]；block rows [8,9]，target=4 => insertAt=4，
    // 4 <= afterDeleteStart(4)：平移支，不是撑大——见 merges.test.ts 同构用例。
    const originalMerges: CellRange[] = [{ startRow: 4, startColumn: 0, endRow: 6, endColumn: 0 }];
    const model = wideMergedFixture(originalMerges);
    const mutation: SheetMutation = { type: "move-rows", start: 8, count: 2, target: 4 };
    const inverse = invertSheetMutation(model, mutation);
    apply(model, mutation);
    // 正向确实搬动了合并区（平移 +2，不是原地不动）——不然下面的复原断言会空过。
    expect(model.merges()).toStrictEqual([{ startRow: 6, startColumn: 0, endRow: 8, endColumn: 0 }]);
    apply(model, inverse);
    expect(model.merges()).toStrictEqual(originalMerges);
  });

  it("move-columns：outside-block 合并区在允许成功的搬移下，往返后合并区数组精确复原（Finding 1 round-trip，列轴同构）", () => {
    const originalMerges: CellRange[] = [{ startRow: 0, startColumn: 4, endRow: 0, endColumn: 6 }];
    const model = wideMergedFixture(originalMerges);
    const mutation: SheetMutation = { type: "move-columns", start: 8, count: 2, target: 4 };
    const inverse = invertSheetMutation(model, mutation);
    apply(model, mutation);
    expect(model.merges()).toStrictEqual([{ startRow: 0, startColumn: 6, endRow: 0, endColumn: 8 }]);
    apply(model, inverse);
    expect(model.merges()).toStrictEqual(originalMerges);
  });
});

/**
 * 决策 3：`delete-rows` / `delete-columns` 的求逆必须捕获 `mergesBefore`，理由是
 * "全覆盖消失"这个信息丢失场景——`insert` 无从凭空造出一条被删掉的合并区。这里的
 * 反例是决策 3 文本里手算的那个场景的一个具体实例：删除范围**全覆盖**合并区，
 * 合并区从数组里整条消失，撤销必须能把它精确复原，而不是靠"重新插入"能凑出来的。
 */
describe("invertSheetMutation — delete-rows/delete-columns 全覆盖合并区的撤销（决策 3）", () => {
  it("delete-rows：删除范围全覆盖合并区，合并区消失，撤销精确复原", () => {
    const workbook = mergedWorkbookFixture();
    const model = workbook.sheet();
    const before = workbook.toSnapshot();
    const mutation: SheetMutation = { type: "delete-rows", start: 0, count: 2 };
    const inverse = invertSheetMutation(model, mutation);
    // 求逆必须捕获施加删除之前的整份合并区数组
    expect((inverse as { mergesBefore: readonly unknown[] }).mergesBefore).toStrictEqual([
      { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 }
    ]);
    apply(model, mutation);
    // 全覆盖：delStart(0) <= mStart(0) && delEnd(1) >= mEnd(1)，合并区整条消失
    expect(model.merges()).toStrictEqual([]);
    apply(model, inverse);
    expect(model.merges()).toStrictEqual([{ startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 }]);
    expect(workbook.toSnapshot()).toStrictEqual(before);
  });

  it("delete-columns：同理，列轴的合并区全覆盖消失后也能精确复原", () => {
    const workbook = mergedWorkbookFixture();
    const model = workbook.sheet();
    const before = workbook.toSnapshot();
    const mutation: SheetMutation = { type: "delete-columns", start: 0, count: 2 };
    const inverse = invertSheetMutation(model, mutation);
    expect((inverse as { mergesBefore: readonly unknown[] }).mergesBefore).toStrictEqual([
      { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 }
    ]);
    apply(model, mutation);
    expect(model.merges()).toStrictEqual([]);
    apply(model, inverse);
    expect(model.merges()).toStrictEqual([{ startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 }]);
    expect(workbook.toSnapshot()).toStrictEqual(before);
  });
});

/**
 * 决策 1：`set-merges` 是整体替换的哑变更，求逆读的是施加前的整份 `merges()`——
 * 与 `set-row-meta` 的 `meta: model.getRowMeta(…)` 同一个模式。
 */
describe("invertSheetMutation — set-merges", () => {
  it("往返后逐字段相等", () => {
    expectExactRoundTrip({
      type: "set-merges",
      merges: [{ startRow: 2, startColumn: 2, endRow: 3, endColumn: 3 }]
    });
  });

  it("逆是施加前的那份合并区数组，不是施加后的", () => {
    const model = mergedFixture();
    const inverse = invertSheetMutation(model, {
      type: "set-merges",
      merges: [{ startRow: 3, startColumn: 3, endRow: 4, endColumn: 4 }]
    });
    expect(inverse).toStrictEqual({
      type: "set-merges",
      merges: [{ startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 }]
    });
  });

  it("求逆不改模型", () => {
    expectNonInvasive({ type: "set-merges", merges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 }] });
  });
});

describe("invertSheetMutation — 结构删除的逆带回插件侧表的条目", () => {
  it("删行：被删区间里的插件行表条目跟着逆回来", () => {
    // 区间 [1, 3) 内有 rowKeys 的第 1 条，区间外有第 4 条：前者靠 restore-rows 的
    // tables 带回，后者靠 insertRows 自己的重键移回，两条路都要走到。
    expectExactRoundTrip(
      { type: "delete-rows", start: 1, count: 2 },
      { precondition: (model) => assertBandHasPluginEntries(model, "row", 1, 2) }
    );
  });

  it("删列：被删区间里的插件列表条目跟着逆回来", () => {
    // 列轴同理。ColumnCells 的键是行号而 tables 的下标是**列号**，两者不是一回事
    expectExactRoundTrip(
      { type: "delete-columns", start: 0, count: 2 },
      { precondition: (model) => assertBandHasPluginEntries(model, "column", 0, 2) }
    );
  });

  it("多张插件行表：条目各回各的表，不串位", () => {
    const seed = (workbook: WorkbookModel): void => {
      // 第二张行表，条目落在第 2 行——与 rowKeys 的第 1 条同在区间 [1, 3) 内，
      // 于是"按表名装、按表名写回"一旦写成"按顺序装、按顺序写回"就会串位
      workbook.addResource({ name: "com.acme.rowFlags", indexedBy: "row", data: { 2: "f2" } });
    };
    expectExactRoundTrip(
      { type: "delete-rows", start: 1, count: 2 },
      { seed, precondition: (model) => assertBandHasPluginEntries(model, "row", 1, 2) }
    );

    // 下面这两条直接读表，**不只是**红灯时好定位的诊断——它们独立地扛着一段覆盖：
    // 把 `writeTables` 里"跳过 undefined"那条稀疏性守卫拆掉，若同时有人把上面两条
    // 往返断言从 `toStrictEqual` 降成 `toEqual`（那时 `{ 2: undefined }` 与"没有第 2
    // 个键"就再也分不开），全场只剩这里还红。所以它不是可以顺手删掉的重复设置。
    const workbook = workbookFixture();
    seed(workbook);
    const model = workbook.sheet();
    const mutation: SheetMutation = { type: "delete-rows", start: 1, count: 2 };
    const inverse = invertSheetMutation(model, mutation);
    apply(model, mutation);
    apply(model, inverse);
    expect(model.table("com.acme.rowKeys")?.read()).toStrictEqual({ 1: "k1", 4: "k4" });
    expect(model.table("com.acme.rowFlags")?.read()).toStrictEqual({ 2: "f2" });
  });

  it("表名叫 __proto__ 的插件表，条目照样带得回来", () => {
    // `"__proto__"` 是一条能一路走到 `capturePluginTables` 的**合法**表名：登记那层
    // （`assertValidResourceEntry`）只要求非空字符串，底下的登记处是 `Map`，对它没有
    // 任何特殊待遇。而往一个**对象字面量**上写这个键不会建出自有属性，它会打到
    // `Object.prototype.__proto__` 的 setter：键凭空消失，这张表的条目被整段静默丢掉
    // ——比任务三那道守卫还差，那道守卫至少会抛。
    const seed = (workbook: WorkbookModel): void => {
      workbook.addResource({ name: "__proto__", indexedBy: "row", data: { 1: "p1", 4: "p4" } });
    };
    expectExactRoundTrip(
      { type: "delete-rows", start: 1, count: 2 },
      { seed, precondition: (model) => assertBandHasPluginEntries(model, "row", 1, 2) }
    );

    const workbook = workbookFixture();
    seed(workbook);
    const model = workbook.sheet();
    const tables = (
      invertSheetMutation(model, { type: "delete-rows", start: 1, count: 2 }) as {
        tables: Record<string, unknown>;
      }
    ).tables;

    // 这里必须走 `getOwnPropertyDescriptor`，**不能**写 `tables["__proto__"]`：真出这个
    // bug 时，那个读法拿回来的正是被顶上去的原型（也就是那个数组本身），断言会照样
    // 变绿——一条读法就能让整条用例空过。
    expect(Object.keys(tables)).toContain("__proto__");
    expect(Object.getOwnPropertyDescriptor(tables, "__proto__")?.value).toStrictEqual(["p1", undefined]);
    // 原型被换成数组的话，这条逆变更会带着一个 `length` 属性躺进历史栈
    expect(Object.getPrototypeOf(tables)).toBeNull();
  });
});

/**
 * 任务三在 `restore-*` 还没有装侧表条目的字段时，让 `delete-rows` / `delete-columns`
 * 的求逆**拒绝**（`MutationNotSupported`），理由是"静默错搬比拒绝糟得多"。那是对的
 * 临时处置，但它同时让 spec §7.2 承诺的"删行"在有插件资源时不可用——到接线那一刻
 * 就是功能倒退。
 *
 * **一道临时守卫如果没有测试盯着它被撤除，它会一直活下去**，而后来的人只会看到
 * "有插件资源就不能删行"，以为那是设计。这一条就是那双眼睛。
 */
describe("invertSheetMutation — 任务三那道临时守卫已撤除", () => {
  it("注册了插件侧表时，删除的逆不再拒绝，而是把条目装进 restore-* 带走", () => {
    const model = fixture();

    // 只断言"不抛"是不够的：一条悄悄丢掉条目的实现也不抛。所以连形状一起钉住。
    expect(invertSheetMutation(model, { type: "delete-rows", start: 1, count: 2 })).toStrictEqual({
      type: "restore-rows",
      start: 1,
      rows: [{ 0: { v: "a1" }, 2: { v: "c1", s: "s1" } }, { 1: { v: "b2", f: "=1+1" } }],
      meta: [{ h: 40 }, undefined],
      tables: sideTables({ "com.acme.rowKeys": ["k1", undefined] }),
      mergesBefore: [], validationsBefore: []
    });

    expect(invertSheetMutation(model, { type: "delete-columns", start: 0, count: 1 })).toStrictEqual({
      type: "restore-columns",
      start: 0,
      count: 1,
      cells: [{ 0: { v: "a0" }, 1: { v: "a1" }, 3: { v: "a3" } }],
      meta: [{ w: 100 }],
      tables: sideTables({ "com.acme.colTags": ["t0"] }),
      mergesBefore: [], validationsBefore: []
    });
  });
});

/**
 * `capturePluginTables` 的两条过滤规则——**按轴过滤**、**跳过核心表**——各有一条用例
 * 盯着，而且盯着它们的**只有**下面这两条。
 *
 * **为什么不能靠往返用例覆盖。** 这不是"还没写"，是**量过之后确认盖不住**：把任何一
 * 条过滤删掉，六条往返用例**零条**变红。机制是这样的——行轴的变更根本不移动列表的
 * 下标，于是一条被错抓的列表条目会被原样写回原下标，那是一次**精确的空操作**；
 * `rowData` 同理，它那条条目 `writeRows` 已经按同一个值恢复过一遍了。捕获侧与写回侧
 * 是**一起错的**，而"施加逆之后快照相等"这个判据在结构上就看不见一起错的两边。
 *
 * 所以：**将来重构时，不要因为"往返用例已经覆盖了"而删掉这两条。** 往返用例没有覆盖，
 * 这句话是量出来的，不是猜的。这两条一旦没了，`capturePluginTables` 有一半函数体就
 * 完全没有测试盯着，而且不会有任何一条用例变红来提醒。
 */
describe("invertSheetMutation — capturePluginTables 的两条过滤规则", () => {
  it("轴向过滤：行删除的逆里只装行表，列表不进来（反向亦然）", () => {
    const model = fixture();

    const onRowDelete = tableNamesOf(invertSheetMutation(model, { type: "delete-rows", start: 1, count: 2 }));
    // 正反两面都要断言：只断言"列表不在"的话，一个什么都不装的实现照样过
    expect(onRowDelete).toContain("com.acme.rowKeys");
    expect(onRowDelete).not.toContain("com.acme.colTags");

    const onColumnDelete = tableNamesOf(
      invertSheetMutation(model, { type: "delete-columns", start: 0, count: 1 })
    );
    expect(onColumnDelete).toContain("com.acme.colTags");
    expect(onColumnDelete).not.toContain("com.acme.rowKeys");
  });

  it("核心表排除：rowData / columnData 一个都不进 tables", () => {
    const model = fixture();

    // 两张核心表是被**两套**机制挡住的：同轴那张（行删除时的 rowData）靠按名字跳过，
    // 跨轴那张（行删除时的 columnData）靠轴向过滤。所以两个名字都要点到。
    const onRowDelete = tableNamesOf(invertSheetMutation(model, { type: "delete-rows", start: 1, count: 2 }));
    expect(onRowDelete).toContain("com.acme.rowKeys");
    expect(onRowDelete).not.toContain("rowData");
    expect(onRowDelete).not.toContain("columnData");

    const onColumnDelete = tableNamesOf(
      invertSheetMutation(model, { type: "delete-columns", start: 0, count: 1 })
    );
    expect(onColumnDelete).toContain("com.acme.colTags");
    expect(onColumnDelete).not.toContain("columnData");
    expect(onColumnDelete).not.toContain("rowData");
  });
});

/**
 * `set-row-meta` / `set-column-meta` 是 4B 加的两个成员，理由不是补齐对称性：行高
 * 权威从 `LayoutEngine` 迁进 `rowData` 之后（spec §5.1），改行高就是在改 `SheetModel`
 * 内部的状态；这条路径若不走 `SheetMutation`，"同一个事务里既改行高又做结构变更"
 * 就成了两个参与者写同一份状态，行高那一半绕过影子模型直接写活模型。
 *
 * 这里钉的是它们的逆，重点在 `undefined` 那一支：**原先没有条目时，逆必须是删条目，
 * 不是写一个空的 `{}`**。写空对象不会报错——它只是把这一行留在 `rowExtent()` 的并集
 * 里，坐标空间与快照各多出一格。对照物是 `WorkbookModel.toSnapshot()` 的逐字段比较，
 * 它看得见这一格。
 */
describe("invertSheetMutation — 行/列元数据", () => {
  it("set-row-meta 覆盖一条已有的元数据：往返后逐字段相等", () => {
    expectExactRoundTrip(
      { type: "set-row-meta", row: 1, meta: { h: 99, hp: "manual" } },
      {
        // 挡掉"覆盖的是一条本来就不存在的元数据"那种空过
        precondition: (model) => {
          expect(model.getRowMeta(1)).toStrictEqual({ h: 40 });
        }
      }
    );
  });

  it("set-row-meta 写到一个原先没有条目的行：逆携带 undefined，施加下去把条目删掉", () => {
    const model = fixture();
    expect(model.getRowMeta(0)).toBeUndefined();
    const mutation: SheetMutation = { type: "set-row-meta", row: 0, meta: { h: 99 } };
    const inverse = invertSheetMutation(model, mutation);
    expect(inverse).toStrictEqual({ type: "set-row-meta", row: 0, meta: undefined });
    apply(model, mutation);
    expect(model.getRowMeta(0)).toStrictEqual({ h: 99 });
    apply(model, inverse);
    // `toBeUndefined` 分不清"删掉了"和"存了一个空对象"，所以再钉一次快照
    expect(model.getRowMeta(0)).toBeUndefined();
    expect(model.toSnapshot().rowData).toStrictEqual(fixture().toSnapshot().rowData);
  });

  it("set-row-meta 写到没有条目的行，整表往返后逐字段相等", () => {
    expectExactRoundTrip({ type: "set-row-meta", row: 0, meta: { h: 99, hd: true } });
  });

  it("set-column-meta 的两支同理", () => {
    expectExactRoundTrip(
      { type: "set-column-meta", column: 2, meta: { w: 7 } },
      {
        precondition: (model) => {
          expect(model.getColumnMeta(2)).toStrictEqual({ w: 200, t: "number" });
        }
      }
    );
    expectExactRoundTrip({ type: "set-column-meta", column: 1, meta: { w: 7, hd: true } });
  });

  it("元数据变更不查合并区——`setRowMeta` / `setColumnMeta` 在 commit 期也不查", () => {
    const model = mergedFixture();
    expect(() => invertSheetMutation(model, { type: "set-row-meta", row: 0, meta: { h: 9 } })).not.toThrow();
    expect(() =>
      invertSheetMutation(model, { type: "set-column-meta", column: 0, meta: { w: 9 } })
    ).not.toThrow();
  });

  it("求逆不改模型", () => {
    expectNonInvasive({ type: "set-row-meta", row: 1, meta: { h: 99 } });
    expectNonInvasive({ type: "set-column-meta", column: 0, meta: { w: 9 } });
  });
});

/**
 * `restore-*` 的 `tables` 三条不变量。4A 把它们记成"4B 的前置条件，不是可选项"，当时
 * 的理由是"没有外部生产者"；4B 补上不是因为出现了外部生产者，而是因为**历史栈**：
 * 一条逆变更会在栈里存活很久，其间插件完全可以注销一张侧表。
 *
 * 三条的共同点是**失败方式都是静默的**，所以每条都要有一条自己的用例——尤其是第一条，
 * 它破坏的是被恢复区间之外、压根没被碰过的数据。
 */
describe("invertSheetMutation — restore-* 的 tables 校验", () => {
  it("条目数与被恢复的行数不符：拒绝，而不是让它写出界", () => {
    const model = fixture();
    const mutation: SheetMutation = {
      type: "restore-rows",
      start: 1,
      rows: [{}, {}],
      meta: [undefined, undefined],
      // 两行的 restore 带了三条条目，第三条会落到 insert 往下推的那一行上
      tables: sideTables({ "com.acme.rowKeys": ["a", "b", "c"] }),
      mergesBefore: [], validationsBefore: []
    };
    const error = grabError(() => invertSheetMutation(model, mutation));
    expect(error).toBeInstanceOf(CyExcelError);
    expect((error as CyExcelError).code).toBe("invalid-mutation");
    expect(error?.message).toContain("expected 2");
  });

  it("点名一张没登记的表：拒绝", () => {
    const error = grabError(() =>
      invertSheetMutation(fixture(), {
        type: "restore-rows",
        start: 1,
        rows: [{}],
        meta: [undefined],
        tables: sideTables({ "com.acme.nope": ["a"] }),
        mergesBefore: [], validationsBefore: []
      })
    );
    expect(error).toBeInstanceOf(CyExcelError);
    expect((error as CyExcelError).code).toBe("invalid-mutation");
  });

  it("表的轴与被恢复的轴不符：拒绝——写进去不会报错，只会静默错位", () => {
    const error = grabError(() =>
      invertSheetMutation(fixture(), {
        type: "restore-rows",
        start: 1,
        rows: [{}],
        meta: [undefined],
        // 这是一张列表，拿它的条目去写行下标
        tables: sideTables({ "com.acme.colTags": ["a"] }),
        mergesBefore: [], validationsBefore: []
      })
    );
    expect(error).toBeInstanceOf(CyExcelError);
    expect((error as CyExcelError).code).toBe("invalid-mutation");
    expect(error?.message).toContain("indexed by column");
  });

  it("求逆自己产出的 restore-* 一定过得了这三条", () => {
    const model = fixture();
    const inverse = invertSheetMutation(model, { type: "delete-rows", start: 1, count: 2 });
    // 反空洞：这条逆确实带着一张插件行表，不然下面那句是在给一条空 `tables` 背书
    expect(tableNamesOf(inverse)).toContain("com.acme.rowKeys");
    expect(() => invertSheetMutation(model, inverse)).not.toThrow();
  });
});
