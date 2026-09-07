import { describe, expect, it } from "vitest";
import { CyExcelError, CyExcelErrorCode } from "../../../../src/core/errors";
import { RevisionGate } from "../../../../src/core/model";
import { applySheetMutation } from "../../../../src/core/sheet/apply";
import type { SheetMutation } from "../../../../src/core/sheet/mutation";
import { SheetTransaction } from "../../../../src/core/sheet/SheetTransaction";
import type { PreparedSheetMutation } from "../../../../src/core/sheet/SheetTransaction";
import { WorkbookModel } from "../../../../src/core/sheet/WorkbookModel";

/**
 * 一张够小、但每一类状态都占着的表：格子、行元数据、列元数据、两张插件侧表。第 4 行
 * **只有元数据、没有格子**，与 `invert.test.ts` / `clone.test.ts` 是同一个刻意安排。
 *
 * 每个格子的值都不一样，这样"预览前进了没有"才能逐格断言，而不是只能数个数。
 */
function fixture(): WorkbookModel {
  const workbook = WorkbookModel.createEmpty({ sheet: { blankRows: 2, blankColumns: 2 } });
  const model = workbook.sheet();
  model.setRecord(0, 0, { v: "a0" });
  model.setRecord(0, 1, { v: "b0" });
  model.setRecord(1, 0, { v: "a1" });
  model.setRecord(2, 0, { v: "a2" });
  model.setRecord(3, 0, { v: "a3" });
  model.setRowMeta(1, { h: 40 });
  model.setRowMeta(4, { hd: true });
  model.setColumnMeta(0, { w: 100 });
  workbook.addResource({ name: "com.acme.rowKeys", indexedBy: "row", data: { 1: "k1", 3: "k3" } });
  workbook.addResource({ name: "com.acme.colTags", indexedBy: "column", data: { 0: "t0" } });
  return workbook;
}

function revision() {
  return new RevisionGate().revision;
}

function transactionOn(workbook: WorkbookModel): SheetTransaction {
  return new SheetTransaction(workbook.sheet(), revision());
}

function grabError(run: () => void): Error | undefined {
  try {
    run();
  } catch (error) {
    return error as Error;
  }
  return undefined;
}

/** 一个只交换头两行、其余原样的排列——够简单，能逐格验。 */
function swapFirstTwo(length: number): Int32Array {
  const order = new Int32Array(length);
  for (let index = 0; index < length; index += 1) {
    order[index] = index;
  }
  order[0] = 1;
  order[1] = 0;
  return order;
}

describe("SheetTransaction — prepare 不得触碰活状态", () => {
  it("暂存一条格子变更，活模型一个字节都不动", () => {
    const workbook = fixture();
    const before = workbook.toSnapshot();
    const transaction = transactionOn(workbook);
    transaction.prepareMutation({ type: "set-cell", row: 0, column: 0, record: { v: "x" } });
    expect(workbook.toSnapshot()).toStrictEqual(before);
  });

  it("暂存一条结构变更，活模型一个字节都不动", () => {
    const workbook = fixture();
    const before = workbook.toSnapshot();
    const transaction = transactionOn(workbook);
    transaction.prepareMutation({ type: "delete-rows", start: 1, count: 2 });
    transaction.prepareMutation({ type: "insert-columns", start: 0, count: 1 });
    expect(workbook.toSnapshot()).toStrictEqual(before);
  });

  /**
   * 校验也在这一步发生：`invertSheetMutation` 头一件事就是 `assertMutationIsAppliable`。
   * 这正是 `MutationCoordinator` 要的"能查出来的都在 prepare 查出来"——而它抛出的时候，
   * 活模型同样必须原封不动。
   */
  it("参数非法时在 prepare 期抛，且抛之前没写过活模型", () => {
    const workbook = fixture();
    const before = workbook.toSnapshot();
    const transaction = transactionOn(workbook);
    expect(() => transaction.prepareMutation({ type: "delete-rows", start: 99, count: 1 })).toThrow(
      RangeError
    );
    expect(workbook.toSnapshot()).toStrictEqual(before);
  });

  it("`restore-*` 点名一张没登记的侧表，同样在 prepare 期被拦下", () => {
    // 这条是 T1 那三条 `tables` 不变量在事务这条路上的出口：一条逆变更会在历史栈里
    // 存活很久，其间插件完全可以注销一张表，等到撤销重放时表就不在了。
    const workbook = fixture();
    const before = workbook.toSnapshot();
    const transaction = transactionOn(workbook);
    expect(() =>
      transaction.prepareMutation({
        type: "restore-rows",
        start: 0,
        rows: [{}],
        meta: [undefined],
        tables: { "com.acme.notRegistered": [undefined] },
        mergesBefore: []
      })
    ).toThrow(CyExcelError);
    expect(workbook.toSnapshot()).toStrictEqual(before);
  });
});

describe("SheetTransaction — 预览随每一条暂存前进", () => {
  it("格子变更后读到的是新值，活模型读到的还是旧值", () => {
    const workbook = fixture();
    const transaction = transactionOn(workbook);
    transaction.prepareMutation({ type: "set-cell", row: 0, column: 0, record: { v: "x" } });
    expect(transaction.getCell(0, 0)).toBe("x");
    expect(workbook.sheet().getCell(0, 0)).toBe("a0");
  });

  it("结构变更后坐标空间与格子位置一起前进", () => {
    const workbook = fixture();
    const transaction = transactionOn(workbook);
    const rowsBefore = transaction.getRowCount();
    transaction.prepareMutation({ type: "insert-rows", start: 1, rows: [{}] });
    expect(transaction.getRowCount()).toBe(rowsBefore + 1);
    expect(transaction.getCell(2, 0)).toBe("a1");
    expect(transaction.getCell(1, 0)).toBeUndefined();
  });

  it("`getRecord` 给的是完整记录，不是裸值——变更层要靠它保住 `t`/`s`/`f`", () => {
    const workbook = fixture();
    const transaction = transactionOn(workbook);
    transaction.prepareMutation({
      type: "set-cell",
      row: 0,
      column: 0,
      record: { v: 1, t: "number", s: "s1" }
    });
    expect(transaction.getRecord(0, 0)).toStrictEqual({ v: 1, t: "number", s: "s1" });
    expect(transaction.getCell(0, 0)).toBe(1);
  });

  /**
   * 组合的**要害**：第 2 条的逆必须在"第 1 条已生效"的状态上求。
   *
   * 逐条按相反顺序回滚时，撤掉第 2 条之后应当停在第 1 条的结果（`"x"`）上。若逆是
   * 对着**活模型**求的，那一条逆会是"恢复成 `a0`"——两条测试断言里只有这一条能把它
   * 认出来，最终状态是分辨不出来的（两种实现都回到 `a0`）。
   */
  it("第二条的逆求在第一条之上，不是求在活模型之上", () => {
    const workbook = fixture();
    const transaction = transactionOn(workbook);
    const first = transaction.prepareMutation({
      type: "set-cell",
      row: 0,
      column: 0,
      record: { v: "x" }
    });
    const second = transaction.prepareMutation({
      type: "set-cell",
      row: 0,
      column: 0,
      record: { v: "y" }
    });
    first.commit();
    second.commit();
    expect(workbook.sheet().getCell(0, 0)).toBe("y");

    second.rollback();
    expect(workbook.sheet().getCell(0, 0)).toBe("x");
    first.rollback();
    expect(workbook.sheet().getCell(0, 0)).toBe("a0");
  });
});

/**
 * 设计稿 §7 点名的那条组合序列。
 *
 * 比较一律走 `WorkbookModel.toSnapshot()`——`SheetModel.toSnapshot()` **不**序列化插件
 * 侧表（4A 已经踩过这个坑），拿它比会让侧表上的错漏静默通过。
 */
describe("SheetTransaction — 组合序列 setCell → insertRows → setCell → deleteRows → reorderRows", () => {
  function stageSequence(transaction: SheetTransaction): {
    mutations: SheetMutation[];
    prepared: PreparedSheetMutation[];
  } {
    const mutations: SheetMutation[] = [];
    const prepared: PreparedSheetMutation[] = [];
    const stage = (mutation: SheetMutation) => {
      mutations.push(mutation);
      prepared.push(transaction.prepareMutation(mutation));
    };

    stage({ type: "set-cell", row: 0, column: 0, record: { v: "x" } });
    expect(transaction.getCell(0, 0)).toBe("x");

    stage({ type: "insert-rows", start: 1, rows: [{}] });
    expect(transaction.getCell(2, 0)).toBe("a1");

    stage({ type: "set-cell", row: 1, column: 0, record: { v: "inserted" } });
    expect(transaction.getCell(1, 0)).toBe("inserted");

    // 删掉的是原来的第 2 行（插入把它推到了第 3 行）
    stage({ type: "delete-rows", start: 3, count: 1 });
    expect(transaction.getCell(3, 0)).toBe("a3");

    // `order` 的长度取坐标空间：它恒 >= 占用范围（`rowCount = maxOccupied + 1 +
    // blankRows`），所以 `reorderRows` 那条"排列必须盖住每一个占用行"的检查过得去。
    stage({ type: "reorder-rows", order: swapFirstTwo(transaction.getRowCount()) });
    expect(transaction.getCell(0, 0)).toBe("inserted");
    expect(transaction.getCell(1, 0)).toBe("x");

    return { mutations, prepared };
  }

  it("逐条 commit 之后，活模型等于直接施加这五条的结果", () => {
    // 判据取一份**独立算出来的**期望：同样五条变更直接施加在另一张同样的表上。
    // 拿事务自己的预览当判据会变成自证。
    const workbook = fixture();
    const transaction = transactionOn(workbook);
    const { mutations, prepared } = stageSequence(transaction);

    for (const participant of prepared) {
      participant.commit();
    }

    const oracle = fixture();
    for (const mutation of mutations) {
      applySheetMutation(oracle.sheet(), mutation);
    }
    expect(workbook.toSnapshot()).toStrictEqual(oracle.toSnapshot());
  });

  it("按相反顺序回滚之后，逐字段回到基态（含插件侧表）", () => {
    const workbook = fixture();
    const base = workbook.toSnapshot();
    const transaction = transactionOn(workbook);
    const { prepared } = stageSequence(transaction);

    for (const participant of prepared) {
      participant.commit();
    }
    expect(workbook.toSnapshot()).not.toStrictEqual(base);

    for (let index = prepared.length - 1; index >= 0; index -= 1) {
      (prepared[index] as PreparedSheetMutation).rollback();
    }
    expect(workbook.toSnapshot()).toStrictEqual(base);
  });

  /**
   * `MutationCoordinator.commit` 的失败路径：第 k 个参与者抛出时，它把**已经开始的
   * 那几个**按相反顺序回滚，后面没开始的不碰。这里模拟前三条 commit 成功、第四条
   * 之前就失败。
   */
  it("只 commit 了前几条时，回滚这几条也能回到基态", () => {
    const workbook = fixture();
    const base = workbook.toSnapshot();
    const transaction = transactionOn(workbook);
    const { prepared } = stageSequence(transaction);

    const started = prepared.slice(0, 3);
    for (const participant of started) {
      participant.commit();
    }
    for (let index = started.length - 1; index >= 0; index -= 1) {
      (started[index] as PreparedSheetMutation).rollback();
    }
    expect(workbook.toSnapshot()).toStrictEqual(base);
    // 没开始的那两条被 rollback 到也必须是无害的（协调器会对全部 started 调用，
    // 而"started"的判定在它那边，这里保守地全调一遍）
    for (let index = prepared.length - 1; index >= 0; index -= 1) {
      (prepared[index] as PreparedSheetMutation).rollback();
    }
    expect(workbook.toSnapshot()).toStrictEqual(base);
  });
});

describe("SheetTransaction — 回滚的两条护栏", () => {
  it("没 commit 过的 rollback 是空操作", () => {
    const workbook = fixture();
    const before = workbook.toSnapshot();
    const transaction = transactionOn(workbook);
    const prepared = transaction.prepareMutation({ type: "delete-rows", start: 0, count: 2 });
    prepared.rollback();
    expect(workbook.toSnapshot()).toStrictEqual(before);
  });

  it("rollback 幂等——第二次不再施加一遍逆", () => {
    // `PreparedMutationParticipant` 的原话："rollback() must be idempotent"。这里
    // 挑的是 `delete-rows`：它的逆是 `restore-rows`，施加两遍会**插两次行**，快照
    // 立刻能看出来。
    const workbook = fixture();
    const before = workbook.toSnapshot();
    const transaction = transactionOn(workbook);
    const prepared = transaction.prepareMutation({ type: "delete-rows", start: 0, count: 2 });
    prepared.commit();
    prepared.rollback();
    prepared.rollback();
    expect(workbook.toSnapshot()).toStrictEqual(before);
  });
});

describe("SheetTransaction — 生命周期与估值", () => {
  it("dispose 之后再读或再暂存都抛 MutationConflict", () => {
    const workbook = fixture();
    const transaction = transactionOn(workbook);
    transaction.dispose();
    for (const read of [
      () => transaction.getCell(0, 0),
      () => transaction.getRecord(0, 0),
      () => transaction.prepareMutation({ type: "set-cell", row: 0, column: 0, record: undefined })
    ]) {
      const error = grabError(read);
      expect(error).toBeInstanceOf(CyExcelError);
      expect((error as CyExcelError).code).toBe(CyExcelErrorCode.MutationConflict);
    }
  });

  it("dispose 不销毁已经交出去的逆变更——它别名的值归活模型所有", () => {
    // 历史栈里的记录要在事务结束之后长期有效，这条钉住 `dispose()` 只是关句柄。
    const workbook = fixture();
    const transaction = transactionOn(workbook);
    const prepared = transaction.prepareMutation({ type: "delete-rows", start: 1, count: 1 });
    prepared.commit();
    transaction.dispose();
    prepared.rollback();
    expect(workbook.sheet().getCell(1, 0)).toBe("a1");
    expect(workbook.sheet().table("com.acme.rowKeys")?.get(1)).toBe("k1");
  });

  it("估值算的是正逆两条之和，且随被撤销的内容增长", () => {
    // 真正占历史栈的是**逆**：`delete-rows` 自己只有几个下标，它的
    // `restore-rows` 却背着整段记录、元数据与插件侧表条目。所以只估正变更会把
    // 一条极贵的删除算成一条极便宜的变更。
    const workbook = fixture();
    const transaction = transactionOn(workbook);
    const heavy = transaction.prepareMutation({ type: "delete-rows", start: 0, count: 4 });

    // 同样是一条 `delete-rows`，只是删在空白区（第 5 行既没有格子也没有任何侧表
    // 条目）：正变更一模一样，逆却是空的。两者的差值全部来自逆。
    const empty = fixture();
    const emptyTransaction = transactionOn(empty);
    const light = emptyTransaction.prepareMutation({ type: "delete-rows", start: 5, count: 1 });

    expect(Number.isFinite(heavy.estimatedBytes)).toBe(true);
    expect(heavy.estimatedBytes).toBeGreaterThan(light.estimatedBytes);
  });

  it("交出去的 `forward` 就是传进来的那一条，`inverse` 能把它抵消", () => {
    const workbook = fixture();
    const transaction = transactionOn(workbook);
    const mutation: SheetMutation = { type: "delete-rows", start: 1, count: 2 };
    const prepared = transaction.prepareMutation(mutation);
    expect(prepared.forward).toBe(mutation);
    expect(prepared.inverse.type).toBe("restore-rows");
  });
});
