import { CyExcelError, CyExcelErrorCode } from "../errors";
import { assertValidRowOrder } from "../layout/rowOrder";
import { validateAxisChange } from "../layout/structuralChange";
import {
  MutationCoordinator,
  type CommitReceipt,
  type Mutation,
  type MutationPrepareContext,
  type MutationPrepareSession,
  type PreparedMutationParticipant
} from "../mutations";
import { rowValuesToCells, valueToRecord } from "../sheet/import";
import type { SheetMutation } from "../sheet/mutation";
import { changesForSheetMutation, type SheetAxisSpace } from "../sheet/patches";
import type { SheetModel } from "../sheet/SheetModel";
import { createSheetPlanningResources } from "../sheet/planning";
import type { CellRange, Patch } from "../types";
import { assertValidCoordinate, assertValidRange } from "../validation";
import type { ChangeSet } from "./changes";

export interface TableStoreCommitResult {
  readonly patches: readonly Patch[];
  readonly changeSet: ChangeSet;
  readonly version: number;
}

class TableSheetMutation implements Mutation {
  readonly type: string;
  readonly changes: readonly Patch[];
  readonly #sheet: SheetModel;
  readonly #mutation: SheetMutation;

  constructor(sheet: SheetModel, mutation: SheetMutation, changes: readonly Patch[]) {
    this.#sheet = sheet;
    this.#mutation = mutation;
    this.type = `data:${mutation.type}`;
    this.changes = Object.freeze([...changes]);
  }

  prepare(context: MutationPrepareContext): PreparedMutationParticipant {
    // 施加**之前**的坐标空间要先读下来：`prepareData` 返回时预览已经吃过这条变更了。
    const before = spaceOf(context);
    assertValueWriteIsInsideSpace(this.#mutation, before);
    const prepared = context.prepareData(this.#sheet, this.#mutation);
    const space = unionSpace(before, spaceOf(context));

    const forward = new TableSheetMutation(
      this.#sheet,
      prepared.forward,
      changesForSheetMutation(prepared.forward, space)
    );
    const inverse = new TableSheetMutation(
      this.#sheet,
      prepared.inverse,
      changesForSheetMutation(prepared.inverse, space)
    );
    return {
      hasEffect: true,
      forward,
      inverse,
      estimatedBytes: prepared.estimatedBytes,
      commit: () => prepared.commit(),
      rollback: () => prepared.rollback(),
      dispose: () => prepared.dispose()
    };
  }
}

function spaceOf(context: MutationPrepareContext): SheetAxisSpace {
  return { rowCount: context.model.getRowCount(), columnCount: context.model.getColumnCount() };
}

/**
 * 值写不许改坐标空间——越界的 `set-cell` / `set-range` 在这里被挡下。
 *
 * 旧模型里挡它的是行数组本身（`rows[5]` 不存在，`ArrayDataSource.setCell()` 抛
 * `invalid-coordinate`）。稀疏矩阵不会拒绝：它会把第 5 行建出来，`rowCount()`
 * 随之变大，而这条变更只带一条 `cell-value` 补丁——渲染层不知道表变长了，选择、
 * 布局、行头三处各自算出不同的行数。轴只能由结构变更改动，这条是那个约束的闸。
 *
 * 校验发生在 prepare 而不是构造时：同一个事务里"先插行、再往新行里写"是合法的
 * （`rowCommands.test.ts` 的暂存视图用例），而那时新行只存在于影子模型上。
 * `context.model` 正是影子，`before` 是这条变更施加前的暂存空间。
 */
function assertValueWriteIsInsideSpace(mutation: SheetMutation, space: SheetAxisSpace): void {
  if (mutation.type === "set-cell") {
    assertInsideSpace(mutation.row, mutation.column, space);
    return;
  }
  if (mutation.type === "set-range") {
    assertInsideSpace(mutation.range.startRow, mutation.range.startColumn, space);
    assertInsideSpace(mutation.range.endRow, mutation.range.endColumn, space);
  }
}

function assertInsideSpace(row: number, column: number, space: SheetAxisSpace): void {
  if (row >= space.rowCount) {
    throw new CyExcelError(CyExcelErrorCode.InvalidCoordinate, `Row ${row} does not exist`);
  }
  if (column >= space.columnCount) {
    throw new CyExcelError(CyExcelErrorCode.ColumnNotFound, `Column ${column} does not exist`);
  }
}

/**
 * 正逆两条补丁共用一个坐标空间：**施加前与施加后的并集**。
 *
 * `row-layout` 那条补丁的 `range` 是"这一段几何失效了"的依据，取窄了是渲染层**画少
 * 了**、不报错。正变更该取施加后的空间（与 `RowHeightMutation` / `LayoutRootMutation`
 * 同口径），而逆是在回滚时施加的，它要盖住的是回滚**之后**的空间——那正好是施加
 * 前的那个。两者取并集，两条都不会偏小；偏大的那一头只是多重画一段。
 *
 * 这不是假想的：`set-row-meta` 会改坐标空间（`rowData` 的下标进 `rowCount()` 的并集），
 * 于是"删掉第 5000 行的行高"这条变更施加后空间缩到 10 行，它的逆再把那条元数据放
 * 回去——只按施加后的空间算，逆那条补丁的 `range` 会停在第 9 行。
 */
function unionSpace(before: SheetAxisSpace, after: SheetAxisSpace): SheetAxisSpace {
  return {
    rowCount: Math.max(before.rowCount, after.rowCount),
    columnCount: Math.max(before.columnCount, after.columnCount)
  };
}

/** Mutable table model whose writes are finalized by one shared coordinator. */
export class TableStore {
  readonly #sheet: SheetModel;
  readonly #coordinator: MutationCoordinator;
  #nextDirectTransactionId = 1;

  constructor(sheet: SheetModel, coordinator = new MutationCoordinator()) {
    this.#sheet = sheet;
    this.#coordinator = coordinator;
  }

  get version(): number {
    return this.#coordinator.revisionGate.revision;
  }

  get coordinator(): MutationCoordinator {
    return this.#coordinator;
  }

  /** 这次事务写的是哪张表——`CyExcelCore` 要拿它去读快照与侧表。 */
  get sheet(): SheetModel {
    return this.#sheet;
  }

  beginMutationSession(transactionId: string): MutationPrepareSession {
    return this.#coordinator.begin(
      transactionId,
      (baseRevision) => createSheetPlanningResources(this.#sheet, baseRevision)
    );
  }

  getRowCount(): number {
    return this.#sheet.rowCount();
  }

  getColumnCount(): number {
    return this.#sheet.columnCount();
  }

  getCell(row: number, column: number): unknown {
    this.#assertCellExists(row, column);
    return this.#sheet.getCell(row, column);
  }

  occupiedCells(): Generator<{ row: number; column: number; value: unknown }> {
    return this.#sheet.occupiedCells();
  }

  /**
   * 只带值的一条变更**整条替换**记录，不与该格已有的 `t`/`s`/`f`/`si` 合并。
   *
   * 今天这样是对的：4B 之后没有任何活路径产出带这些字段的记录（样式属阶段 5、公式
   * 不在范围内），所以此刻做合并等于给一条还不存在的状态写代码。
   *
   * **等到有了那些字段，合并必须发生在事务里、读影子的 `getRecord`**——不能在这里
   * 读活模型：同一个事务里连写两次同一格时，第二条会拿到一份还没被第一条更新过的
   * 旧样式，把第一条的结果悄悄盖掉。届时的落点是给 `TableSheetMutation.prepare` 一条
   * "带值的 set-cell"变体，在 `prepareData` 之前就地问影子要记录。
   */
  createSetCellMutation(row: number, column: number, value: unknown): Mutation {
    // 这里只查形状；"在不在坐标空间里"要等到 prepare 才知道——同一个事务里可以
    // 先插行再往新行写，那时新行只存在于影子模型上（见 `assertValueWriteIsInsideSpace`）。
    assertValidCoordinate(row, column);
    return new TableSheetMutation(
      this.#sheet,
      { type: "set-cell", row, column, record: valueToRecord(value) },
      [{ type: "cell-value", row, column }]
    );
  }

  createSetRangeMutation(range: CellRange, values: readonly (readonly unknown[])[]): Mutation {
    assertValidRange(range);
    const expectedRows = range.endRow - range.startRow + 1;
    const expectedColumns = range.endColumn - range.startColumn + 1;
    if (values.length !== expectedRows || values.some((row) => row.length !== expectedColumns)) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidRange,
        `Range values must have shape ${expectedRows}x${expectedColumns}`
      );
    }
    return new TableSheetMutation(
      this.#sheet,
      { type: "set-range", range, records: values.map((row) => row.map(valueToRecord)) },
      [{ type: "range-value", range }]
    );
  }

  /**
   * 整体替换合并区表。不经 `changesForSheetMutation`——`merge-layout` 是一个不带
   * 字段的整表重绘标记，不依赖坐标空间，与 `set-range`/`set-cell` 走同一类字面量
   * 补丁数组的判断（`changesForSheetMutation` 是给需要按 `space` 算范围的补丁用的）。
   */
  createSetMergesMutation(merges: readonly CellRange[]): Mutation {
    return new TableSheetMutation(
      this.#sheet,
      { type: "set-merges", merges: merges.map((range) => ({ ...range })) },
      [{ type: "merge-layout" }]
    );
  }

  createInsertRowsMutation(start: number, rows: readonly (readonly unknown[])[]): Mutation {
    this.#assertStructuralShape("row", "insert", start, rows.length);
    return new TableSheetMutation(
      this.#sheet,
      { type: "insert-rows", start, rows: rows.map(rowValuesToCells) },
      [{ type: "row-structure", operation: "insert", start, count: rows.length }]
    );
  }

  createDeleteRowsMutation(start: number, count: number): Mutation {
    this.#assertStructuralShape("row", "delete", start, count);
    return new TableSheetMutation(
      this.#sheet,
      { type: "delete-rows", start, count },
      [{ type: "row-structure", operation: "delete", start, count }]
    );
  }

  createMoveRowsMutation(start: number, count: number, target: number): Mutation {
    this.#assertStructuralShape("row", "move", start, count, target);
    return new TableSheetMutation(
      this.#sheet,
      { type: "move-rows", start, count, target },
      [{ type: "row-structure", operation: "move", start, count, target }]
    );
  }

  /**
   * Whole-axis permutation (Step 22 sorting). Unlike the other row mutations
   * there is no start/count/target shape to validate against the row axis,
   * so this skips `#assertStructuralShape` and instead validates `order` is
   * a permutation of the current row count directly.
   */
  createReorderRowsMutation(order: Int32Array): Mutation {
    assertValidRowOrder(order, this.getRowCount());
    return new TableSheetMutation(
      this.#sheet,
      { type: "reorder-rows", order },
      [{ type: "row-reorder", order }]
    );
  }

  /**
   * 插列只要个数。列不再有身份（ADR §3.8.2），所以这里不再收字段映射——
   * "第 i 列对应哪个字段"是导入那一刻的事实，不是表运行期的状态。
   */
  createInsertColumnsMutation(start: number, count: number): Mutation {
    this.#assertStructuralShape("column", "insert", start, count);
    return new TableSheetMutation(
      this.#sheet,
      { type: "insert-columns", start, count },
      [{ type: "column-structure", operation: "insert", start, count }]
    );
  }

  createDeleteColumnsMutation(start: number, count: number): Mutation {
    this.#assertStructuralShape("column", "delete", start, count);
    return new TableSheetMutation(
      this.#sheet,
      { type: "delete-columns", start, count },
      [{ type: "column-structure", operation: "delete", start, count }]
    );
  }

  createMoveColumnsMutation(start: number, count: number, target: number): Mutation {
    this.#assertStructuralShape("column", "move", start, count, target);
    return new TableSheetMutation(
      this.#sheet,
      { type: "move-columns", start, count, target },
      [{ type: "column-structure", operation: "move", start, count, target }]
    );
  }

  setCell(row: number, column: number, value: unknown): TableStoreCommitResult {
    return this.#execute(this.createSetCellMutation(row, column, value));
  }

  setRange(range: CellRange, values: readonly (readonly unknown[])[]): TableStoreCommitResult {
    return this.#execute(this.createSetRangeMutation(range, values));
  }

  updateCell(row: number, column: number, updater: (current: unknown) => unknown): TableStoreCommitResult {
    return this.setCell(row, column, updater(this.getCell(row, column)));
  }

  insertRows(start: number, rows: readonly (readonly unknown[])[]): TableStoreCommitResult {
    return this.#execute(this.createInsertRowsMutation(start, rows));
  }

  deleteRows(start: number, count: number): TableStoreCommitResult {
    return this.#execute(this.createDeleteRowsMutation(start, count));
  }

  moveRows(start: number, count: number, target: number): TableStoreCommitResult {
    return this.#execute(this.createMoveRowsMutation(start, count, target));
  }

  reorderRows(order: Int32Array): TableStoreCommitResult {
    return this.#execute(this.createReorderRowsMutation(order));
  }

  insertColumns(start: number, count: number): TableStoreCommitResult {
    return this.#execute(this.createInsertColumnsMutation(start, count));
  }

  deleteColumns(start: number, count: number): TableStoreCommitResult {
    return this.#execute(this.createDeleteColumnsMutation(start, count));
  }

  moveColumns(start: number, count: number, target: number): TableStoreCommitResult {
    return this.#execute(this.createMoveColumnsMutation(start, count, target));
  }

  #execute(mutation: Mutation): TableStoreCommitResult {
    const session = this.beginMutationSession(`store-${this.#nextDirectTransactionId++}`);
    let receipt: CommitReceipt;
    try {
      session.stage(mutation);
      const prepared = session.seal();
      try {
        receipt = this.#coordinator.commit(prepared);
      } catch (error) {
        this.#coordinator.discard(prepared);
        throw error;
      }
    } catch (error) {
      session.discard();
      throw error;
    }
    return this.#toCommitResult(receipt);
  }

  #toCommitResult(receipt: CommitReceipt): TableStoreCommitResult {
    return { patches: receipt.patches, changeSet: receipt.changeSet, version: receipt.afterRevision };
  }

  #assertCellExists(row: number, column: number): void {
    assertValidCoordinate(row, column);
    if (row >= this.getRowCount()) {
      throw new CyExcelError(CyExcelErrorCode.InvalidCoordinate, `Row ${row} does not exist`);
    }
    if (column >= this.getColumnCount()) {
      throw new CyExcelError(CyExcelErrorCode.ColumnNotFound, `Column ${column} does not exist`);
    }
  }

  #assertStructuralShape(
    axis: "row" | "column",
    operation: "insert" | "delete" | "move",
    start: number,
    count: number,
    target?: number
  ): void {
    const minimumLength = Math.max(start + (operation === "insert" ? 0 : count), target ?? 0);
    validateAxisChange({ axis, operation, start, count, ...(target === undefined ? {} : { target }) }, minimumLength);
  }
}
