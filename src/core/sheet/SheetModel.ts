import type { CellRange } from "../types";
import { CyExcelError, CyExcelErrorCode } from "../errors";
import { CellMatrix } from "./CellMatrix";
import { assertValidBlankCount, resolveCoordinateSpace } from "./coordinateSpace";
import { deserializeSheet, serializeSheet } from "./serialize";
import type { CellRecord, ColumnMeta, RowMeta, SheetSnapshot } from "./types";
import { DEFAULT_BLANK_COLUMNS, DEFAULT_BLANK_ROWS } from "./types";
import {
  buildDeleteRowsPlan,
  buildInsertRowsPlan,
  buildMoveRowsPlan
} from "./rowPlan";
import { remapByColumnShift, remapByRowPlan } from "./indexShift";
import {
  buildDeleteColumnsShift,
  buildInsertColumnsShift,
  buildMoveColumnsShift
} from "./columnShift";
import { validateAxisChange } from "../layout/structuralChange";

export interface SheetModelOptions {
  readonly id?: string;
  readonly name?: string;
  readonly blankRows?: number;
  readonly blankColumns?: number;
  readonly defaultRowHeight?: number;
  readonly defaultColumnWidth?: number;
  readonly freeze?: { readonly rows: number; readonly columns: number };
}

/**
 * 一张表的模型：稀疏记录矩阵 + 行列元数据 + 由 `blank*` 派生的坐标空间。
 *
 * 本阶段不接线到 `TableStore` / `LayoutEngine` / `CyExcelCore`——那是阶段 3。
 */
export class SheetModel {
  readonly #matrix: CellMatrix;
  #rowData: Readonly<Record<number, RowMeta>>;
  #columnData: Readonly<Record<number, ColumnMeta>>;
  #mergeData: readonly CellRange[];
  readonly #id: string;
  readonly #name: string;
  readonly #blankRows: number;
  readonly #blankColumns: number;
  readonly #defaultRowHeight: number;
  readonly #defaultColumnWidth: number;
  readonly #freeze: { readonly rows: number; readonly columns: number } | undefined;

  private constructor(init: {
    matrix: CellMatrix;
    rowData: Readonly<Record<number, RowMeta>>;
    columnData: Readonly<Record<number, ColumnMeta>>;
    mergeData: readonly CellRange[];
    id: string;
    name: string;
    blankRows: number;
    blankColumns: number;
    defaultRowHeight: number;
    defaultColumnWidth: number;
    freeze?: { readonly rows: number; readonly columns: number };
  }) {
    assertValidBlankCount(init.blankRows, "blankRows");
    assertValidBlankCount(init.blankColumns, "blankColumns");
    this.#matrix = init.matrix;
    this.#rowData = init.rowData;
    this.#columnData = init.columnData;
    this.#mergeData = init.mergeData;
    this.#id = init.id;
    this.#name = init.name;
    this.#blankRows = init.blankRows;
    this.#blankColumns = init.blankColumns;
    this.#defaultRowHeight = init.defaultRowHeight;
    this.#defaultColumnWidth = init.defaultColumnWidth;
    this.#freeze = init.freeze ?? undefined;
  }

  static createEmpty(options: SheetModelOptions = {}): SheetModel {
    return new SheetModel({
      matrix: new CellMatrix(),
      rowData: {},
      columnData: {},
      mergeData: [],
      id: options.id ?? "sheet-1",
      name: options.name ?? "Sheet1",
      blankRows: options.blankRows ?? DEFAULT_BLANK_ROWS,
      blankColumns: options.blankColumns ?? DEFAULT_BLANK_COLUMNS,
      defaultRowHeight: options.defaultRowHeight ?? 28,
      defaultColumnWidth: options.defaultColumnWidth ?? 120,
      ...(options.freeze !== undefined && { freeze: options.freeze })
    });
  }

  static fromSnapshot(snapshot: SheetSnapshot): SheetModel {
    const restored = deserializeSheet(snapshot);
    return new SheetModel({
      matrix: restored.matrix,
      // deserializeSheet 已经拷贝过 rowData/columnData/mergeData，这里不需要再拷一层
      rowData: restored.rowData,
      columnData: restored.columnData,
      mergeData: restored.mergeData,
      id: restored.id,
      name: restored.name,
      blankRows: restored.blankRows,
      blankColumns: restored.blankColumns,
      defaultRowHeight: restored.defaultRowHeight,
      defaultColumnWidth: restored.defaultColumnWidth,
      ...(restored.freeze !== undefined ? { freeze: restored.freeze } : {})
    });
  }

  /** 裸值，不是记录（ADR 0004 §3.6）。 */
  getCell(row: number, column: number): unknown {
    return this.#matrix.get(row, column)?.v;
  }

  getRecord(row: number, column: number): CellRecord | undefined {
    return this.#matrix.get(row, column);
  }

  /** 只改值，保留该格已有的 `t`/`s`/`f`/`si`。 */
  setCell(row: number, column: number, value: unknown): void {
    const existing = this.#matrix.get(row, column);
    this.#matrix.set(row, column, existing ? { ...existing, v: value } : { v: value });
  }

  setRecord(row: number, column: number, record: CellRecord): void {
    this.#matrix.set(row, column, record);
  }

  deleteCell(row: number, column: number): void {
    this.#matrix.delete(row, column);
  }

  /** 坐标空间上界（spec §3.3）。 */
  rowCount(): number {
    return resolveCoordinateSpace(this.#matrix.maxOccupiedRow(), this.#blankRows);
  }

  columnCount(): number {
    return resolveCoordinateSpace(this.#matrix.maxOccupiedColumn(), this.#blankColumns);
  }

  /** 最大占用行号加一——与 `rowCount()` 是两个不同的数（spec §3.3 后果 ④）。 */
  occupiedRowCount(): number {
    const max = this.#matrix.maxOccupiedRow();
    return max === undefined ? 0 : max + 1;
  }

  toSnapshot(): SheetSnapshot {
    return serializeSheet({
      id: this.#id,
      name: this.#name,
      blankRows: this.#blankRows,
      blankColumns: this.#blankColumns,
      defaultRowHeight: this.#defaultRowHeight,
      defaultColumnWidth: this.#defaultColumnWidth,
      matrix: this.#matrix,
      rowData: this.#rowData,
      columnData: this.#columnData,
      mergeData: this.#mergeData,
      ...(this.#freeze !== undefined ? { freeze: this.#freeze } : {})
    });
  }

  /**
   * 合并区跟着结构操作正确搬移是一整块 Excel 语义（跨插入点延展、被删范围
   * 完全覆盖则消失、部分覆盖则收缩），属于阶段 5。在那之前**拒绝**而不是
   * 错着搬——静默错搬比拒绝糟得多。仓库先例：headerGroups 同样挡住列结构变更。
   */
  #assertNoMerges(operation: string): void {
    if (this.#mergeData.length > 0) {
      throw new CyExcelError(
        CyExcelErrorCode.MutationNotSupported,
        `${operation} is not supported while merged ranges exist`
      );
    }
  }

  /**
   * 行的"占用范围"必须取**所有按行下标索引的表的并集**，不能只看格子。
   *
   * 这个 extent 驱动每一次结构操作：矩阵的搬运，以及每张侧表的重键。而侧表
   * 可以有格子范围之外的条目——用户给一个没有任何格子的空行设了行高，是完全
   * 正常的状态。漏掉一张表，超出格子范围的那部分就被**每一次**结构操作静默
   * 丢弃。本分支上这件事已经出现过三次。
   *
   * **当前参与并集的表，就是下面逐一列举的这些：**
   * 1. `#matrix` 的格子（`maxOccupiedRow() + 1`）
   * 2. `#rowData`（行高与行高 provenance，spec §5）
   *
   * **给后来者的义务：任何新增的按行下标索引的表都必须加进这个并集。**
   * 阶段三的 `filteredRows` 与 spec §7.2 的按行索引插件资源区是第三、第四张，
   * 什么都没有强制它们进来——**漏掉是默认结果，不是风险**。
   * 依据：spec §11 不变量 8（声明 `indexedBy: "row"` 的资源，其重键由核心提供）
   * 与 ADR 0004 §9 不变量 4（每新增一张按下标索引的侧表，必须同时为结构操作
   * 提供重键与逆操作，做不到就不要新增这张表）。
   */
  #occupiedRowExtent(): number {
    const maxCellRow = this.#matrix.maxOccupiedRow();
    let extent = maxCellRow === undefined ? 0 : maxCellRow + 1;
    for (const key of Object.keys(this.#rowData)) {
      const row = Number(key);
      if (row + 1 > extent) {
        extent = row + 1;
      }
    }
    return extent;
  }

  #applyRowPlan(plan: Int32Array, sourceExtent: number): void {
    this.#matrix.applyRowPlan(plan, sourceExtent);
    this.#rowData = remapByRowPlan(this.#rowData, plan);
  }

  /**
   * 六个结构方法共用的参数防线，复用被替换的 `ArrayDataSource` 那一份
   * （`../layout/structuralChange`），不另写。三条口径：
   *
   * ① **放在 `#assertNoMerges` 之前**——参数非法与有没有合并区无关，
   *    先报 `MutationNotSupported` 会误导调用方。
   * ② **轴长取 `rowCount()` / `columnCount()`（坐标空间，spec §3.3），
   *    不是占用范围**——"在空白区里选几行删掉"是 Excel 允许的合法操作，
   *    用占用范围会把它误挡掉。
   * ③ **抛的是 `RangeError` 而不是 `CyExcelError`，不包装**——与被替换的
   *    `ArrayDataSource` 保持同一种失败方式，比与 `#assertNoMerges` 内部
   *    一致更重要；错误类型的统一属于后续公开 API 收口。
   *
   * 没有这道防线时的静默后果：`insertRows(-2, 1)` 让 `newRow - count`
   * 算出 `-1`，而 `-1` 就是 `ROW_PLAN_HOLE`——哨兵撞车，一行真数据变成
   * 一个洞；`moveRows(0, 1, -1)` 走 `rest.splice(负数)` 从末尾倒插；
   * `moveColumns(0, 2, 1)` 算出 `insertAt = -1`，产出负的列键。
   */
  insertRows(start: number, count: number): void {
    validateAxisChange({ axis: "row", operation: "insert", start, count }, this.rowCount());
    this.#assertNoMerges("insertRows");
    const extent = this.#occupiedRowExtent();
    this.#applyRowPlan(buildInsertRowsPlan(extent, start, count), extent);
  }

  deleteRows(start: number, count: number): void {
    validateAxisChange({ axis: "row", operation: "delete", start, count }, this.rowCount());
    this.#assertNoMerges("deleteRows");
    const extent = this.#occupiedRowExtent();
    this.#applyRowPlan(buildDeleteRowsPlan(extent, start, count), extent);
  }

  moveRows(start: number, count: number, target: number): void {
    validateAxisChange(
      { axis: "row", operation: "move", start, count, target, targetSpace: "pre-removal" },
      this.rowCount()
    );
    this.#assertNoMerges("moveRows");
    const extent = this.#occupiedRowExtent();
    this.#applyRowPlan(buildMoveRowsPlan(extent, start, count, target), extent);
  }

  reorderRows(order: Int32Array): void {
    this.#assertNoMerges("reorderRows");
    const extent = this.#occupiedRowExtent();
    if (extent > order.length) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidRange,
        `occupied row ${extent - 1} >= order length ${order.length}, would be silently dropped`
      );
    }
    this.#matrix.reorderRows(order);
    this.#rowData = remapByRowPlan(this.#rowData, order);
  }

  /**
   * 列轴对 A1 那一类"占用范围外丢数据"天然免疫：`#remapColumns` 是逐个
   * **占用列**过的纯算术，不以 extent 为界——`deleteColumns(10, 2)` 在只占用
   * 0–4 列时全部 `< 10`，原样不动。所以列轴只补上面那道参数校验。
   */
  insertColumns(start: number, count: number): void {
    validateAxisChange(
      { axis: "column", operation: "insert", start, count },
      this.columnCount()
    );
    this.#assertNoMerges("insertColumns");
    this.#matrix.insertColumns(start, count);
    this.#columnData = remapByColumnShift(
      this.#columnData,
      buildInsertColumnsShift(start, count)
    );
  }

  deleteColumns(start: number, count: number): void {
    validateAxisChange(
      { axis: "column", operation: "delete", start, count },
      this.columnCount()
    );
    this.#assertNoMerges("deleteColumns");
    this.#matrix.deleteColumns(start, count);
    this.#columnData = remapByColumnShift(
      this.#columnData,
      buildDeleteColumnsShift(start, count)
    );
  }

  moveColumns(start: number, count: number, target: number): void {
    validateAxisChange(
      { axis: "column", operation: "move", start, count, target, targetSpace: "pre-removal" },
      this.columnCount()
    );
    this.#assertNoMerges("moveColumns");
    this.#matrix.moveColumns(start, count, target);
    this.#columnData = remapByColumnShift(
      this.#columnData,
      buildMoveColumnsShift(start, count, target)
    );
  }

  getRowMeta(row: number): RowMeta | undefined {
    return this.#rowData[row];
  }

  setRowMeta(row: number, meta: RowMeta): void {
    this.#rowData = { ...this.#rowData, [row]: meta };
  }

  getColumnMeta(column: number): ColumnMeta | undefined {
    return this.#columnData[column];
  }

  setColumnMeta(column: number, meta: ColumnMeta): void {
    this.#columnData = { ...this.#columnData, [column]: meta };
  }

  merges(): readonly CellRange[] {
    return this.#mergeData;
  }

  occupiedCount(): number {
    return this.#matrix.occupiedCount();
  }

  /** 最大占用列号加一——与 `columnCount()`（坐标空间）是两个不同的数。 */
  occupiedColumnCount(): number {
    const max = this.#matrix.maxOccupiedColumn();
    return max === undefined ? 0 : max + 1;
  }
}
