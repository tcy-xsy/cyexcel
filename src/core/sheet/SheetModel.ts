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
   * 行的"占用范围"必须取**格子与元数据的并集**，不能只看格子。
   *
   * 这个 extent 同时驱动两件事：矩阵的搬运，以及 `rowData` 的重键。而元数据
   * 可以存在于没有任何格子的行上——用户给一个空行设了行高，是完全正常的状态。
   * 只按格子算的话，超出格子范围的那些行高会被每一次结构操作静默丢弃。
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

  insertRows(start: number, count: number): void {
    this.#assertNoMerges("insertRows");
    const extent = this.#occupiedRowExtent();
    this.#applyRowPlan(buildInsertRowsPlan(extent, start, count), extent);
  }

  deleteRows(start: number, count: number): void {
    this.#assertNoMerges("deleteRows");
    const extent = this.#occupiedRowExtent();
    this.#applyRowPlan(buildDeleteRowsPlan(extent, start, count), extent);
  }

  moveRows(start: number, count: number, target: number): void {
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

  insertColumns(start: number, count: number): void {
    this.#assertNoMerges("insertColumns");
    this.#matrix.insertColumns(start, count);
    this.#columnData = remapByColumnShift(this.#columnData, (column) =>
      column < start ? column : column + count
    );
  }

  deleteColumns(start: number, count: number): void {
    this.#assertNoMerges("deleteColumns");
    this.#matrix.deleteColumns(start, count);
    this.#columnData = remapByColumnShift(this.#columnData, (column) => {
      if (column < start) {
        return column;
      }
      if (column < start + count) {
        return undefined;
      }
      return column - count;
    });
  }

  moveColumns(start: number, count: number, target: number): void {
    this.#assertNoMerges("moveColumns");
    this.#matrix.moveColumns(start, count, target);
    const insertAt = target > start ? target - count : target;
    this.#columnData = remapByColumnShift(this.#columnData, (column) => {
      if (column >= start && column < start + count) {
        return insertAt + (column - start);
      }
      const withoutBlock = column > start ? column - count : column;
      return withoutBlock >= insertAt ? withoutBlock + count : withoutBlock;
    });
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
