import type { CellRange } from "../types";
import { CellMatrix } from "./CellMatrix";
import { assertValidBlankCount, resolveCoordinateSpace } from "./coordinateSpace";
import { deserializeSheet, serializeSheet } from "./serialize";
import type { CellRecord, ColumnMeta, RowMeta, SheetSnapshot } from "./types";
import { DEFAULT_BLANK_COLUMNS, DEFAULT_BLANK_ROWS } from "./types";

export interface SheetModelOptions {
  readonly id?: string;
  readonly name?: string;
  readonly blankRows?: number;
  readonly blankColumns?: number;
  readonly defaultRowHeight?: number;
  readonly defaultColumnWidth?: number;
}

/**
 * 一张表的模型：稀疏记录矩阵 + 行列元数据 + 由 `blank*` 派生的坐标空间。
 *
 * 本阶段不接线到 `TableStore` / `LayoutEngine` / `CyExcelCore`——那是阶段 3。
 */
export class SheetModel {
  readonly #matrix: CellMatrix;
  readonly #rowData: Record<number, RowMeta>;
  readonly #columnData: Record<number, ColumnMeta>;
  #mergeData: CellRange[];
  readonly #id: string;
  readonly #name: string;
  readonly #blankRows: number;
  readonly #blankColumns: number;
  readonly #defaultRowHeight: number;
  readonly #defaultColumnWidth: number;

  private constructor(init: {
    matrix: CellMatrix;
    rowData: Record<number, RowMeta>;
    columnData: Record<number, ColumnMeta>;
    mergeData: CellRange[];
    id: string;
    name: string;
    blankRows: number;
    blankColumns: number;
    defaultRowHeight: number;
    defaultColumnWidth: number;
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
      defaultColumnWidth: options.defaultColumnWidth ?? 120
    });
  }

  static fromSnapshot(snapshot: SheetSnapshot): SheetModel {
    const restored = deserializeSheet(snapshot);
    return new SheetModel({
      matrix: restored.matrix,
      rowData: { ...restored.rowData },
      columnData: { ...restored.columnData },
      mergeData: [...restored.mergeData],
      id: restored.id,
      name: restored.name,
      blankRows: restored.blankRows,
      blankColumns: restored.blankColumns,
      defaultRowHeight: restored.defaultRowHeight,
      defaultColumnWidth: restored.defaultColumnWidth
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
      mergeData: this.#mergeData
    });
  }
}
