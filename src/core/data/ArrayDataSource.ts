import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { DataSource, RowData } from "../types";

export interface ArrayDataSourceColumn {
  readonly id: string;
  readonly field: string;
}

export interface ArrayDataSourceOptions {
  readonly readonly?: boolean;
}

/**
 * DataSource backed by a plain array of row objects. Rows and cell values
 * are kept as-is — no deep clone, no Proxy — per ADR 0001 §5 (complex
 * values are atomic).
 */
export class ArrayDataSource implements DataSource {
  readonly #rows: RowData[];
  readonly #columns: readonly ArrayDataSourceColumn[];
  readonly #readonly: boolean;

  constructor(
    rows: RowData[],
    columns: readonly ArrayDataSourceColumn[],
    options: ArrayDataSourceOptions = {}
  ) {
    this.#rows = rows;
    this.#columns = columns;
    this.#readonly = options.readonly ?? false;
  }

  getRowCount(): number {
    return this.#rows.length;
  }

  getColumnCount(): number {
    return this.#columns.length;
  }

  getCell(row: number, column: number): unknown {
    const field = this.#columns[column]?.field;
    if (field === undefined) {
      return undefined;
    }
    return this.#rows[row]?.[field];
  }

  setCell(row: number, column: number, value: unknown): void {
    if (this.#readonly) {
      throw new CyExcelError(
        CyExcelErrorCode.ReadonlyDataSource,
        "Cannot write to a readonly ArrayDataSource"
      );
    }

    const field = this.#columns[column]?.field;
    if (field === undefined) {
      throw new CyExcelError(
        CyExcelErrorCode.ColumnNotFound,
        `Column ${column} does not exist`
      );
    }

    const rowData = this.#rows[row];
    if (rowData === undefined) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidCoordinate,
        `Row ${row} does not exist`
      );
    }

    rowData[field] = value;
  }

  /**
   * `rows` is one array of column-ordered values per new row (same
   * convention as `setRange()`) — turned into `RowData` objects here,
   * using this instance's own field mapping, before splicing them in.
   * `splice()` shifts the existing array in place; it never allocates a
   * full copy of `#rows`, so this stays cheap regardless of total row
   * count (Step 12's "不整表复制" exit criterion).
   */
  insertRows(start: number, rows: readonly (readonly unknown[])[]): void {
    if (this.#readonly) {
      throw new CyExcelError(
        CyExcelErrorCode.ReadonlyDataSource,
        "Cannot write to a readonly ArrayDataSource"
      );
    }

    const rowObjects = rows.map((values) => {
      const rowData: RowData = {};
      this.#columns.forEach((column, index) => {
        rowData[column.field] = values[index];
      });
      return rowData;
    });
    this.#rows.splice(start, 0, ...rowObjects);
  }

  deleteRows(start: number, count: number): void {
    if (this.#readonly) {
      throw new CyExcelError(
        CyExcelErrorCode.ReadonlyDataSource,
        "Cannot write to a readonly ArrayDataSource"
      );
    }

    this.#rows.splice(start, count);
  }

  /**
   * `target` follows the same convention `LayoutEngine`/`SelectionModel`
   * assume: the insertion index in the *original* (pre-removal) index
   * space. The moved rows keep their exact object references — nothing
   * about "which row is this" changes, only position (see
   * doc/progress.md's Step 12 entry for why that's enough to satisfy
   * "stable identity" without a dedicated row-ID API).
   */
  moveRows(start: number, count: number, target: number): void {
    if (this.#readonly) {
      throw new CyExcelError(
        CyExcelErrorCode.ReadonlyDataSource,
        "Cannot write to a readonly ArrayDataSource"
      );
    }

    const moved = this.#rows.splice(start, count);
    const adjustedTarget = target > start ? target - count : target;
    this.#rows.splice(adjustedTarget, 0, ...moved);
  }
}
