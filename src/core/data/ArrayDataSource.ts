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
}
