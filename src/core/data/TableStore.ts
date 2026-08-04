import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { CellRange, DataSource, Patch } from "../types";
import { assertValidCoordinate, assertValidRange } from "../validation";

export interface TableStoreCommitResult {
  readonly patches: readonly Patch[];
  readonly version: number;
}

/**
 * Sole entry point for reading and writing cell data. Every write goes
 * through here so version increments and Patch emission stay exact — there
 * is no separate path that reaches the underlying DataSource directly.
 */
export class TableStore {
  readonly #dataSource: DataSource;
  #version = 0;

  constructor(dataSource: DataSource) {
    this.#dataSource = dataSource;
  }

  get version(): number {
    return this.#version;
  }

  getCell(row: number, column: number): unknown {
    assertValidCoordinate(row, column);
    this.#assertColumnExists(column);
    return this.#dataSource.getCell(row, column);
  }

  setCell(row: number, column: number, value: unknown): TableStoreCommitResult {
    assertValidCoordinate(row, column);
    this.#assertColumnExists(column);

    this.#dataSource.setCell(row, column, value);

    return this.#commit([{ type: "cell-value", row, column }]);
  }

  setRange(
    range: CellRange,
    values: readonly (readonly unknown[])[]
  ): TableStoreCommitResult {
    assertValidRange(range);
    this.#assertColumnExists(range.endColumn);

    for (let row = range.startRow; row <= range.endRow; row += 1) {
      const rowValues = values[row - range.startRow];
      for (let column = range.startColumn; column <= range.endColumn; column += 1) {
        this.#dataSource.setCell(row, column, rowValues?.[column - range.startColumn]);
      }
    }

    return this.#commit([{ type: "range-value", range }]);
  }

  updateCell(
    row: number,
    column: number,
    updater: (current: unknown) => unknown
  ): TableStoreCommitResult {
    assertValidCoordinate(row, column);
    this.#assertColumnExists(column);

    const current = this.#dataSource.getCell(row, column);
    this.#dataSource.setCell(row, column, updater(current));

    return this.#commit([{ type: "cell-value", row, column }]);
  }

  #assertColumnExists(column: number): void {
    if (column >= this.#dataSource.getColumnCount()) {
      throw new CyExcelError(
        CyExcelErrorCode.ColumnNotFound,
        `Column ${column} does not exist`
      );
    }
  }

  #commit(patches: readonly Patch[]): TableStoreCommitResult {
    this.#version += 1;
    return { patches: Object.freeze([...patches]), version: this.#version };
  }
}
