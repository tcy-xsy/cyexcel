import { CyExcelError, CyExcelErrorCode } from "../../src/core/errors";
import type {
  DataMutation,
  DataSource,
  DataSourceTransaction,
  PreparedDataSourceMutation
} from "../../src/core/types";
import type { ModelRevision } from "../../src/core/model";

/**
 * DataSource that counts getCell/setCell calls without materializing a
 * rowCount x columnCount array, so it stays cheap to construct even at
 * million-row scale. Cell values are synthesized deterministically unless
 * overwritten via setCell.
 */
export class CountingDataSource implements DataSource {
  readonly #rowCount: number;
  readonly #columnCount: number;
  readonly #overrides = new Map<string, unknown>();
  #readCount = 0;
  #writeCount = 0;

  constructor(rowCount: number, columnCount: number) {
    this.#rowCount = rowCount;
    this.#columnCount = columnCount;
  }

  get readCount(): number {
    return this.#readCount;
  }

  get writeCount(): number {
    return this.#writeCount;
  }

  getRowCount(): number {
    return this.#rowCount;
  }

  getColumnCount(): number {
    return this.#columnCount;
  }

  getCell(row: number, column: number): unknown {
    this.#readCount += 1;
    const key = `${row}:${column}`;
    return this.#overrides.has(key)
      ? this.#overrides.get(key)
      : row * this.#columnCount + column;
  }

  setCell(row: number, column: number, value: unknown): void {
    this.#writeCount += 1;
    this.#overrides.set(`${row}:${column}`, value);
  }

  beginTransaction(baseRevision: ModelRevision): DataSourceTransaction {
    const overlay = new Map<string, unknown>();
    return {
      baseRevision,
      getRowCount: () => this.#rowCount,
      getColumnCount: () => this.#columnCount,
      getCell: (row, column) => overlay.has(`${row}:${column}`)
        ? overlay.get(`${row}:${column}`)
        : this.getCell(row, column),
      prepareMutation: (mutation) => this.#prepareMutation(mutation, overlay),
      dispose: () => overlay.clear()
    };
  }

  #prepareMutation(
    mutation: DataMutation,
    overlay: Map<string, unknown>
  ): PreparedDataSourceMutation {
    if (mutation.type !== "set-cell") {
      throw new CyExcelError(CyExcelErrorCode.MutationNotSupported, "CountingDataSource only stages set-cell");
    }
    const key = `${mutation.row}:${mutation.column}`;
    const hadOverride = this.#overrides.has(key);
    const previous = overlay.has(key)
      ? overlay.get(key)
      : hadOverride
        ? this.#overrides.get(key)
        : mutation.row * this.#columnCount + mutation.column;
    overlay.set(key, mutation.value);
    let commitStarted = false;
    let rolledBack = false;
    return {
      forward: mutation,
      inverse: {
        type: "set-cell",
        row: mutation.row,
        column: mutation.column,
        value: previous
      },
      estimatedBytes: 64,
      commit: () => {
        commitStarted = true;
        this.setCell(mutation.row, mutation.column, mutation.value);
      },
      rollback: () => {
        if (!commitStarted || rolledBack) {
          return;
        }
        rolledBack = true;
        if (hadOverride) {
          this.#overrides.set(key, previous);
        } else {
          this.#overrides.delete(key);
        }
      },
      dispose: () => {}
    };
  }
}

export function createCountingDataSource(
  rowCount: number,
  columnCount: number
): CountingDataSource {
  return new CountingDataSource(rowCount, columnCount);
}
