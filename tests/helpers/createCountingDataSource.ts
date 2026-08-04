import type { DataSource } from "../../src/core";

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
}

export function createCountingDataSource(
  rowCount: number,
  columnCount: number
): CountingDataSource {
  return new CountingDataSource(rowCount, columnCount);
}
