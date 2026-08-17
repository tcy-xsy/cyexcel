export interface ColumnAxisSnapshotOptions {
  readonly columnWidths: readonly number[];
  readonly hiddenColumnWidths?: ReadonlyMap<number, number>;
  readonly frozenColumnCount: number;
}

function assertNonNegativeInteger(name: string, value: number): void {
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new RangeError(`${name} must be a non-negative safe integer: ${value}`);
  }
}

function assertNonNegativeFinite(name: string, value: number): void {
  if (!Number.isFinite(value) || value < 0) {
    throw new RangeError(`${name} must be a non-negative finite number: ${value}`);
  }
}

function buildOffsets(columnWidths: readonly number[]): readonly number[] {
  const offsets: number[] = [0];
  let total = 0;
  for (const width of columnWidths) {
    total += width;
    offsets.push(total);
  }
  return Object.freeze(offsets);
}

/**
 * Immutable column geometry shared by Layout roots and every Skeleton/Scene
 * publication derived from them. Dense widths and offsets are copied exactly
 * once when a column mutation creates a new axis.
 */
export class ColumnAxisSnapshot {
  readonly #columnWidths: readonly number[];
  readonly #columnOffsets: readonly number[];
  readonly #hiddenColumnWidths: ReadonlyMap<number, number>;

  readonly columnCount: number;
  readonly frozenColumnCount: number;

  constructor(options: ColumnAxisSnapshotOptions) {
    const columnWidths = options.columnWidths.map((width, column) => {
      assertNonNegativeFinite(`columnWidths[${column}]`, width);
      return width;
    });
    assertNonNegativeInteger("frozenColumnCount", options.frozenColumnCount);
    if (options.frozenColumnCount > columnWidths.length) {
      throw new RangeError("frozenColumnCount cannot exceed columnCount");
    }

    const hiddenColumnWidths = new Map<number, number>();
    for (const [column, width] of options.hiddenColumnWidths ?? []) {
      assertNonNegativeInteger("hiddenColumnWidths key", column);
      assertNonNegativeFinite("hiddenColumnWidths value", width);
      if (column >= columnWidths.length) {
        throw new RangeError(`hiddenColumnWidths key exceeds columnCount: ${column}`);
      }
      if (columnWidths[column] !== 0) {
        throw new RangeError(`hidden column ${column} must have zero layout width`);
      }
      hiddenColumnWidths.set(column, width);
    }

    this.#columnWidths = Object.freeze(columnWidths);
    this.#columnOffsets = buildOffsets(columnWidths);
    this.#hiddenColumnWidths = hiddenColumnWidths;
    this.columnCount = columnWidths.length;
    this.frozenColumnCount = options.frozenColumnCount;
    Object.freeze(this);
  }

  /** Frozen boundary table shared with the Layout root. */
  get columnOffsets(): readonly number[] {
    return this.#columnOffsets;
  }

  getColumnWidth(column: number): number {
    return this.#columnWidths[column] ?? 0;
  }

  /** Accepts the boundary after the final column (`column === columnCount`). */
  getColumnOffset(column: number): number {
    return this.#columnOffsets[column] ?? 0;
  }

  getColumnRestoreWidth(column: number): number {
    return this.#hiddenColumnWidths.get(column) ?? this.getColumnWidth(column);
  }

  isColumnHidden(column: number): boolean {
    return this.#hiddenColumnWidths.has(column);
  }

  getTotalWidth(): number {
    return this.#columnOffsets[this.columnCount] as number;
  }

  /** Largest boundary index whose offset is <= x. */
  getColumnAtOffset(x: number): number {
    let low = 0;
    let high = this.#columnOffsets.length - 1;
    while (low < high) {
      const mid = Math.ceil((low + high) / 2);
      if ((this.#columnOffsets[mid] as number) <= x) {
        low = mid;
      } else {
        high = mid - 1;
      }
    }
    return low;
  }

  /** Explicit O(columnCount) mutation/serialization boundary. */
  copyColumnWidths(): number[] {
    return [...this.#columnWidths];
  }

  /** Returns a detached mutable copy; the retained hidden-width map never escapes. */
  copyHiddenColumnWidths(): Map<number, number> {
    return new Map(this.#hiddenColumnWidths);
  }
}
