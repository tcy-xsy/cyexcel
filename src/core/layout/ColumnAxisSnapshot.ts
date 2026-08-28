export interface ColumnAxisSnapshotOptions {
  readonly columnWidths: readonly number[];
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
 *
 * **只有几何宽度，没有第二份存储**（ADR 0004 阶段 4B T8b，设计稿 §4.4）。列宽与
 * 「这列是不是隐藏的」的权威在 `SheetModel.columnData[c]` 的 `w`/`hd`；这里存的是
 * 由它派生的那个数——隐藏列即 0。曾经并存的 `hiddenColumnWidths`（记「取消隐藏后
 * 恢复成多宽」）连同 `getColumnRestoreWidth`/`isColumnHidden` 一起删除：那份记忆
 * 现在就是权威里的 `w`，留在这里只会多出一处能与权威分叉的地方。
 */
export class ColumnAxisSnapshot {
  readonly #columnWidths: readonly number[];
  readonly #columnOffsets: readonly number[];

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

    this.#columnWidths = Object.freeze(columnWidths);
    this.#columnOffsets = buildOffsets(columnWidths);
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
}
