import { CyExcelError, CyExcelErrorCode } from "../errors";

export interface LayoutOptions {
  readonly rowCount: number;
  readonly columnCount: number;
  readonly rowHeight: number;
  /** One width per column; length must equal `columnCount`. A width of 0 is a "hidden" column. */
  readonly columnWidths: readonly number[];
  readonly freeze?: {
    readonly rows?: number;
    readonly columns?: number;
  };
  readonly overscanRows?: number;
  readonly overscanColumns?: number;
}

export interface ViewportState {
  readonly scrollTop: number;
  readonly scrollLeft: number;
  readonly width: number;
  readonly height: number;
}

/** `count === 0` means nothing is visible; `start` and `count` are always >= 0. */
export interface IndexRange {
  readonly start: number;
  readonly count: number;
}

export interface CellRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

function isNonNegativeInteger(value: number): boolean {
  return Number.isInteger(value) && value >= 0;
}

function isPositiveFinite(value: number): boolean {
  return Number.isFinite(value) && value > 0;
}

function isNonNegativeFinite(value: number): boolean {
  return Number.isFinite(value) && value >= 0;
}

function invalidLayoutOption(field: string, value: unknown): CyExcelError {
  return new CyExcelError(
    CyExcelErrorCode.InvalidOptions,
    `Invalid LayoutOptions.${field}: ${String(value)}`
  );
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

function buildColumnOffsets(columnWidths: readonly number[]): number[] {
  const offsets: number[] = [0];
  let total = 0;
  for (const width of columnWidths) {
    total += width;
    offsets.push(total);
  }
  return offsets;
}

/** Largest index `i` in `offsets` such that `offsets[i] <= value` (offsets is ascending, offsets[0] === 0). */
function findOffsetIndex(offsets: readonly number[], value: number): number {
  let low = 0;
  let high = offsets.length - 1;

  while (low < high) {
    const mid = Math.ceil((low + high) / 2);
    // `mid` is always within [low, high] ⊆ [0, offsets.length - 1] by the
    // loop invariant, so this index always exists.
    const midValue = offsets[mid] as number;
    if (midValue <= value) {
      low = mid;
    } else {
      high = mid - 1;
    }
  }

  return low;
}

function validateLayoutOptions(options: LayoutOptions): void {
  if (!isNonNegativeInteger(options.rowCount)) {
    throw invalidLayoutOption("rowCount", options.rowCount);
  }
  if (!isNonNegativeInteger(options.columnCount)) {
    throw invalidLayoutOption("columnCount", options.columnCount);
  }
  if (!isPositiveFinite(options.rowHeight)) {
    throw invalidLayoutOption("rowHeight", options.rowHeight);
  }
  if (options.columnWidths.length !== options.columnCount) {
    throw invalidLayoutOption("columnWidths.length", options.columnWidths.length);
  }
  for (const width of options.columnWidths) {
    if (!isNonNegativeFinite(width)) {
      throw invalidLayoutOption("columnWidths[]", width);
    }
  }
  if (
    options.overscanRows !== undefined &&
    !isNonNegativeInteger(options.overscanRows)
  ) {
    throw invalidLayoutOption("overscanRows", options.overscanRows);
  }
  if (
    options.overscanColumns !== undefined &&
    !isNonNegativeInteger(options.overscanColumns)
  ) {
    throw invalidLayoutOption("overscanColumns", options.overscanColumns);
  }
  if (
    options.freeze?.rows !== undefined &&
    !isNonNegativeInteger(options.freeze.rows)
  ) {
    throw invalidLayoutOption("freeze.rows", options.freeze.rows);
  }
  if (
    options.freeze?.columns !== undefined &&
    !isNonNegativeInteger(options.freeze.columns)
  ) {
    throw invalidLayoutOption("freeze.columns", options.freeze.columns);
  }
}

/**
 * Pure geometry: given row/column counts, a fixed row height, per-column
 * widths, frozen row/column counts, and a scroll position, compute visible
 * ranges, cell rectangles, and point hit-testing. Never touches Canvas or
 * the DOM, and never takes a DPR — see doc/architecture.md §5-6 and
 * `toPhysicalPixelSize` (./dpr.ts) for the physical-pixel concern.
 */
export class LayoutEngine {
  readonly #rowCount: number;
  readonly #columnCount: number;
  readonly #rowHeight: number;
  readonly #columnWidths: readonly number[];
  readonly #columnOffsets: readonly number[];
  readonly #frozenRowCount: number;
  readonly #frozenColumnCount: number;
  readonly #overscanRows: number;
  readonly #overscanColumns: number;

  constructor(options: LayoutOptions) {
    validateLayoutOptions(options);

    this.#rowCount = options.rowCount;
    this.#columnCount = options.columnCount;
    this.#rowHeight = options.rowHeight;
    this.#columnWidths = options.columnWidths;
    this.#columnOffsets = buildColumnOffsets(options.columnWidths);
    this.#frozenRowCount = Math.min(options.freeze?.rows ?? 0, options.rowCount);
    this.#frozenColumnCount = Math.min(
      options.freeze?.columns ?? 0,
      options.columnCount
    );
    this.#overscanRows = options.overscanRows ?? 0;
    this.#overscanColumns = options.overscanColumns ?? 0;
  }

  get rowCount(): number {
    return this.#rowCount;
  }

  get columnCount(): number {
    return this.#columnCount;
  }

  get rowHeight(): number {
    return this.#rowHeight;
  }

  get frozenRowCount(): number {
    return this.#frozenRowCount;
  }

  get frozenColumnCount(): number {
    return this.#frozenColumnCount;
  }

  getColumnWidth(column: number): number {
    return this.#columnWidths[column] ?? 0;
  }

  getTotalWidth(): number {
    // #columnOffsets has length columnCount + 1, so this index always exists.
    return this.#columnOffsets[this.#columnCount] as number;
  }

  getTotalHeight(): number {
    return this.#rowCount * this.#rowHeight;
  }

  getCellRect(row: number, column: number): CellRect {
    return {
      x: this.#columnOffsets[column] ?? 0,
      y: row * this.#rowHeight,
      width: this.getColumnWidth(column),
      height: this.#rowHeight
    };
  }

  getRowAtY(y: number): number {
    if (this.#rowCount === 0) {
      return 0;
    }
    return clamp(this.#rowIndexFromOffset(y), 0, this.#rowCount - 1);
  }

  getColumnAtX(x: number): number {
    if (this.#columnCount === 0) {
      return 0;
    }
    return clamp(this.#columnIndexFromOffset(x), 0, this.#columnCount - 1);
  }

  getVisibleRowRange(viewport: ViewportState): IndexRange {
    const scrollableRowCount = this.#rowCount - this.#frozenRowCount;
    const frozenHeight = this.#frozenRowCount * this.#rowHeight;
    const scrollableHeight = Math.max(0, viewport.height - frozenHeight);

    if (scrollableRowCount <= 0 || scrollableHeight <= 0) {
      return { start: this.#frozenRowCount, count: 0 };
    }

    const scrollTop = Math.max(0, viewport.scrollTop);
    const firstIndex = this.#frozenRowCount + Math.floor(scrollTop / this.#rowHeight);
    const visibleCount = Math.ceil(scrollableHeight / this.#rowHeight) + 1;
    const lastIndex = firstIndex + visibleCount - 1;

    const start = clamp(
      firstIndex - this.#overscanRows,
      this.#frozenRowCount,
      this.#rowCount - 1
    );
    const end = clamp(
      lastIndex + this.#overscanRows,
      this.#frozenRowCount,
      this.#rowCount - 1
    );

    return { start, count: end - start + 1 };
  }

  getVisibleColumnRange(viewport: ViewportState): IndexRange {
    const scrollableColumnCount = this.#columnCount - this.#frozenColumnCount;
    const frozenWidth = this.#frozenColumnWidth();
    const scrollableWidth = Math.max(0, viewport.width - frozenWidth);

    if (scrollableColumnCount <= 0 || scrollableWidth <= 0) {
      return { start: this.#frozenColumnCount, count: 0 };
    }

    const scrollLeft = Math.max(0, viewport.scrollLeft);
    const contentLeft = frozenWidth + scrollLeft;
    const contentRight = contentLeft + scrollableWidth;

    const firstIndex = clamp(
      this.#columnIndexFromOffset(contentLeft),
      this.#frozenColumnCount,
      this.#columnCount - 1
    );
    const lastIndex = clamp(
      this.#columnIndexFromOffset(contentRight),
      this.#frozenColumnCount,
      this.#columnCount - 1
    );

    const start = clamp(
      firstIndex - this.#overscanColumns,
      this.#frozenColumnCount,
      this.#columnCount - 1
    );
    const end = clamp(
      lastIndex + this.#overscanColumns,
      this.#frozenColumnCount,
      this.#columnCount - 1
    );

    return { start, count: end - start + 1 };
  }

  getCellAtPoint(
    point: { x: number; y: number },
    viewport: ViewportState
  ): { row: number; column: number } | null {
    if (point.x < 0 || point.y < 0) {
      return null;
    }
    if (this.#rowCount === 0 || this.#columnCount === 0) {
      return null;
    }

    const frozenHeight = this.#frozenRowCount * this.#rowHeight;
    const frozenWidth = this.#frozenColumnWidth();

    let row: number;
    if (point.y < frozenHeight) {
      row = clamp(this.#rowIndexFromOffset(point.y), 0, this.#frozenRowCount - 1);
    } else {
      const scrollTop = Math.max(0, viewport.scrollTop);
      row = this.#rowIndexFromOffset(scrollTop + point.y);
      if (row >= this.#rowCount) {
        return null;
      }
    }

    let column: number;
    if (point.x < frozenWidth) {
      column = clamp(
        this.#columnIndexFromOffset(point.x),
        0,
        this.#frozenColumnCount - 1
      );
    } else {
      const scrollLeft = Math.max(0, viewport.scrollLeft);
      column = this.#columnIndexFromOffset(scrollLeft + point.x);
      if (column >= this.#columnCount) {
        return null;
      }
    }

    return { row, column };
  }

  #rowIndexFromOffset(offset: number): number {
    return Math.floor(offset / this.#rowHeight);
  }

  #columnIndexFromOffset(offset: number): number {
    return findOffsetIndex(this.#columnOffsets, offset);
  }

  #frozenColumnWidth(): number {
    // frozenColumnCount is clamped to [0, columnCount] in the constructor,
    // and #columnOffsets has length columnCount + 1, so this index always exists.
    return this.#columnOffsets[this.#frozenColumnCount] as number;
  }
}
