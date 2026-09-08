import type { CellRange } from "../types";
import { ColumnAxisSnapshot } from "./ColumnAxisSnapshot";
import type { CellRect, IndexRange, LayoutEngine, LayoutRoot } from "./LayoutEngine";
import { RowMetrics } from "./RowMetrics";

interface SkeletonSnapshotBaseOptions {
  readonly modelVersion: number;
  readonly layoutVersion: number;
  readonly columnAxis: ColumnAxisSnapshot;
  readonly frozenRowCount: number;
}

export type SkeletonSnapshotOptions = SkeletonSnapshotBaseOptions & (
  | { readonly rowHeights: ArrayLike<number>; readonly rowMetrics?: never }
  | { readonly rowMetrics: RowMetrics; readonly rowHeights?: never }
);

function assertNonNegativeInteger(name: string, value: number): void {
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new RangeError(`${name} must be a non-negative safe integer: ${value}`);
  }
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

/**
 * Immutable logical sheet geometry for one committed model/layout version.
 * Layout publications retain the immutable RowMetrics root, so old and new
 * scenes can coexist without copying million-row geometry.
 */
export class SkeletonSnapshot {
  readonly rowMetrics: RowMetrics;
  readonly columnAxis: ColumnAxisSnapshot;

  readonly modelVersion: number;
  readonly layoutVersion: number;
  readonly rowCount: number;
  readonly columnCount: number;
  readonly frozenRowCount: number;
  readonly frozenColumnCount: number;

  constructor(options: SkeletonSnapshotOptions) {
    assertNonNegativeInteger("modelVersion", options.modelVersion);
    assertNonNegativeInteger("layoutVersion", options.layoutVersion);

    const rowMetrics = "rowMetrics" in options && options.rowMetrics !== undefined
      ? options.rowMetrics
      : RowMetrics.fromDenseValues(options.rowHeights);
    assertNonNegativeInteger("frozenRowCount", options.frozenRowCount);
    if (options.frozenRowCount > rowMetrics.rowCount) {
      throw new RangeError("frozenRowCount cannot exceed rowCount");
    }
    if (!(options.columnAxis instanceof ColumnAxisSnapshot)) {
      throw new TypeError("columnAxis must be a ColumnAxisSnapshot");
    }

    this.modelVersion = options.modelVersion;
    this.layoutVersion = options.layoutVersion;
    this.rowCount = rowMetrics.rowCount;
    this.columnCount = options.columnAxis.columnCount;
    this.frozenRowCount = options.frozenRowCount;
    this.frozenColumnCount = options.columnAxis.frozenColumnCount;
    this.rowMetrics = rowMetrics;
    this.columnAxis = options.columnAxis;
    Object.freeze(this);
  }

  static fromLayout(layout: LayoutEngine, modelVersion: number): SkeletonSnapshot {
    return SkeletonSnapshot.fromRoot(layout.root, modelVersion);
  }

  static fromRoot(root: LayoutRoot, modelVersion: number): SkeletonSnapshot {
    return new SkeletonSnapshot({
      modelVersion,
      layoutVersion: root.version,
      // `visibleRowMetrics`, not the authoritative `rowMetrics` — this is the
      // one place filtering enters rendering/scroll/hit-test (ADR 0003 §3.2).
      rowMetrics: root.visibleRowMetrics,
      columnAxis: root.columnAxis,
      frozenRowCount: root.frozenRowCount,
    });
  }

  getRowHeight(row: number): number {
    if (row < 0 || row >= this.rowCount) {
      return 0;
    }
    return this.rowMetrics.heightAt(row);
  }

  getRowOffset(row: number): number {
    if (!Number.isSafeInteger(row) || row < 0 || row > this.rowCount) {
      return 0;
    }
    return this.rowMetrics.offsetAt(row);
  }

  getColumnWidth(column: number): number {
    return this.columnAxis.getColumnWidth(column);
  }

  getColumnLeft(column: number): number {
    return this.columnAxis.getColumnOffset(column);
  }

  getCellRect(row: number, column: number): CellRect {
    return {
      x: this.getColumnLeft(column),
      y: this.getRowOffset(row),
      width: this.getColumnWidth(column),
      height: this.getRowHeight(row)
    };
  }

  getRangeRect(range: CellRange): CellRect {
    const topLeft = this.getCellRect(range.startRow, range.startColumn);
    const bottomRight = this.getCellRect(range.endRow, range.endColumn);
    return {
      x: topLeft.x,
      y: topLeft.y,
      width: bottomRight.x + bottomRight.width - topLeft.x,
      height: bottomRight.y + bottomRight.height - topLeft.y
    };
  }

  getTotalWidth(): number {
    return this.columnAxis.getTotalWidth();
  }

  getTotalHeight(): number {
    return this.rowMetrics.totalHeight;
  }

  getRowAtY(y: number): number {
    if (this.rowCount === 0) {
      return 0;
    }
    return clamp(this.rowMetrics.rowAtOffset(y), 0, this.rowCount - 1);
  }

  getColumnAtX(x: number): number {
    if (this.columnCount === 0) {
      return 0;
    }
    return clamp(this.columnAxis.getColumnAtOffset(x), 0, this.columnCount - 1);
  }

  getVisibleRowRange(viewport: {
    readonly scrollTop: number;
    readonly height: number;
    readonly overscanRows?: number;
  }): IndexRange {
    const scrollableRowCount = this.rowCount - this.frozenRowCount;
    const frozenHeight = this.getRowOffset(this.frozenRowCount);
    const scrollableHeight = Math.max(0, viewport.height - frozenHeight);

    if (scrollableRowCount <= 0 || scrollableHeight <= 0) {
      return { start: this.frozenRowCount, count: 0 };
    }

    const contentTop = frozenHeight + Math.max(0, viewport.scrollTop);
    const firstIndex = this.rowMetrics.rowAtOffset(contentTop);
    const lastIndex = this.rowMetrics.rowAtOffset(contentTop + scrollableHeight);
    const overscanRows = viewport.overscanRows ?? 0;
    const start = clamp(firstIndex - overscanRows, this.frozenRowCount, this.rowCount - 1);
    const end = clamp(lastIndex + overscanRows, this.frozenRowCount, this.rowCount - 1);
    return { start, count: end - start + 1 };
  }

  getVisibleColumnRange(viewport: {
    readonly scrollLeft: number;
    readonly width: number;
    readonly overscanColumns?: number;
  }): IndexRange {
    const scrollableColumnCount = this.columnCount - this.frozenColumnCount;
    const frozenWidth = this.getColumnLeft(this.frozenColumnCount);
    const scrollableWidth = Math.max(0, viewport.width - frozenWidth);

    if (scrollableColumnCount <= 0 || scrollableWidth <= 0) {
      return { start: this.frozenColumnCount, count: 0 };
    }

    const contentLeft = frozenWidth + Math.max(0, viewport.scrollLeft);
    const firstIndex = clamp(
      this.columnAxis.getColumnAtOffset(contentLeft),
      this.frozenColumnCount,
      this.columnCount - 1
    );
    const lastIndex = clamp(
      this.columnAxis.getColumnAtOffset(contentLeft + scrollableWidth),
      this.frozenColumnCount,
      this.columnCount - 1
    );
    const overscanColumns = viewport.overscanColumns ?? 0;
    const start = clamp(firstIndex - overscanColumns, this.frozenColumnCount, this.columnCount - 1);
    const end = clamp(lastIndex + overscanColumns, this.frozenColumnCount, this.columnCount - 1);
    return { start, count: end - start + 1 };
  }
}
