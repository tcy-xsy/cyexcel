import { CyExcelError, CyExcelErrorCode } from '../errors';
import { ColumnAxisSnapshot } from './ColumnAxisSnapshot';
import { RowMetrics, type RowMetricsFragment } from './RowMetrics';
import type { ColumnStructureChange } from './columnStructure';
import {
  remapIndex,
  validateAxisChange,
  type AxisChangeDescriptor,
} from './structuralChange';

export interface LayoutOptions {
  readonly rowCount: number;
  readonly columnCount: number;
  /** Default height for any row without an explicit override. */
  readonly rowHeight: number;
  /** Row index -> explicit height, only for rows that differ from `rowHeight`. */
  readonly rowHeights?: ReadonlyMap<number, number>;
  /** One geometric width per column; length must equal `columnCount`. Visibility is tracked independently. */
  readonly columnWidths: readonly number[];
  readonly freeze?: {
    readonly rows?: number;
    readonly columns?: number;
  };
  readonly overscanRows?: number;
  readonly overscanColumns?: number;
}

export type RowHeightProvenance = 'default' | 'auto' | 'manual';

export interface LayoutRowHeightUpdate {
  readonly height: number;
  readonly provenance: RowHeightProvenance;
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

export interface LayoutRootSnapshot {
  readonly version: number;
  readonly rowCount: number;
  readonly rowMetrics: RowMetrics;
  readonly rowHeightProvenance: ReadonlyMap<number, Exclude<RowHeightProvenance, 'default'>>;
  readonly columnAxis: ColumnAxisSnapshot;
}

export interface LayoutRoot {
  readonly version: number;
  readonly rowCount: number;
  readonly columnCount: number;
  readonly rowMetrics: RowMetrics;
  /** Sparse ownership metadata. Missing rows have `default` provenance. */
  readonly rowHeightProvenance: ReadonlyMap<number, Exclude<RowHeightProvenance, 'default'>>;
  readonly columnAxis: ColumnAxisSnapshot;
  readonly frozenRowCount: number;
  readonly frozenColumnCount: number;
}

export type PreparedLayoutRoot = LayoutRoot;

export interface LayoutStructurePayload {
  readonly rowHeights?: ReadonlyMap<number, number> | undefined;
  readonly rowMetricsFragment?: RowMetricsFragment | undefined;
  /** Row offsets within an inserted fragment. Missing offsets are `default`. */
  readonly rowHeightProvenance?: ReadonlyMap<number, Exclude<RowHeightProvenance, 'default'>> | undefined;
  readonly columnWidths?: readonly number[] | undefined;
  readonly hiddenColumnOffsets?: ReadonlyMap<number, number> | undefined;
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
    `Invalid LayoutOptions.${field}: ${String(value)}`,
  );
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

function createReadonlyMap<K, V>(source: ReadonlyMap<K, V>): ReadonlyMap<K, V> {
  const snapshot = new Map(source);
  const readonlyMap = {
    get size() {
      return snapshot.size;
    },
    get(key: K): V | undefined {
      return snapshot.get(key);
    },
    has(key: K): boolean {
      return snapshot.has(key);
    },
    entries() {
      return snapshot.entries();
    },
    keys() {
      return snapshot.keys();
    },
    values() {
      return snapshot.values();
    },
    forEach(
      callback: (value: V, key: K, map: ReadonlyMap<K, V>) => void,
      thisArg?: unknown,
    ) {
      for (const [key, value] of snapshot) {
        callback.call(thisArg, value, key, readonlyMap);
      }
    },
    [Symbol.iterator]() {
      return snapshot[Symbol.iterator]();
    },
  } as ReadonlyMap<K, V>;
  return Object.freeze(readonlyMap);
}


function validateLayoutOptions(options: LayoutOptions): void {
  if (!isNonNegativeInteger(options.rowCount)) {
    throw invalidLayoutOption('rowCount', options.rowCount);
  }
  if (!isNonNegativeInteger(options.columnCount)) {
    throw invalidLayoutOption('columnCount', options.columnCount);
  }
  if (!isPositiveFinite(options.rowHeight)) {
    throw invalidLayoutOption('rowHeight', options.rowHeight);
  }
  if (options.rowHeights) {
    for (const [row, height] of options.rowHeights) {
      if (!isNonNegativeInteger(row) || row >= options.rowCount) {
        throw invalidLayoutOption('rowHeights key', row);
      }
      if (!isPositiveFinite(height)) {
        throw invalidLayoutOption('rowHeights value', height);
      }
    }
  }
  if (options.columnWidths.length !== options.columnCount) {
    throw invalidLayoutOption(
      'columnWidths.length',
      options.columnWidths.length,
    );
  }
  for (const width of options.columnWidths) {
    if (!isNonNegativeFinite(width)) {
      throw invalidLayoutOption('columnWidths[]', width);
    }
  }
  if (
    options.overscanRows !== undefined &&
    !isNonNegativeInteger(options.overscanRows)
  ) {
    throw invalidLayoutOption('overscanRows', options.overscanRows);
  }
  if (
    options.overscanColumns !== undefined &&
    !isNonNegativeInteger(options.overscanColumns)
  ) {
    throw invalidLayoutOption('overscanColumns', options.overscanColumns);
  }
  if (
    options.freeze?.rows !== undefined &&
    !isNonNegativeInteger(options.freeze.rows)
  ) {
    throw invalidLayoutOption('freeze.rows', options.freeze.rows);
  }
  if (
    options.freeze?.columns !== undefined &&
    !isNonNegativeInteger(options.freeze.columns)
  ) {
    throw invalidLayoutOption('freeze.columns', options.freeze.columns);
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
  #root: LayoutRoot;
  /** Default height for inserted/grown rows; #root.rowMetrics is the source of truth. */
  readonly #rowHeight: number;
  /** The freeze.rows the caller originally asked for, before clamping to the current rowCount — structural root replacements re-clamp against this, so a shrink-then-grow sequence can recover frozen rows instead of only ever losing them. */
  readonly #requestedFrozenRows: number;
  /** Column-axis counterpart to #requestedFrozenRows — same "re-clamp against the original request, not the previous value" reasoning. */
  readonly #requestedFrozenColumns: number;
  readonly #overscanRows: number;
  readonly #overscanColumns: number;

  constructor(options: LayoutOptions) {
    validateLayoutOptions(options);

    this.#rowHeight = options.rowHeight;
    this.#requestedFrozenRows = options.freeze?.rows ?? 0;
    this.#requestedFrozenColumns = options.freeze?.columns ?? 0;
    this.#overscanRows = options.overscanRows ?? 0;
    this.#overscanColumns = options.overscanColumns ?? 0;
    const columnAxis = this.#createColumnAxis(options.columnWidths, new Map());
    this.#root = this.#createRoot(
      0,
      options.rowCount,
      RowMetrics.create(
        options.rowCount,
        options.rowHeight,
        options.rowHeights ?? new Map(),
      ),
      new Map(
        Array.from(options.rowHeights?.keys() ?? [], (row) => [row, 'manual' as const]),
      ),
      columnAxis,
    );
  }

  get rowCount(): number {
    return this.#root.rowCount;
  }

  get columnCount(): number {
    return this.#root.columnCount;
  }

  /** Default height for rows without an explicit override — not necessarily any particular row's actual height, see getRowHeight(). */
  get rowHeight(): number {
    return this.#rowHeight;
  }

  get frozenRowCount(): number {
    return this.#root.frozenRowCount;
  }

  get frozenColumnCount(): number {
    return this.#root.frozenColumnCount;
  }

  /** Bumped by exactly 1 on every successful setColumnWidth()/setRowCount() call; starts at 0. */
  get version(): number {
    return this.#root.version;
  }

  getColumnWidth(column: number): number {
    return this.#root.columnAxis.getColumnWidth(column);
  }

  /** Width to restore if the column is hidden, otherwise its current width. */
  getColumnRestoreWidth(column: number): number {
    return this.#root.columnAxis.getColumnRestoreWidth(column);
  }

  getColumnWidthsSnapshot(): readonly number[] {
    return Object.freeze(this.#root.columnAxis.copyColumnWidths());
  }

  /** Same x a row-0 getCellRect() would report, without needing a row index — for header code that has no row. */
  getColumnLeft(column: number): number {
    return this.#root.columnAxis.getColumnOffset(column);
  }

  /** `row`'s own height — the override if one was set via setRowHeight(), otherwise the default `rowHeight`. Out-of-range returns 0, mirroring getColumnWidth(). */
  getRowHeight(row: number): number {
    if (row < 0 || row >= this.#root.rowCount) {
      return 0;
    }
    return this.#root.rowMetrics.heightAt(row);
  }

  getRowHeightProvenance(row: number): RowHeightProvenance {
    if (!isNonNegativeInteger(row) || row >= this.#root.rowCount) {
      throw invalidLayoutOption('row', row);
    }
    return this.#root.rowHeightProvenance.get(row) ?? 'default';
  }

  /** Explicit O(rowCount) serialization boundary. Never use inside layout transactions. */
  exportDenseRowHeights(): Float64Array {
    return this.#root.rowMetrics.exportDenseValues();
  }

  /** Same y a column-0 getCellRect() would report, without needing a column index — for header/frozen-band code that has no column. `row === rowCount` is a valid "boundary past the last row" query, equal to getTotalHeight(). */
  getRowOffset(row: number): number {
    if (!isNonNegativeInteger(row) || row > this.#root.rowCount) {
      return 0;
    }
    return this.#root.rowMetrics.offsetAt(row);
  }

  /**
   * Resizes one column. Every later column's offset shifts by the width
   * delta (a column's absolute width still requires knowing everything to
   * its left, so a single-column change unavoidably moves everything after
   * it) — but that's a plain shift by a constant, not a full recompute from
   * scratch. `version` is the one thing callers should actually watch:
   * ColumnManager/HeaderManager's own cached RuntimeColumn/RuntimeHeader
   * arrays are untouched (this doesn't re-run their merge chains), only
   * `version` changes, so a render path that caches "have I painted this
   * version yet" knows to repaint without anything upstream re-resolving.
   */
  setColumnWidth(column: number, width: number): void {
    this.replaceRoot(this.prepareSetColumnWidth(column, width));
  }

  /**
   * Row-axis counterpart to setColumnWidth() — persistent path-copy instead
   * of column's O(n) offset shift, since row counts
   * can reach into the millions (see Step 20's spike, doc/progress.md).
   * Sets `row`'s height to exactly `height` — this is a plain overwrite, not
   * a floor/ceiling; a caller wanting "grow only if bigger" (e.g.
   * content-driven auto-height) does that comparison itself before calling.
   */
  setRowHeight(row: number, height: number): void {
    this.replaceRoot(this.prepareSetRowHeight(row, height));
  }

  /** Applies several explicit row heights with one dense snapshot/root build. */
  setRowHeights(updates: ReadonlyMap<number, LayoutRowHeightUpdate>): void {
    this.replaceRoot(this.prepareSetRowHeights(updates));
  }

  /**
   * Changes the persistent row rope: growing appends `rowHeight`-height
   * rows at the tail, shrinking truncates and **drops** any override at an
   * index `>= rowCount` (unlike hideColumn()/showColumn(), there is no
   * "remember and restore on regrow" here — a deliberate, documented
   * asymmetry, not an oversight; see LayoutEngine.test.ts's setRowCount
   * suite). Re-clamps `frozenRowCount` against the *originally requested*
   * freeze.rows (not the previous frozenRowCount), so a shrink that clips
   * frozen rows can recover them on a later grow.
   */
  setRowCount(rowCount: number): void {
    this.replaceRoot(this.prepareSetRowCount(rowCount));
  }

  /**
   * Splices `count` new rows in at `start` — mirrors insertColumns()'s
   * splice-then-rebuild shape, but the "widths" here (`heights`) are
   * optional and sparse: most inserted rows have no custom height, unlike a
   * new column which always needs an explicit width. `heights`' keys are
   * relative to the inserted block (`0..count-1`), same convention
   * `insertColumns(start, widths)` uses for its (dense) `widths` array.
   * Persistent split/concat keeps the cost proportional to tree depth and
   * the inserted fragment, rather than the total row count.
   */
  insertRows(
    start: number,
    count: number,
    heights?: ReadonlyMap<number, number>,
  ): void {
    this.replaceRoot(this.prepareInsertRows(start, count, heights));
  }

  /** Row-axis counterpart to deleteColumns() — a deleted row's height override goes with it; heights of rows after the deleted band shift down to fill the gap, via plain array splice (no bespoke index-remap logic needed). */
  deleteRows(start: number, count: number): void {
    this.replaceRoot(this.prepareDeleteRows(start, count));
  }

  /** `target` is the insertion index in the pre-move index space — same convention `moveColumns()`/`ArrayDataSource.moveRows()` use. Row count doesn't change, so (like moveColumns()) frozenRowCount doesn't need re-clamping. */
  moveRows(start: number, count: number, target: number): void {
    this.replaceRoot(this.prepareMoveRows(start, count, target));
  }

  /**
   * Splices `widths` in at `start` — unlike setRowCount() (a single shared
   * rowHeight, so structural row changes are just a count update),
   * per-column width means the immutable column axis genuinely changes.
   * That O(columnCount) rebuild happens once in the mutation; subsequent
   * Skeleton and Scene publications retain the resulting axis in O(1).
   * Re-clamps frozenColumnCount the same way setRowCount() does.
   */
  insertColumns(start: number, widths: readonly number[]): void {
    this.replaceRoot(this.prepareInsertColumns(start, widths));
  }

  deleteColumns(start: number, count: number): void {
    this.replaceRoot(this.prepareDeleteColumns(start, count));
  }

  /** `target` is the insertion index in the pre-move index space — same convention `remapColumnIndex()`/`ArrayDataSource.moveColumns()` use. */
  moveColumns(start: number, count: number, target: number): void {
    this.replaceRoot(this.prepareMoveColumns(start, count, target));
  }

  /**
   * Sets a column's width to 0 (this class's existing "hidden" convention —
   * see LayoutOptions.columnWidths) after remembering its current width, so
   * showColumn() can restore it exactly. A no-op on the remembered width if
   * the column is already hidden (hiding twice must not forget the real
   * original width and remember 0 instead).
   */
  hideColumn(column: number): void {
    this.replaceRoot(this.prepareHideColumn(column));
  }

  /** Restores a hidden column's remembered width. A no-op if the column isn't currently hidden. */
  showColumn(column: number): void {
    this.replaceRoot(this.prepareShowColumn(column));
  }

  isColumnHidden(column: number): boolean {
    return this.#root.columnAxis.isColumnHidden(column);
  }

  get root(): LayoutRoot {
    return this.#root;
  }

  snapshotRoot(): LayoutRootSnapshot {
    return Object.freeze({
      version: this.#root.version,
      rowCount: this.#root.rowCount,
      rowMetrics: this.#root.rowMetrics,
      rowHeightProvenance: this.#root.rowHeightProvenance,
      columnAxis: this.#root.columnAxis,
    });
  }

  prepareRoot(snapshot: LayoutRootSnapshot): PreparedLayoutRoot {
    return this.#createRoot(
      this.#root.version + 1,
      snapshot.rowCount,
      snapshot.rowMetrics,
      snapshot.rowHeightProvenance,
      snapshot.columnAxis,
    );
  }

  replaceRoot(root: PreparedLayoutRoot): void {
    this.#root = root;
  }

  prepareSetColumnWidth(column: number, width: number, base: LayoutRoot = this.#root): PreparedLayoutRoot {
    if (!isNonNegativeInteger(column) || column >= base.columnCount) {
      throw invalidLayoutOption('column', column);
    }
    if (!isNonNegativeFinite(width)) {
      throw invalidLayoutOption('width', width);
    }

    const columnWidths = base.columnAxis.copyColumnWidths();
    const hiddenColumnWidths = base.columnAxis.copyHiddenColumnWidths();
    if (hiddenColumnWidths.has(column)) {
      hiddenColumnWidths.set(column, width);
      columnWidths[column] = 0;
    } else {
      columnWidths[column] = width;
    }
    return this.#createRoot(
      base.version + 1,
      base.rowCount,
      base.rowMetrics,
      base.rowHeightProvenance,
      this.#createColumnAxis(columnWidths, hiddenColumnWidths),
    );
  }

  prepareSetRowHeight(
    row: number,
    height: number,
    base: LayoutRoot = this.#root,
  ): PreparedLayoutRoot {
    return this.prepareSetRowHeights(
      new Map([[row, Object.freeze({ height, provenance: 'manual' as const })]]),
      base,
    );
  }

  /** Prepares a batch without publishing a partially updated live root. */
  prepareSetRowHeights(
    updates: ReadonlyMap<number, LayoutRowHeightUpdate>,
    base: LayoutRoot = this.#root,
  ): PreparedLayoutRoot {
    let changed = false;
    const heights = new Map<number, number>();
    const rowHeightProvenance = new Map(base.rowHeightProvenance);
    for (const [row, update] of updates) {
      if (!isNonNegativeInteger(row) || row >= base.rowCount) {
        throw invalidLayoutOption('row', row);
      }
      if (!isPositiveFinite(update.height)) {
        throw invalidLayoutOption('height', update.height);
      }
      if (update.provenance !== 'default' && update.provenance !== 'auto' && update.provenance !== 'manual') {
        throw invalidLayoutOption('row height provenance', update.provenance);
      }
      heights.set(row, update.height);
      changed ||= base.rowMetrics.heightAt(row) !== update.height ||
        (base.rowHeightProvenance.get(row) ?? 'default') !== update.provenance;
      if (update.provenance === 'default') {
        rowHeightProvenance.delete(row);
      } else {
        rowHeightProvenance.set(row, update.provenance);
      }
    }
    if (updates.size === 0) {
      return base;
    }

    const rowMetrics = base.rowMetrics.withHeights(heights);
    if (!changed) {
      return base;
    }

    return this.#createRoot(
      base.version + 1,
      base.rowCount,
      rowMetrics,
      rowHeightProvenance,
      base.columnAxis,
    );
  }

  prepareSetRowCount(rowCount: number): PreparedLayoutRoot {
    if (!isNonNegativeInteger(rowCount)) {
      throw invalidLayoutOption('rowCount', rowCount);
    }

    const rowMetrics = rowCount < this.#root.rowCount
      ? this.#root.rowMetrics.delete(rowCount, this.#root.rowCount - rowCount).metrics
      : rowCount > this.#root.rowCount
        ? this.#root.rowMetrics.insert(
            this.#root.rowCount,
            RowMetrics.uniform(rowCount - this.#root.rowCount, this.#rowHeight),
          )
        : this.#root.rowMetrics;
    const rowHeightProvenance = new Map(
      Array.from(this.#root.rowHeightProvenance).filter(([row]) => row < rowCount),
    );
    return this.#createRoot(
      this.#root.version + 1,
      rowCount,
      rowMetrics,
      rowHeightProvenance,
      this.#root.columnAxis,
    );
  }

  prepareInsertRows(
    start: number,
    count: number,
    heights?: ReadonlyMap<number, number>,
  ): PreparedLayoutRoot {
    const rowHeightProvenance = heights === undefined
      ? undefined
      : new Map(Array.from(heights.keys(), (row) => [row, 'manual' as const]));
    return this.prepareStructureRoot(
      { axis: 'row', operation: 'insert', start, count },
      heights === undefined ? {} : { rowHeights: heights, rowHeightProvenance },
    );
  }

  prepareDeleteRows(start: number, count: number): PreparedLayoutRoot {
    return this.prepareStructureRoot({
      axis: 'row',
      operation: 'delete',
      start,
      count,
    });
  }

  prepareMoveRows(
    start: number,
    count: number,
    target: number,
  ): PreparedLayoutRoot {
    return this.prepareStructureRoot({
      axis: 'row',
      operation: 'move',
      start,
      count,
      target,
    });
  }

  prepareInsertColumns(
    start: number,
    widths: readonly number[],
  ): PreparedLayoutRoot {
    return this.prepareStructureRoot(
      { axis: 'column', operation: 'insert', start, count: widths.length },
      { columnWidths: widths },
    );
  }

  prepareDeleteColumns(start: number, count: number): PreparedLayoutRoot {
    return this.prepareStructureRoot({
      axis: 'column',
      operation: 'delete',
      start,
      count,
    });
  }

  prepareMoveColumns(
    start: number,
    count: number,
    target: number,
  ): PreparedLayoutRoot {
    return this.prepareStructureRoot({
      axis: 'column',
      operation: 'move',
      start,
      count,
      target,
    });
  }

  prepareHideColumn(column: number, base: LayoutRoot = this.#root): PreparedLayoutRoot {
    if (!isNonNegativeInteger(column) || column >= base.columnCount) {
      throw invalidLayoutOption('column', column);
    }
    const columnWidths = base.columnAxis.copyColumnWidths();
    const hiddenColumnWidths = base.columnAxis.copyHiddenColumnWidths();
    if (!hiddenColumnWidths.has(column)) {
      hiddenColumnWidths.set(column, columnWidths[column] as number);
    }
    columnWidths[column] = 0;
    return this.#createRoot(
      base.version + 1,
      base.rowCount,
      base.rowMetrics,
      base.rowHeightProvenance,
      this.#createColumnAxis(columnWidths, hiddenColumnWidths),
    );
  }

  prepareShowColumn(column: number, base: LayoutRoot = this.#root): PreparedLayoutRoot {
    if (!isNonNegativeInteger(column) || column >= base.columnCount) {
      throw invalidLayoutOption('column', column);
    }
    const width = base.columnAxis.isColumnHidden(column)
      ? base.columnAxis.getColumnRestoreWidth(column)
      : undefined;
    if (width === undefined) {
      return this.#createRoot(
        base.version + 1,
        base.rowCount,
        base.rowMetrics,
        base.rowHeightProvenance,
        this.#createColumnAxis(
          base.columnAxis.copyColumnWidths(),
          base.columnAxis.copyHiddenColumnWidths(),
        ),
      );
    }
    const columnWidths = base.columnAxis.copyColumnWidths();
    const hiddenColumnWidths = base.columnAxis.copyHiddenColumnWidths();
    hiddenColumnWidths.delete(column);
    columnWidths[column] = width;
    return this.#createRoot(
      base.version + 1,
      base.rowCount,
      base.rowMetrics,
      base.rowHeightProvenance,
      this.#createColumnAxis(columnWidths, hiddenColumnWidths),
    );
  }

  prepareStructureRoot(
    change: AxisChangeDescriptor,
    payload: LayoutStructurePayload = {},
    base: LayoutRoot = this.#root,
  ): PreparedLayoutRoot {
    validateAxisChange(
      change,
      change.axis === 'row' ? base.rowCount : base.columnCount,
    );

    if (change.axis === 'row') {
      let rowMetrics: RowMetrics;
      const rowHeightProvenance = this.#remapRowHeightProvenance(
        base.rowHeightProvenance,
        change,
      );
      if (change.operation === 'insert') {
        const inserted = payload.rowMetricsFragment ?? RowMetrics.create(
          change.count,
          this.#rowHeight,
          payload.rowHeights ?? new Map(),
        );
        if (inserted.rowCount !== change.count) {
          throw invalidLayoutOption('rowMetricsFragment.rowCount', inserted.rowCount);
        }
        const insertedProvenance = payload.rowHeightProvenance ?? new Map(
          Array.from(payload.rowHeights?.keys() ?? [], (row) => [row, 'manual' as const]),
        );
        for (const [offset, provenance] of insertedProvenance) {
          if (!isNonNegativeInteger(offset) || offset >= change.count) {
            throw invalidLayoutOption('rowHeightProvenance key', offset);
          }
          if (provenance !== 'auto' && provenance !== 'manual') {
            throw invalidLayoutOption('rowHeightProvenance value', provenance);
          }
          rowHeightProvenance.set(change.start + offset, provenance);
        }
        rowMetrics = base.rowMetrics.insert(change.start, inserted);
      } else if (change.operation === 'delete') {
        rowMetrics = base.rowMetrics.delete(change.start, change.count).metrics;
      } else {
        rowMetrics = base.rowMetrics.move(change.start, change.count, change.target as number);
      }

      return this.#createRoot(
        base.version + 1,
        rowMetrics.rowCount,
        rowMetrics,
        rowHeightProvenance,
        base.columnAxis,
      );
    }

    const columnWidths = base.columnAxis.copyColumnWidths();
    let hiddenColumnWidths = base.columnAxis.copyHiddenColumnWidths();

    if (change.operation === 'insert') {
      const widths = payload.columnWidths;
      if (widths === undefined || widths.length !== change.count) {
        throw invalidLayoutOption('columnWidths.length', widths?.length);
      }
      for (const width of widths) {
        if (!isNonNegativeFinite(width)) {
          throw invalidLayoutOption('width', width);
        }
      }
      columnWidths.splice(change.start, 0, ...widths);
      hiddenColumnWidths = this.#remapHiddenColumns(hiddenColumnWidths, {
        operation: 'insert',
        start: change.start,
        count: change.count,
      });
      for (const [offset, width] of payload.hiddenColumnOffsets ?? []) {
        if (!isNonNegativeInteger(offset) || offset >= change.count) {
          throw invalidLayoutOption('hiddenColumnOffsets key', offset);
        }
        if (!isNonNegativeFinite(width)) {
          throw invalidLayoutOption('hiddenColumnOffsets value', width);
        }
        const column = change.start + offset;
        hiddenColumnWidths.set(column, width);
        columnWidths[column] = 0;
      }
    } else if (change.operation === 'delete') {
      columnWidths.splice(change.start, change.count);
      hiddenColumnWidths = this.#remapHiddenColumns(hiddenColumnWidths, {
        operation: 'delete',
        start: change.start,
        count: change.count,
      });
    } else {
      const moved = columnWidths.splice(change.start, change.count);
      const target = change.target as number;
      const adjustedTarget =
        target > change.start ? target - change.count : target;
      columnWidths.splice(adjustedTarget, 0, ...moved);
      hiddenColumnWidths = this.#remapHiddenColumns(hiddenColumnWidths, {
        operation: 'move',
        start: change.start,
        count: change.count,
        target,
      });
    }

    return this.#createRoot(
      base.version + 1,
      base.rowCount,
      base.rowMetrics,
      base.rowHeightProvenance,
      this.#createColumnAxis(columnWidths, hiddenColumnWidths),
    );
  }

  #createRoot(
    version: number,
    rowCount: number,
    rowMetrics: RowMetrics,
    rowHeightProvenance: ReadonlyMap<number, Exclude<RowHeightProvenance, 'default'>>,
    columnAxis: ColumnAxisSnapshot,
  ): LayoutRoot {
    if (!isNonNegativeInteger(version)) {
      throw invalidLayoutOption('version', version);
    }
    if (!isNonNegativeInteger(rowCount)) {
      throw invalidLayoutOption('rowCount', rowCount);
    }
    if (rowMetrics.rowCount !== rowCount) {
      throw invalidLayoutOption('rowMetrics.rowCount', rowMetrics.rowCount);
    }
    if (!(columnAxis instanceof ColumnAxisSnapshot)) {
      throw invalidLayoutOption('columnAxis', columnAxis);
    }
    if (
      columnAxis.frozenColumnCount !==
      Math.min(this.#requestedFrozenColumns, columnAxis.columnCount)
    ) {
      throw invalidLayoutOption('columnAxis.frozenColumnCount', columnAxis.frozenColumnCount);
    }

    const provenance = new Map<number, Exclude<RowHeightProvenance, 'default'>>();
    for (const [row, value] of rowHeightProvenance) {
      if (!isNonNegativeInteger(row) || row >= rowCount) {
        throw invalidLayoutOption('rowHeightProvenance key', row);
      }
      if (value !== 'auto' && value !== 'manual') {
        throw invalidLayoutOption('rowHeightProvenance value', value);
      }
      provenance.set(row, value);
    }
    return Object.freeze({
      version,
      rowCount,
      columnCount: columnAxis.columnCount,
      rowMetrics,
      rowHeightProvenance: createReadonlyMap(provenance),
      columnAxis,
      frozenRowCount: Math.min(this.#requestedFrozenRows, rowCount),
      frozenColumnCount: columnAxis.frozenColumnCount,
    });
  }

  #createColumnAxis(
    columnWidths: readonly number[],
    hiddenColumnWidths: ReadonlyMap<number, number>,
  ): ColumnAxisSnapshot {
    try {
      return new ColumnAxisSnapshot({
        columnWidths,
        hiddenColumnWidths,
        frozenColumnCount: Math.min(this.#requestedFrozenColumns, columnWidths.length),
      });
    } catch (error) {
      if (error instanceof RangeError) {
        throw invalidLayoutOption('columnAxis', error.message);
      }
      throw error;
    }
  }

  #remapRowHeightProvenance(
    provenance: ReadonlyMap<number, Exclude<RowHeightProvenance, 'default'>>,
    change: AxisChangeDescriptor,
  ): Map<number, Exclude<RowHeightProvenance, 'default'>> {
    const remapped = new Map<number, Exclude<RowHeightProvenance, 'default'>>();
    for (const [row, value] of provenance) {
      if (
        change.operation === 'delete' &&
        row >= change.start &&
        row < change.start + change.count
      ) {
        continue;
      }
      remapped.set(remapIndex(row, change), value);
    }
    return remapped;
  }

  #remapHiddenColumns(
    hiddenColumnWidths: ReadonlyMap<number, number>,
    change: ColumnStructureChange,
  ): Map<number, number> {
    const remapped = new Map<number, number>();
    for (const [column, width] of hiddenColumnWidths) {
      if (
        change.operation === 'delete' &&
        column >= change.start &&
        column < change.start + change.count
      ) {
        continue;
      }
      remapped.set(remapIndex(column, change), width);
    }
    return remapped;
  }

  getTotalWidth(): number {
    return this.#root.columnAxis.getTotalWidth();
  }

  getTotalHeight(): number {
    return this.#root.rowMetrics.totalHeight;
  }

  getCellRect(row: number, column: number): CellRect {
    return {
      x: this.#root.columnAxis.getColumnOffset(column),
      y: this.getRowOffset(row),
      width: this.getColumnWidth(column),
      height: this.getRowHeight(row),
    };
  }

  getRowAtY(y: number): number {
    if (this.#root.rowCount === 0) {
      return 0;
    }
    return clamp(this.#rowIndexFromOffset(y), 0, this.#root.rowCount - 1);
  }

  getColumnAtX(x: number): number {
    if (this.#root.columnCount === 0) {
      return 0;
    }
    return clamp(this.#columnIndexFromOffset(x), 0, this.#root.columnCount - 1);
  }

  getVisibleRowRange(viewport: ViewportState): IndexRange {
    const scrollableRowCount = this.#root.rowCount - this.#root.frozenRowCount;
    const frozenHeight = this.getRowOffset(this.#root.frozenRowCount);
    const scrollableHeight = Math.max(0, viewport.height - frozenHeight);

    if (scrollableRowCount <= 0 || scrollableHeight <= 0) {
      return { start: this.#root.frozenRowCount, count: 0 };
    }

    const scrollTop = Math.max(0, viewport.scrollTop);
    const contentTop = frozenHeight + scrollTop;
    const firstIndex = this.#rowIndexFromOffset(contentTop);
    const lastIndex = this.#rowIndexFromOffset(contentTop + scrollableHeight);

    const start = clamp(
      firstIndex - this.#overscanRows,
      this.#root.frozenRowCount,
      this.#root.rowCount - 1,
    );
    const end = clamp(
      lastIndex + this.#overscanRows,
      this.#root.frozenRowCount,
      this.#root.rowCount - 1,
    );

    return { start, count: end - start + 1 };
  }

  getVisibleColumnRange(viewport: ViewportState): IndexRange {
    const scrollableColumnCount =
      this.#root.columnCount - this.#root.frozenColumnCount;
    const frozenWidth = this.#frozenColumnWidth();
    const scrollableWidth = Math.max(0, viewport.width - frozenWidth);

    if (scrollableColumnCount <= 0 || scrollableWidth <= 0) {
      return { start: this.#root.frozenColumnCount, count: 0 };
    }

    const scrollLeft = Math.max(0, viewport.scrollLeft);
    const contentLeft = frozenWidth + scrollLeft;
    const contentRight = contentLeft + scrollableWidth;

    const firstIndex = clamp(
      this.#columnIndexFromOffset(contentLeft),
      this.#root.frozenColumnCount,
      this.#root.columnCount - 1,
    );
    const lastIndex = clamp(
      this.#columnIndexFromOffset(contentRight),
      this.#root.frozenColumnCount,
      this.#root.columnCount - 1,
    );

    const start = clamp(
      firstIndex - this.#overscanColumns,
      this.#root.frozenColumnCount,
      this.#root.columnCount - 1,
    );
    const end = clamp(
      lastIndex + this.#overscanColumns,
      this.#root.frozenColumnCount,
      this.#root.columnCount - 1,
    );

    return { start, count: end - start + 1 };
  }

  getCellAtPoint(
    point: { x: number; y: number },
    viewport: ViewportState,
  ): { row: number; column: number } | null {
    if (point.x < 0 || point.y < 0) {
      return null;
    }
    if (this.#root.rowCount === 0 || this.#root.columnCount === 0) {
      return null;
    }

    const frozenHeight = this.getRowOffset(this.#root.frozenRowCount);
    const frozenWidth = this.#frozenColumnWidth();

    let row: number;
    if (point.y < frozenHeight) {
      row = clamp(
        this.#rowIndexFromOffset(point.y),
        0,
        this.#root.frozenRowCount - 1,
      );
    } else {
      const scrollTop = Math.max(0, viewport.scrollTop);
      row = this.#rowIndexFromOffset(scrollTop + point.y);
      if (row >= this.#root.rowCount) {
        return null;
      }
    }

    let column: number;
    if (point.x < frozenWidth) {
      column = clamp(
        this.#columnIndexFromOffset(point.x),
        0,
        this.#root.frozenColumnCount - 1,
      );
    } else {
      const scrollLeft = Math.max(0, viewport.scrollLeft);
      column = this.#columnIndexFromOffset(scrollLeft + point.x);
      if (column >= this.#root.columnCount) {
        return null;
      }
    }

    return { row, column };
  }

  #rowIndexFromOffset(offset: number): number {
    return this.#root.rowMetrics.rowAtOffset(offset);
  }

  #columnIndexFromOffset(offset: number): number {
    return this.#root.columnAxis.getColumnAtOffset(offset);
  }

  #frozenColumnWidth(): number {
    return this.#root.columnAxis.getColumnOffset(this.#root.frozenColumnCount);
  }
}
