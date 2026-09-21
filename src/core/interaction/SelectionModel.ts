import { CyExcelError, CyExcelErrorCode } from "../errors";
import { remapColumnIndex, remapRowIndex, type ColumnStructureChange, type RowStructureChange } from "../layout";
import type { CellRange } from "../types";

export type SelectionKind = "cell" | "row" | "column" | "all";

export interface CellPosition {
  readonly row: number;
  readonly column: number;
}

export interface SelectionBounds {
  readonly rowCount: number;
  readonly columnCount: number;
}

export interface SelectionSnapshot {
  readonly kind: SelectionKind;
  readonly active: CellPosition;
  readonly anchor: CellPosition;
  /** The active (most recently touched, still extendable) region — always equal to `ranges.at(-1)`. */
  readonly range: CellRange;
  /**
   * Every selected region, non-active ones first, active region always
   * last. Non-active regions are frozen `CellRange` snapshots (see
   * `SelectionModel`'s class doc comment) — only the active region can
   * still be extended via `extendTo()`/`moveActiveTo(..., true)`.
   */
  readonly ranges: readonly CellRange[];
}

/** Immutable owner root swapped atomically by selection mutations. */
export interface SelectionRoot {
  readonly bounds: SelectionBounds;
  readonly kind: SelectionKind;
  readonly active: CellPosition;
  readonly anchor: CellPosition;
  readonly otherRanges: readonly CellRange[];
}

function isNonNegativeInteger(value: number): boolean {
  return Number.isInteger(value) && value >= 0;
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

function validateBounds(bounds: SelectionBounds): void {
  if (!isNonNegativeInteger(bounds.rowCount) || !isNonNegativeInteger(bounds.columnCount)) {
    throw new CyExcelError(
      CyExcelErrorCode.InvalidOptions,
      `Invalid SelectionBounds: rowCount=${bounds.rowCount}, columnCount=${bounds.columnCount}`
    );
  }
}

function computeRange(
  kind: SelectionKind,
  anchor: CellPosition,
  active: CellPosition,
  bounds: SelectionBounds
): CellRange {
  const maxRow = Math.max(0, bounds.rowCount - 1);
  const maxColumn = Math.max(0, bounds.columnCount - 1);

  switch (kind) {
    case "all":
      return { startRow: 0, startColumn: 0, endRow: maxRow, endColumn: maxColumn };
    case "row":
      return {
        startRow: Math.min(anchor.row, active.row),
        endRow: Math.max(anchor.row, active.row),
        startColumn: 0,
        endColumn: maxColumn
      };
    case "column":
      return {
        startRow: 0,
        endRow: maxRow,
        startColumn: Math.min(anchor.column, active.column),
        endColumn: Math.max(anchor.column, active.column)
      };
    case "cell":
      return {
        startRow: Math.min(anchor.row, active.row),
        endRow: Math.max(anchor.row, active.row),
        startColumn: Math.min(anchor.column, active.column),
        endColumn: Math.max(anchor.column, active.column)
      };
  }
}

function freezeRoot(root: SelectionRoot): SelectionRoot {
  return Object.freeze({
    bounds: Object.freeze({ ...root.bounds }),
    kind: root.kind,
    active: Object.freeze({ ...root.active }),
    anchor: Object.freeze({ ...root.anchor }),
    otherRanges: Object.freeze(root.otherRanges.map((range) => Object.freeze({ ...range })))
  });
}

function clampPosition(position: CellPosition, bounds: SelectionBounds): CellPosition {
  return {
    row: clamp(position.row, 0, Math.max(0, bounds.rowCount - 1)),
    column: clamp(position.column, 0, Math.max(0, bounds.columnCount - 1))
  };
}

function clampRange(range: CellRange, bounds: SelectionBounds): CellRange {
  const start = clampPosition({ row: range.startRow, column: range.startColumn }, bounds);
  const end = clampPosition({ row: range.endRow, column: range.endColumn }, bounds);
  return { startRow: start.row, startColumn: start.column, endRow: end.row, endColumn: end.column };
}

/**
 * Holds active cell, anchor cell, and selection kind for the *active*
 * region; that region's range is always derived, never stored directly.
 * Additional, non-active regions (Ctrl/Cmd+click multi-select, Milestone 2
 * Step 16) are kept as frozen `CellRange` snapshots in `#otherRanges` —
 * once a region stops being active it can no longer be extended via
 * `extendTo()`, matching Excel: Shift+drag only ever extends the most
 * recently touched region, earlier regions stay exactly where they were
 * added.
 */
export class SelectionModel {
  #root: SelectionRoot;

  /**
   * 选区变了就叫一声。工具栏（Step 25）的按钮高亮要跟着选区走，而"选区变了"这件事
   * 此前没有任何出口——`CyExcelCore` 里十六处调用 `invalidateSelection()` 只是让
   * 画布重绘，不构成一个可订阅的事实。放在这里而不是在那十六处各发一次，是因为
   * 这里是 `#root` 唯一能被改的地方，漏不掉。
   */
  #onChange: (() => void) | undefined;

  constructor(bounds: SelectionBounds, onChange?: () => void) {
    validateBounds(bounds);
    this.#onChange = onChange;
    this.#root = freezeRoot({
      bounds,
      kind: "cell",
      anchor: { row: 0, column: 0 },
      active: { row: 0, column: 0 },
      otherRanges: []
    });
  }

  get bounds(): SelectionBounds {
    return this.#root.bounds;
  }

  get root(): SelectionRoot {
    return this.#root;
  }

  /**
   * Updates the bounds a selection is clamped into — e.g. after a
   * structural row or column operation changes `rowCount`/`columnCount`.
   * Without `rowChange`/`columnChange`, the current active/anchor
   * positions are just clamped into the new bounds (fine for a plain
   * count change with no reordering). With `rowChange`/`columnChange`,
   * positions on that axis are first remapped via `remapRowIndex()`/
   * `remapColumnIndex()` so a selection follows the *same logical row/
   * column* through an insert/delete/move instead of landing wherever
   * clamping happens to leave it — e.g. a selection on row 10 should still
   * be on row 10's data after two rows are inserted above it (now row 12),
   * not clamped back to row 10's new, different content. The two axes are
   * independent and can both apply in the same call (e.g. a paste that
   * simultaneously grows row and column count — not currently something
   * any Command does in one step, but the two remaps don't interact, so
   * there's no reason to forbid it).
   */
  setBounds(
    bounds: SelectionBounds,
    rowChange?: RowStructureChange,
    columnChange?: ColumnStructureChange
  ): void {
    this.replaceRoot(this.prepareBoundsRoot(bounds, rowChange, columnChange));
  }

  prepareBoundsRoot(
    bounds: SelectionBounds,
    rowChange?: RowStructureChange,
    columnChange?: ColumnStructureChange,
    base: SelectionRoot = this.#root
  ): SelectionRoot {
    validateBounds(bounds);
    let anchor = base.anchor;
    let active = base.active;
    let otherRanges = [...base.otherRanges];

    if (rowChange) {
      anchor = { ...anchor, row: remapRowIndex(anchor.row, rowChange) };
      active = { ...active, row: remapRowIndex(active.row, rowChange) };
      otherRanges = otherRanges.map((range) => ({
        ...range,
        startRow: remapRowIndex(range.startRow, rowChange),
        endRow: remapRowIndex(range.endRow, rowChange)
      }));
    }

    if (columnChange) {
      anchor = { ...anchor, column: remapColumnIndex(anchor.column, columnChange) };
      active = { ...active, column: remapColumnIndex(active.column, columnChange) };
      otherRanges = otherRanges.map((range) => ({
        ...range,
        startColumn: remapColumnIndex(range.startColumn, columnChange),
        endColumn: remapColumnIndex(range.endColumn, columnChange)
      }));
    }

    return freezeRoot({
      bounds,
      kind: base.kind,
      anchor: clampPosition(anchor, bounds),
      active: clampPosition(active, bounds),
      otherRanges: otherRanges.map((range) => clampRange(range, bounds))
    });
  }

  replaceRoot(root: SelectionRoot): void {
    this.#setRoot(root);
  }

  /** `#root` 的唯一写入口：只有真的换了对象才通知，省掉一串重复的空刷新。 */
  #setRoot(next: SelectionRoot): void {
    if (next === this.#root) {
      return;
    }
    this.#root = next;
    this.#onChange?.();
  }

  get snapshot(): SelectionSnapshot {
    const activeRange = Object.freeze(
      computeRange(this.#root.kind, this.#root.anchor, this.#root.active, this.#root.bounds)
    );
    return Object.freeze({
      kind: this.#root.kind,
      active: this.#root.active,
      anchor: this.#root.anchor,
      range: activeRange,
      ranges: Object.freeze([...this.#root.otherRanges, activeRange])
    });
  }

  selectCell(row: number, column: number): SelectionSnapshot {
    const position = this.#clampPosition(row, column);
    this.#replaceInteractionRoot("cell", position, position, []);
    return this.snapshot;
  }

  selectRow(row: number): SelectionSnapshot {
    const position = { row: clamp(row, 0, Math.max(0, this.#root.bounds.rowCount - 1)), column: 0 };
    this.#replaceInteractionRoot("row", position, position, []);
    return this.snapshot;
  }

  selectColumn(column: number): SelectionSnapshot {
    const position = {
      row: 0,
      column: clamp(column, 0, Math.max(0, this.#root.bounds.columnCount - 1))
    };
    this.#replaceInteractionRoot("column", position, position, []);
    return this.snapshot;
  }

  selectAll(): SelectionSnapshot {
    this.#replaceInteractionRoot("all", { row: 0, column: 0 }, {
      row: Math.max(0, this.#root.bounds.rowCount - 1),
      column: Math.max(0, this.#root.bounds.columnCount - 1)
    }, []);
    return this.snapshot;
  }

  /**
   * Freezes the current active region into `#otherRanges` and starts a new,
   * extendable single-cell active region at `(row, column)` — the Ctrl/
   * Cmd+click gesture (`PointerController`). Mirrors `selectCell()` for the
   * new region, just without clearing the regions that came before it.
   */
  addRange(row: number, column: number): SelectionSnapshot {
    const otherRanges = this.#frozenActiveRanges();
    const position = this.#clampPosition(row, column);
    this.#replaceInteractionRoot("cell", position, position, otherRanges);
    return this.snapshot;
  }

  /** Same as `addRange()`, but the new active region is a full-width row (Ctrl/Cmd+click on a row header). */
  addRow(row: number): SelectionSnapshot {
    const otherRanges = this.#frozenActiveRanges();
    const position = { row: clamp(row, 0, Math.max(0, this.#root.bounds.rowCount - 1)), column: 0 };
    this.#replaceInteractionRoot("row", position, position, otherRanges);
    return this.snapshot;
  }

  /** Same as `addRange()`, but the new active region is a full-height column (Ctrl/Cmd+click on a column header). */
  addColumn(column: number): SelectionSnapshot {
    const otherRanges = this.#frozenActiveRanges();
    const position = {
      row: 0,
      column: clamp(column, 0, Math.max(0, this.#root.bounds.columnCount - 1))
    };
    this.#replaceInteractionRoot("column", position, position, otherRanges);
    return this.snapshot;
  }

  extendTo(row: number, column: number): SelectionSnapshot {
    if (this.#root.kind !== "all") {
      this.#replaceInteractionRoot(
        this.#root.kind,
        this.#root.anchor,
        this.#clampPosition(row, column),
        this.#root.otherRanges
      );
    }
    return this.snapshot;
  }

  moveActive(deltaRow: number, deltaColumn: number, extend = false): SelectionSnapshot {
    return this.moveActiveTo(this.#root.active.row + deltaRow, this.#root.active.column + deltaColumn, extend);
  }

  moveActiveTo(row: number, column: number, extend = false): SelectionSnapshot {
    const position = this.#clampPosition(row, column);
    this.#replaceInteractionRoot(
      "cell",
      extend ? this.#root.anchor : position,
      position,
      extend ? this.#root.otherRanges : []
    );
    return this.snapshot;
  }

  #clampPosition(row: number, column: number): CellPosition {
    return clampPosition({ row, column }, this.#root.bounds);
  }

  #frozenActiveRanges(): readonly CellRange[] {
    return [
      ...this.#root.otherRanges,
      computeRange(this.#root.kind, this.#root.anchor, this.#root.active, this.#root.bounds)
    ];
  }

  #replaceInteractionRoot(
    kind: SelectionKind,
    anchor: CellPosition,
    active: CellPosition,
    otherRanges: readonly CellRange[]
  ): void {
    this.#setRoot(freezeRoot({ bounds: this.#root.bounds, kind, anchor, active, otherRanges }));
  }
}
