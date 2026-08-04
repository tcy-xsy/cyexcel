import { CyExcelError, CyExcelErrorCode } from "../errors";
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
  readonly range: CellRange;
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

/**
 * Holds active cell, anchor cell, and selection kind; range is always
 * derived, never stored directly. P0 supports one continuous range only —
 * a multi-region `ranges` array is left for a later P1 step.
 */
export class SelectionModel {
  readonly #bounds: SelectionBounds;
  #kind: SelectionKind = "cell";
  #anchor: CellPosition = { row: 0, column: 0 };
  #active: CellPosition = { row: 0, column: 0 };

  constructor(bounds: SelectionBounds) {
    validateBounds(bounds);
    this.#bounds = bounds;
  }

  get bounds(): SelectionBounds {
    return this.#bounds;
  }

  get snapshot(): SelectionSnapshot {
    return Object.freeze({
      kind: this.#kind,
      active: this.#active,
      anchor: this.#anchor,
      range: computeRange(this.#kind, this.#anchor, this.#active, this.#bounds)
    });
  }

  selectCell(row: number, column: number): SelectionSnapshot {
    const position = this.#clampPosition(row, column);
    this.#kind = "cell";
    this.#anchor = position;
    this.#active = position;
    return this.snapshot;
  }

  selectRow(row: number): SelectionSnapshot {
    const position = { row: clamp(row, 0, Math.max(0, this.#bounds.rowCount - 1)), column: 0 };
    this.#kind = "row";
    this.#anchor = position;
    this.#active = position;
    return this.snapshot;
  }

  selectColumn(column: number): SelectionSnapshot {
    const position = {
      row: 0,
      column: clamp(column, 0, Math.max(0, this.#bounds.columnCount - 1))
    };
    this.#kind = "column";
    this.#anchor = position;
    this.#active = position;
    return this.snapshot;
  }

  selectAll(): SelectionSnapshot {
    this.#kind = "all";
    this.#anchor = { row: 0, column: 0 };
    this.#active = {
      row: Math.max(0, this.#bounds.rowCount - 1),
      column: Math.max(0, this.#bounds.columnCount - 1)
    };
    return this.snapshot;
  }

  extendTo(row: number, column: number): SelectionSnapshot {
    if (this.#kind !== "all") {
      this.#active = this.#clampPosition(row, column);
    }
    return this.snapshot;
  }

  moveActive(deltaRow: number, deltaColumn: number, extend = false): SelectionSnapshot {
    return this.moveActiveTo(this.#active.row + deltaRow, this.#active.column + deltaColumn, extend);
  }

  moveActiveTo(row: number, column: number, extend = false): SelectionSnapshot {
    const position = this.#clampPosition(row, column);
    this.#kind = "cell";
    this.#active = position;
    if (!extend) {
      this.#anchor = position;
    }
    return this.snapshot;
  }

  #clampPosition(row: number, column: number): CellPosition {
    return {
      row: clamp(row, 0, Math.max(0, this.#bounds.rowCount - 1)),
      column: clamp(column, 0, Math.max(0, this.#bounds.columnCount - 1))
    };
  }
}
