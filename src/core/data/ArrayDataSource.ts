import { CyExcelError, CyExcelErrorCode } from "../errors";
import { applyRowOrder, assertValidRowOrder, invertRowOrder } from "../layout/rowOrder";
import { validateAxisChange } from "../layout/structuralChange";
import type { ModelRevision } from "../model";
import type {
  CellDescriptorSnapshot,
  DataMutation,
  DataSource,
  DataSourceTransaction,
  PreparedDataSourceMutation,
  RowData
} from "../types";

type StagedRow = number | RowData;

type RowPiece =
  | { readonly kind: "base"; readonly start: number; readonly length: number }
  | { readonly kind: "inserted"; readonly rows: readonly RowData[] };

export interface ArrayDataSourceColumn {
  readonly id: string;
  readonly field: string;
}

export interface ArrayDataSourceOptions {
  readonly readonly?: boolean;
}

/**
 * DataSource backed by a plain array of row objects. Rows and cell values
 * are kept as-is — no deep clone, no Proxy — per ADR 0001 §5 (complex
 * values are atomic).
 */
export class ArrayDataSource implements DataSource {
  readonly #rows: RowData[];
  /** Mutable (insertColumns/deleteColumns/moveColumns splice this in place) — shallow-copied from the constructor argument so structural changes never resize/reorder a caller-owned array out from under them, same reasoning as CyExcelCore.ts's buildDataSource() protecting the `data` array. */
  #columns: ArrayDataSourceColumn[];
  readonly #readonly: boolean;

  constructor(
    rows: RowData[],
    columns: readonly ArrayDataSourceColumn[],
    options: ArrayDataSourceOptions = {}
  ) {
    // Structural edits own their index array, while cell edits deliberately
    // keep the caller's row objects live, matching Univer's shared model.
    this.#rows = [...rows];
    this.#columns = [...columns];
    this.#readonly = options.readonly ?? false;
  }

  getRowCount(): number {
    return this.#rows.length;
  }

  getColumnCount(): number {
    return this.#columns.length;
  }

  getCell(row: number, column: number): unknown {
    const field = this.#columns[column]?.field;
    if (field === undefined) {
      return undefined;
    }
    return this.#rows[row]?.[field];
  }

  setCell(row: number, column: number, value: unknown): void {
    if (this.#readonly) {
      throw new CyExcelError(
        CyExcelErrorCode.ReadonlyDataSource,
        "Cannot write to a readonly ArrayDataSource"
      );
    }

    const field = this.#columns[column]?.field;
    if (field === undefined) {
      throw new CyExcelError(
        CyExcelErrorCode.ColumnNotFound,
        `Column ${column} does not exist`
      );
    }

    const rowData = this.#rows[row];
    if (rowData === undefined) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidCoordinate,
        `Row ${row} does not exist`
      );
    }

    rowData[field] = value;
  }

  /**
   * `rows` is one array of column-ordered values per new row (same
   * convention as `setRange()`) — turned into `RowData` objects here,
   * using this instance's own field mapping, before splicing them in.
   * `splice()` shifts the existing array in place; it never allocates a
   * full copy of `#rows`, so this stays cheap regardless of total row
   * count (Step 12's "不整表复制" exit criterion).
   */
  insertRows(start: number, rows: readonly (readonly unknown[])[]): void {
    if (this.#readonly) {
      throw new CyExcelError(
        CyExcelErrorCode.ReadonlyDataSource,
        "Cannot write to a readonly ArrayDataSource"
      );
    }
    validateAxisChange({ axis: "row", operation: "insert", start, count: rows.length }, this.#rows.length);

    const rowObjects = rows.map((values) => {
      const rowData: RowData = {};
      this.#columns.forEach((column, index) => {
        rowData[column.field] = values[index];
      });
      return rowData;
    });
    this.#insertItems(this.#rows, start, rowObjects);
  }

  deleteRows(start: number, count: number): void {
    if (this.#readonly) {
      throw new CyExcelError(
        CyExcelErrorCode.ReadonlyDataSource,
        "Cannot write to a readonly ArrayDataSource"
      );
    }
    validateAxisChange({ axis: "row", operation: "delete", start, count }, this.#rows.length);

    this.#rows.splice(start, count);
  }

  /**
   * `target` follows the same convention `LayoutEngine`/`SelectionModel`
   * assume: the insertion index in the *original* (pre-removal) index
   * space. The moved rows keep their exact object references — nothing
   * about "which row is this" changes, only position (see
   * doc/progress.md's Step 12 entry for why that's enough to satisfy
   * "stable identity" without a dedicated row-ID API).
   */
  moveRows(start: number, count: number, target: number): void {
    if (this.#readonly) {
      throw new CyExcelError(
        CyExcelErrorCode.ReadonlyDataSource,
        "Cannot write to a readonly ArrayDataSource"
      );
    }
    validateAxisChange({ axis: "row", operation: "move", start, count, target }, this.#rows.length);

    const moved = this.#rows.splice(start, count);
    const adjustedTarget = target > start ? target - count : target;
    this.#insertItems(this.#rows, adjustedTarget, moved);
  }

  reorderRows(order: Int32Array): void {
    if (this.#readonly) {
      throw new CyExcelError(
        CyExcelErrorCode.ReadonlyDataSource,
        "Cannot write to a readonly ArrayDataSource"
      );
    }
    assertValidRowOrder(order, this.#rows.length);
    const next = applyRowOrder(this.#rows, order);
    for (let index = 0; index < next.length; index += 1) {
      this.#rows[index] = next[index] as RowData;
    }
  }

  /**
   * Unlike insertRows(), never touches `#rows` — cell values live in each
   * row object keyed by `field` (see ADR 0001 §3), not by column position,
   * so a newly inserted column simply reads back as `undefined` on every
   * existing row without any of them needing to change. This makes column
   * structural operations O(columnCount) (typically tens), never
   * O(rowCount) the way the row equivalents necessarily are.
   */
  insertColumns(start: number, columns: readonly ArrayDataSourceColumn[]): void {
    if (this.#readonly) {
      throw new CyExcelError(
        CyExcelErrorCode.ReadonlyDataSource,
        "Cannot write to a readonly ArrayDataSource"
      );
    }
    validateAxisChange({ axis: "column", operation: "insert", start, count: columns.length }, this.#columns.length);

    this.#insertItems(this.#columns, start, columns);
  }

  /**
   * Only removes the column *definition* — the field's actual data is
   * still sitting on every row object, just unreachable through
   * `getCell()` until something maps that field name again (e.g. undoing
   * this by re-inserting the same column definition). Deliberately not
   * cleaned up: matches this DataSource's "atomic values, no proactive
   * deep manipulation" ADR 0001 §5 philosophy, and means undo needs no
   * captured snapshot the way DeleteRowsCommand's does.
   */
  deleteColumns(start: number, count: number): void {
    if (this.#readonly) {
      throw new CyExcelError(
        CyExcelErrorCode.ReadonlyDataSource,
        "Cannot write to a readonly ArrayDataSource"
      );
    }
    validateAxisChange({ axis: "column", operation: "delete", start, count }, this.#columns.length);

    this.#columns.splice(start, count);
  }

  /** Same pre-removal-index-space `target` convention as moveRows(). Data follows automatically (field-keyed, not position-keyed) — no row object is touched. */
  moveColumns(start: number, count: number, target: number): void {
    if (this.#readonly) {
      throw new CyExcelError(
        CyExcelErrorCode.ReadonlyDataSource,
        "Cannot write to a readonly ArrayDataSource"
      );
    }
    validateAxisChange({ axis: "column", operation: "move", start, count, target }, this.#columns.length);

    const moved = this.#columns.splice(start, count);
    const adjustedTarget = target > start ? target - count : target;
    this.#insertItems(this.#columns, adjustedTarget, moved);
  }

  beginTransaction(baseRevision: ModelRevision): DataSourceTransaction {
    const rowPieces: RowPiece[] = this.#rows.length === 0
      ? []
      : [{ kind: "base", start: 0, length: this.#rows.length }];
    const columns: ArrayDataSourceColumn[] = [...this.#columns];
    const cellOverlay = new Map<StagedRow, Map<string, PropertyDescriptor | undefined>>();
    let rowCount = this.#rows.length;
    let disposed = false;

    const assertActive = () => {
      if (disposed) {
        throw new CyExcelError(CyExcelErrorCode.MutationConflict, "DataSource transaction is disposed");
      }
    };
    const rowObject = (token: StagedRow): RowData =>
      typeof token === "number" ? this.#rows[token] as RowData : token;
    const rowTokenAt = (row: number): StagedRow => {
      if (!Number.isInteger(row) || row < 0 || row >= rowCount) {
        throw new CyExcelError(CyExcelErrorCode.InvalidCoordinate, `Row ${row} does not exist`);
      }
      let offset = row;
      for (const piece of rowPieces) {
        const length = piece.kind === "base" ? piece.length : piece.rows.length;
        if (offset < length) {
          return piece.kind === "base" ? piece.start + offset : piece.rows[offset] as RowData;
        }
        offset -= length;
      }
      throw new CyExcelError(CyExcelErrorCode.InvalidCoordinate, `Row ${row} does not exist`);
    };
    const fieldAt = (column: number): string => {
      if (!Number.isInteger(column) || column < 0 || column >= columns.length) {
        throw new CyExcelError(CyExcelErrorCode.ColumnNotFound, `Column ${column} does not exist`);
      }
      return (columns[column] as ArrayDataSourceColumn).field;
    };
    const descriptorAt = (row: number, column: number): PropertyDescriptor | undefined => {
      const token = rowTokenAt(row);
      const field = fieldAt(column);
      const overlay = cellOverlay.get(token);
      if (overlay?.has(field)) {
        return overlay.get(field);
      }
      return Object.getOwnPropertyDescriptor(rowObject(token), field);
    };
    const valueAt = (row: number, column: number): unknown => {
      const token = rowTokenAt(row);
      const field = fieldAt(column);
      const overlay = cellOverlay.get(token);
      if (overlay?.has(field)) {
        const descriptor = overlay.get(field);
        if (!descriptor) {
          return undefined;
        }
        return "value" in descriptor ? descriptor.value : descriptor.get?.call(rowObject(token));
      }
      return rowObject(token)[field];
    };
    const descriptorForValue = (row: number, column: number, value: unknown): PropertyDescriptor => {
      const token = rowTokenAt(row);
      const current = descriptorAt(row, column);
      if (current && (!("value" in current) || current.writable === false)) {
        throw new CyExcelError(
          CyExcelErrorCode.MutationNotSupported,
          `Cell ${row}:${column} is backed by an accessor or non-writable property`
        );
      }
      if (!current && !Object.isExtensible(rowObject(token))) {
        throw new CyExcelError(CyExcelErrorCode.MutationNotSupported, `Row ${row} is not extensible`);
      }
      return current
        ? { ...current, value }
        : { value, writable: true, enumerable: true, configurable: true };
    };
    const setOverlayDescriptor = (row: number, column: number, descriptor: PropertyDescriptor | undefined) => {
      const token = rowTokenAt(row);
      const field = fieldAt(column);
      let overlay = cellOverlay.get(token);
      if (!overlay) {
        overlay = new Map();
        cellOverlay.set(token, overlay);
      }
      overlay.set(field, descriptor);
    };
    const splitRowsAt = (position: number): number => {
      if (position === rowCount) {
        return rowPieces.length;
      }
      let offset = position;
      for (let index = 0; index < rowPieces.length; index += 1) {
        const piece = rowPieces[index] as RowPiece;
        const length = piece.kind === "base" ? piece.length : piece.rows.length;
        if (offset === 0) {
          return index;
        }
        if (offset < length) {
          const left: RowPiece = piece.kind === "base"
            ? { kind: "base", start: piece.start, length: offset }
            : { kind: "inserted", rows: piece.rows.slice(0, offset) };
          const right: RowPiece = piece.kind === "base"
            ? { kind: "base", start: piece.start + offset, length: piece.length - offset }
            : { kind: "inserted", rows: piece.rows.slice(offset) };
          rowPieces.splice(index, 1, left, right);
          return index + 1;
        }
        offset -= length;
      }
      return rowPieces.length;
    };
    const insertRowPiece = (start: number, rows: readonly RowData[]) => {
      if (rows.length === 0) {
        return;
      }
      const index = splitRowsAt(start);
      rowPieces.splice(index, 0, { kind: "inserted", rows });
      rowCount += rows.length;
    };
    const deleteRowPieces = (start: number, count: number): RowData[] => {
      const startIndex = splitRowsAt(start);
      const endIndex = splitRowsAt(start + count);
      const deleted = rowPieces.splice(startIndex, endIndex - startIndex);
      rowCount -= count;
      const rows: RowData[] = [];
      for (const piece of deleted) {
        if (piece.kind === "base") {
          for (let index = 0; index < piece.length; index += 1) {
            rows.push(this.#rows[piece.start + index] as RowData);
          }
        } else {
          rows.push(...piece.rows);
        }
      }
      return rows;
    };
    const moveRowPieces = (start: number, count: number, target: number) => {
      const startIndex = splitRowsAt(start);
      const endIndex = splitRowsAt(start + count);
      const moved = rowPieces.splice(startIndex, endIndex - startIndex);
      const adjustedTarget = target > start ? target - count : target;
      const targetIndex = splitRowsAt(adjustedTarget);
      rowPieces.splice(targetIndex, 0, ...moved);
    };
    /**
     * A reorder touches every row, so — unlike moveRowPieces — there is no
     * cheap segment-preserving edit; the whole piece list collapses into one
     * materialized `{ kind: "inserted" }` piece holding the permuted row
     * objects (Controller Ruling 4). Any in-flight cell overlay is rekeyed
     * along with its row: a "base" row's overlay is keyed by its numeric
     * `#rows` index, which stops being that row's token once it lands in an
     * "inserted" piece, so the overlay would otherwise become unreachable —
     * silently dropping a staged-but-uncommitted `setCell` that preceded the
     * reorder in the same transaction.
     */
    const reorderRowPieces = (order: Int32Array) => {
      const permuted: RowData[] = new Array(rowCount);
      for (let index = 0; index < order.length; index += 1) {
        const oldToken = rowTokenAt(order[index] as number);
        const row = rowObject(oldToken);
        permuted[index] = row;
        if (typeof oldToken === "number") {
          const overlay = cellOverlay.get(oldToken);
          if (overlay) {
            cellOverlay.delete(oldToken);
            cellOverlay.set(row, overlay);
          }
        }
      }
      rowPieces.length = 0;
      rowPieces.push({ kind: "inserted", rows: permuted });
    };
    const prepared = (
      forward: DataMutation,
      inverse: DataMutation,
      estimatedBytes: number,
      apply: () => void,
      revert: () => void
    ): PreparedDataSourceMutation => {
      let commitStarted = false;
      let rolledBack = false;
      return {
        forward,
        inverse,
        estimatedBytes,
        commit: () => {
          commitStarted = true;
          apply();
        },
        rollback: () => {
          if (!commitStarted || rolledBack) {
            return;
          }
          rolledBack = true;
          revert();
        },
        dispose: () => {}
      };
    };

    const transaction: DataSourceTransaction = {
      baseRevision,
      getRowCount: () => rowCount,
      getColumnCount: () => columns.length,
      getCell: (row, column) => {
        assertActive();
        return valueAt(row, column);
      },
      prepareMutation: (mutation) => {
        assertActive();
        this.#assertWritable();
        switch (mutation.type) {
          case "set-cell": {
            const before: CellDescriptorSnapshot[] = [{
              row: mutation.row,
              column: mutation.column,
              descriptor: descriptorAt(mutation.row, mutation.column)
            }];
            const after: CellDescriptorSnapshot[] = [{
              row: mutation.row,
              column: mutation.column,
              descriptor: descriptorForValue(mutation.row, mutation.column, mutation.value)
            }];
            setOverlayDescriptor(mutation.row, mutation.column, after[0]?.descriptor);
            return prepared(
              mutation,
              { type: "restore-cells", cells: before },
              64,
              () => this.#applyCellDescriptors(after),
              () => this.#applyCellDescriptors(before)
            );
          }
          case "set-range": {
            const rowLength = mutation.range.endRow - mutation.range.startRow + 1;
            const columnLength = mutation.range.endColumn - mutation.range.startColumn + 1;
            validateAxisChange({ axis: "row", operation: "delete", start: mutation.range.startRow, count: rowLength }, rowCount);
            validateAxisChange({ axis: "column", operation: "delete", start: mutation.range.startColumn, count: columnLength }, columns.length);
            if (mutation.values.length !== rowLength || mutation.values.some((row) => row.length !== columnLength)) {
              throw new CyExcelError(CyExcelErrorCode.InvalidRange, `Range values must have shape ${rowLength}x${columnLength}`);
            }
            const before: CellDescriptorSnapshot[] = [];
            const after: CellDescriptorSnapshot[] = [];
            for (let row = mutation.range.startRow; row <= mutation.range.endRow; row += 1) {
              for (let column = mutation.range.startColumn; column <= mutation.range.endColumn; column += 1) {
                before.push({ row, column, descriptor: descriptorAt(row, column) });
                const descriptor = descriptorForValue(
                  row,
                  column,
                  mutation.values[row - mutation.range.startRow]?.[column - mutation.range.startColumn]
                );
                after.push({ row, column, descriptor });
                setOverlayDescriptor(row, column, descriptor);
              }
            }
            return prepared(
              mutation,
              { type: "restore-cells", cells: before },
              64 + before.length * 32,
              () => this.#applyCellDescriptors(after),
              () => this.#applyCellDescriptors(before)
            );
          }
          case "restore-cells": {
            const before = mutation.cells.map((cell) => ({
              row: cell.row,
              column: cell.column,
              descriptor: descriptorAt(cell.row, cell.column)
            }));
            const target = mutation.cells.map((cell) => ({ ...cell }));
            for (const cell of target) {
              setOverlayDescriptor(cell.row, cell.column, cell.descriptor);
            }
            return prepared(
              mutation,
              { type: "restore-cells", cells: before },
              64 + before.length * 32,
              () => this.#applyCellDescriptors(target),
              () => this.#applyCellDescriptors(before)
            );
          }
          case "insert-rows": {
            validateAxisChange({ axis: "row", operation: "insert", start: mutation.start, count: mutation.rows.length }, rowCount);
            const rows = mutation.rows.map((values) => {
              const row: RowData = {};
              columns.forEach((column, index) => {
                Object.defineProperty(row, column.field, {
                  value: values[index],
                  writable: true,
                  enumerable: true,
                  configurable: true
                });
              });
              return row;
            });
            insertRowPiece(mutation.start, rows);
            const forward: DataMutation = { type: "insert-row-objects", start: mutation.start, rows };
            return prepared(
              forward,
              { type: "delete-rows", start: mutation.start, count: rows.length },
              64 + rows.length * Math.max(1, columns.length) * 16,
              () => this.#insertItems(this.#rows, mutation.start, rows),
              () => { this.#rows.splice(mutation.start, rows.length); }
            );
          }
          case "insert-row-objects": {
            validateAxisChange({ axis: "row", operation: "insert", start: mutation.start, count: mutation.rows.length }, rowCount);
            const rows = [...mutation.rows];
            insertRowPiece(mutation.start, rows);
            const forward: DataMutation = { ...mutation, rows };
            return prepared(
              forward,
              { type: "delete-rows", start: mutation.start, count: rows.length },
              64 + rows.length * 16,
              () => this.#insertItems(this.#rows, mutation.start, rows),
              () => { this.#rows.splice(mutation.start, rows.length); }
            );
          }
          case "delete-rows": {
            validateAxisChange({ axis: "row", operation: "delete", start: mutation.start, count: mutation.count }, rowCount);
            const deleted = deleteRowPieces(mutation.start, mutation.count);
            return prepared(
              mutation,
              { type: "insert-row-objects", start: mutation.start, rows: deleted },
              64 + deleted.length * 16,
              () => { this.#rows.splice(mutation.start, mutation.count); },
              () => this.#insertItems(this.#rows, mutation.start, deleted)
            );
          }
          case "move-rows": {
            validateAxisChange({ axis: "row", operation: "move", start: mutation.start, count: mutation.count, target: mutation.target }, rowCount);
            moveRowPieces(mutation.start, mutation.count, mutation.target);
            const adjustedTarget = mutation.target > mutation.start ? mutation.target - mutation.count : mutation.target;
            const undoTarget = mutation.start > adjustedTarget ? mutation.start + mutation.count : mutation.start;
            return prepared(
              mutation,
              { type: "move-rows", start: adjustedTarget, count: mutation.count, target: undoTarget },
              64 + mutation.count * 16,
              () => this.moveRows(mutation.start, mutation.count, mutation.target),
              () => this.moveRows(adjustedTarget, mutation.count, undoTarget)
            );
          }
          case "reorder-rows": {
            // Defensively copied, like every sibling case copies its
            // caller-supplied payload (insert-columns' `{...column}`,
            // restore-cells' cell copies): a caller that reuses `order`
            // between prepareMutation() and commit() must not desync
            // apply()/the already-staged view from the precomputed inverse.
            const order = Int32Array.from(mutation.order);
            assertValidRowOrder(order, rowCount);
            const inverse = invertRowOrder(order);
            reorderRowPieces(order);
            return prepared(
              { type: "reorder-rows", order },
              { type: "reorder-rows", order: inverse },
              64 + order.length * 4,
              () => this.reorderRows(order),
              () => this.reorderRows(inverse)
            );
          }
          case "insert-columns": {
            validateAxisChange({ axis: "column", operation: "insert", start: mutation.start, count: mutation.columns.length }, columns.length);
            const inserted = mutation.columns.map((column) => ({ ...column }));
            this.#insertItems(columns, mutation.start, inserted);
            const forward: DataMutation = { ...mutation, columns: inserted };
            return prepared(
              forward,
              { type: "delete-columns", start: mutation.start, count: inserted.length },
              64 + inserted.length * 32,
              () => this.#insertItems(this.#columns, mutation.start, inserted),
              () => { this.#columns.splice(mutation.start, inserted.length); }
            );
          }
          case "delete-columns": {
            validateAxisChange({ axis: "column", operation: "delete", start: mutation.start, count: mutation.count }, columns.length);
            const deleted = columns.splice(mutation.start, mutation.count);
            return prepared(
              mutation,
              { type: "insert-columns", start: mutation.start, columns: deleted },
              64 + deleted.length * 32,
              () => { this.#columns.splice(mutation.start, mutation.count); },
              () => this.#insertItems(this.#columns, mutation.start, deleted)
            );
          }
          case "move-columns": {
            validateAxisChange({ axis: "column", operation: "move", start: mutation.start, count: mutation.count, target: mutation.target }, columns.length);
            this.#moveItems(columns, mutation.start, mutation.count, mutation.target);
            const adjustedTarget = mutation.target > mutation.start ? mutation.target - mutation.count : mutation.target;
            const undoTarget = mutation.start > adjustedTarget ? mutation.start + mutation.count : mutation.start;
            return prepared(
              mutation,
              { type: "move-columns", start: adjustedTarget, count: mutation.count, target: undoTarget },
              64 + mutation.count * 32,
              () => this.moveColumns(mutation.start, mutation.count, mutation.target),
              () => this.moveColumns(adjustedTarget, mutation.count, undoTarget)
            );
          }
        }
      },
      dispose: () => {
        disposed = true;
        rowPieces.length = 0;
        columns.length = 0;
        cellOverlay.clear();
      }
    };
    return transaction;
  }

  #applyCellDescriptors(cells: readonly CellDescriptorSnapshot[]): void {
    for (const cell of cells) {
      const row = this.#rows[cell.row] as RowData;
      const field = this.#columns[cell.column]?.field as string;
      if (cell.descriptor) {
        Object.defineProperty(row, field, cell.descriptor);
      } else if (!Reflect.deleteProperty(row, field)) {
        throw new CyExcelError(CyExcelErrorCode.MutationRuntimeFaulted, `Could not restore cell ${cell.row}:${cell.column}`);
      }
    }
  }

  #assertWritable(): void {
    if (this.#readonly) {
      throw new CyExcelError(CyExcelErrorCode.ReadonlyDataSource, "Cannot write to a readonly ArrayDataSource");
    }
  }

  #insertItems<T>(target: T[], start: number, items: readonly T[]): number {
    let inserted = 0;
    while (inserted < items.length) {
      target.splice(start + inserted, 0, items[inserted] as T);
      inserted += 1;
    }
    return inserted;
  }

  #moveItems<T>(target: T[], start: number, count: number, destination: number): void {
    const moved = target.splice(start, count);
    const adjustedTarget = destination > start ? destination - count : destination;
    this.#insertItems(target, adjustedTarget, moved);
  }

}
