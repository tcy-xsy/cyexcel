import { invertRowOrder } from "../../../src/core/layout/rowOrder";
import type { ModelRevision } from "../../../src/core/model";
import type {
  DataMutation,
  DataSource,
  DataSourceTransaction,
  PreparedDataSourceMutation
} from "../../../src/core/types";
import { MapCellMatrix } from "./MapCellMatrix";
import type { CellRecord } from "./types";

/**
 * ADR 0004 spike: a record matrix implementing the existing `DataSource`
 * interface.
 *
 * **Keeping the interface is a means, not a walk-back of ADR §3.2** — the ADR
 * removes the seam, but removing it means touching CyExcelCore. Borrowing it
 * here is the cheapest way to drive the real Command / Transaction / History /
 * Patch stack from a record matrix. None of the spike's conclusions depend on
 * the seam surviving.
 */

/** One entry of a `restore-records` payload, after narrowing `record: unknown`. */
interface RecordSnapshot {
  readonly row: number;
  readonly column: number;
  readonly record: CellRecord | undefined;
}

/**
 * Task 2 measured the object-nested record matrix at 35.77-35.86 bytes per
 * occupied cell (median 35.78, forced GC, 5M cells). Rounded up to 36: the
 * measurement includes the matrix's own per-cell container overhead, which a
 * history entry does *not* retain, so 36 is deliberately on the generous side
 * of the record's true retained cost.
 */
const RECORD_BYTES = 36;

/**
 * The `{ row, column, record }` wrapper plus its array slot. A three-slot
 * fast-mode JSObject under pointer compression is 12 bytes of header + 3 * 4
 * bytes of in-object slots = 24, plus 4 bytes for the array element.
 */
const CELL_ENTRY_BYTES = 28;

/** Mutation object + payload array header. Matches the repo-wide base of 64. */
const BASE_BYTES = 64;

function restoreBytes(cells: readonly RecordSnapshot[]): number {
  let total = BASE_BYTES;
  for (const cell of cells) {
    total += CELL_ENTRY_BYTES + (cell.record === undefined ? 0 : RECORD_BYTES);
  }
  return total;
}

function narrowRecordSnapshots(
  cells: readonly { readonly row: number; readonly column: number; readonly record: unknown }[]
): RecordSnapshot[] {
  return cells.map((cell) => ({
    row: cell.row,
    column: cell.column,
    record: cell.record as CellRecord | undefined
  }));
}

/**
 * A row permutation that may contain holes: `order[newRow] = oldRow`, or -1
 * for a row that did not exist before. Insert / delete / move / reorder all
 * collapse into this single primitive on a row-bucketed sparse matrix.
 */
function insertOrder(rowCount: number, start: number, count: number): Int32Array {
  const order = new Int32Array(rowCount + count);
  for (let row = 0; row < start; row += 1) {
    order[row] = row;
  }
  for (let offset = 0; offset < count; offset += 1) {
    order[start + offset] = -1;
  }
  for (let row = start; row < rowCount; row += 1) {
    order[row + count] = row;
  }
  return order;
}

function deleteOrder(rowCount: number, start: number, count: number): Int32Array {
  const order = new Int32Array(rowCount - count);
  for (let row = 0; row < start; row += 1) {
    order[row] = row;
  }
  for (let row = start + count; row < rowCount; row += 1) {
    order[row - count] = row;
  }
  return order;
}

function moveOrder(rowCount: number, start: number, count: number, target: number): Int32Array {
  const indices: number[] = Array.from({ length: rowCount }, (_unused, row) => row);
  const moved = indices.splice(start, count);
  const adjustedTarget = target > start ? target - count : target;
  indices.splice(adjustedTarget, 0, ...moved);
  return Int32Array.from(indices);
}

export class SpikeSheetDataSource implements DataSource {
  readonly #matrix: MapCellMatrix;
  #rowCount: number;
  #columnCount: number;

  constructor(rowCount: number, columnCount: number) {
    this.#matrix = new MapCellMatrix(rowCount, columnCount);
    this.#rowCount = rowCount;
    this.#columnCount = columnCount;
  }

  /** Spike-only direct write that bypasses Command, for staging test fixtures. */
  setRecord(row: number, column: number, record: CellRecord): void {
    this.#matrix.set(row, column, record);
  }

  getRecord(row: number, column: number): CellRecord | undefined {
    return this.#matrix.get(row, column);
  }

  occupiedCount(): number {
    return this.#matrix.occupiedCount();
  }

  getRowCount(): number {
    return this.#rowCount;
  }

  getColumnCount(): number {
    return this.#columnCount;
  }

  /** Invariant 1: always the bare value, never the record. */
  getCell(row: number, column: number): unknown {
    return this.#matrix.get(row, column)?.v;
  }

  setCell(row: number, column: number, value: unknown): void {
    this.#writeRecord(row, column, mergeValue(this.#matrix.get(row, column), value));
  }

  insertRows(start: number, rows: readonly (readonly unknown[])[]): void {
    this.#permute(insertOrder(this.#rowCount, start, rows.length), this.#rowCount + rows.length);
    rows.forEach((values, offset) => {
      values.forEach((value, column) => {
        if (value !== undefined) {
          this.#matrix.set(start + offset, column, { v: value });
        }
      });
    });
  }

  deleteRows(start: number, count: number): void {
    this.#permute(deleteOrder(this.#rowCount, start, count), this.#rowCount - count);
  }

  moveRows(start: number, count: number, target: number): void {
    this.#permute(moveOrder(this.#rowCount, start, count, target), this.#rowCount);
  }

  reorderRows(order: Int32Array): void {
    this.#permute(order, order.length);
  }

  /**
   * Column-axis structural ops. Note what this costs: the matrix is bucketed
   * by row, so shifting a column rekeys every occupied cell, and the Task 1
   * `CellMatrix` interface exposes no iteration primitive, so it degenerates
   * into a dense rowCount x columnCount scan. See the Task 4 report.
   */
  insertColumns(start: number, columns: readonly { readonly id: string; readonly field: string }[]): void {
    this.#shiftColumns(start, columns.length);
    this.#columnCount += columns.length;
  }

  deleteColumns(start: number, count: number): void {
    this.#shiftColumns(start, -count);
    this.#columnCount -= count;
  }

  moveColumns(start: number, count: number, target: number): void {
    const captured: (CellRecord | undefined)[][] = [];
    for (let row = 0; row < this.#rowCount; row += 1) {
      const cells: (CellRecord | undefined)[] = [];
      for (let column = 0; column < this.#columnCount; column += 1) {
        cells.push(this.#matrix.get(row, column));
      }
      const moved = cells.splice(start, count);
      const adjustedTarget = target > start ? target - count : target;
      cells.splice(adjustedTarget, 0, ...moved);
      captured.push(cells);
    }
    captured.forEach((cells, row) => {
      cells.forEach((record, column) => {
        this.#writeRecord(row, column, record);
      });
    });
  }

  beginTransaction(baseRevision: ModelRevision): DataSourceTransaction {
    /** Staged row -> live row, or -1 for a row the transaction invented. */
    let rowMap: number[] = Array.from({ length: this.#rowCount }, (_unused, row) => row);
    let stagedColumnCount = this.#columnCount;
    /** Staged cell writes; a stored `undefined` means "staged as deleted". */
    const overlay = new Map<string, CellRecord | undefined>();
    const key = (row: number, column: number): string => `${row}:${column}`;
    const recordAt = (row: number, column: number): CellRecord | undefined => {
      const k = key(row, column);
      if (overlay.has(k)) {
        return overlay.get(k);
      }
      const liveRow = rowMap[row];
      return liveRow === undefined || liveRow < 0 ? undefined : this.#matrix.get(liveRow, column);
    };

    /** Applies a hole-tolerant row permutation to the staged view only. */
    const restageRows = (order: Int32Array): void => {
      const newRowFor = new Map<number, number>();
      const nextRowMap: number[] = [];
      for (let newRow = 0; newRow < order.length; newRow += 1) {
        const oldRow = order[newRow] as number;
        nextRowMap.push(oldRow < 0 ? -1 : (rowMap[oldRow] ?? -1));
        if (oldRow >= 0) {
          newRowFor.set(oldRow, newRow);
        }
      }
      const staged = [...overlay.entries()];
      overlay.clear();
      for (const [entryKey, record] of staged) {
        const [rowPart, columnPart] = entryKey.split(":");
        const newRow = newRowFor.get(Number(rowPart));
        if (newRow !== undefined) {
          overlay.set(`${newRow}:${columnPart ?? "0"}`, record);
        }
      }
      rowMap = nextRowMap;
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

    const prepareRestore = (cells: readonly RecordSnapshot[]): PreparedDataSourceMutation => {
      const before: RecordSnapshot[] = cells.map((cell) => ({
        row: cell.row,
        column: cell.column,
        record: recordAt(cell.row, cell.column)
      }));
      const target = cells.map((cell) => ({ ...cell }));
      for (const cell of target) {
        overlay.set(key(cell.row, cell.column), cell.record);
      }
      return prepared(
        { type: "restore-records", cells: target },
        { type: "restore-records", cells: before },
        restoreBytes(before),
        () => this.#applyRecords(target),
        () => this.#applyRecords(before)
      );
    };

    return {
      baseRevision,
      getRowCount: () => rowMap.length,
      getColumnCount: () => stagedColumnCount,
      getCell: (row, column) => recordAt(row, column)?.v,
      prepareMutation: (mutation: DataMutation): PreparedDataSourceMutation => {
        switch (mutation.type) {
          case "set-cell": {
            const before = recordAt(mutation.row, mutation.column);
            const after = mergeValue(before, mutation.value) as CellRecord;
            overlay.set(key(mutation.row, mutation.column), after);
            const snapshot: RecordSnapshot[] = [
              { row: mutation.row, column: mutation.column, record: before }
            ];
            const applied: RecordSnapshot[] = [
              { row: mutation.row, column: mutation.column, record: after }
            ];
            return prepared(
              mutation,
              // Record-native inverse: put the whole record back. A `record`
              // of undefined means the cell was unoccupied and the inverse
              // must delete it, not write undefined into it.
              { type: "restore-records", cells: snapshot },
              restoreBytes(snapshot),
              () => this.#applyRecords(applied),
              () => this.#applyRecords(snapshot)
            );
          }
          case "set-range": {
            const before: RecordSnapshot[] = [];
            const after: RecordSnapshot[] = [];
            for (let row = mutation.range.startRow; row <= mutation.range.endRow; row += 1) {
              for (let column = mutation.range.startColumn; column <= mutation.range.endColumn; column += 1) {
                const existing = recordAt(row, column);
                const value = mutation.values[row - mutation.range.startRow]?.[column - mutation.range.startColumn];
                const next = mergeValue(existing, value);
                before.push({ row, column, record: existing });
                after.push({ row, column, record: next });
                overlay.set(key(row, column), next);
              }
            }
            return prepared(
              mutation,
              { type: "restore-records", cells: before },
              restoreBytes(before),
              () => this.#applyRecords(after),
              () => this.#applyRecords(before)
            );
          }
          case "restore-records":
            return prepareRestore(narrowRecordSnapshots(mutation.cells));
          case "restore-cells":
            // FINDING: `restore-cells` is ArrayDataSource's PropertyDescriptor
            // payload. A record matrix cannot honour it - there is no caller
            // object to redefine a property on. Retained here only because the
            // variant still exists in the union for ArrayDataSource's sake.
            throw new Error("spike: restore-cells is not expressible on a record matrix");
          case "insert-rows": {
            const count = mutation.rows.length;
            const order = insertOrder(rowMap.length, mutation.start, count);
            const nextRowCount = rowMap.length + count;
            restageRows(order);
            mutation.rows.forEach((values, offset) => {
              values.forEach((value, column) => {
                overlay.set(
                  key(mutation.start + offset, column),
                  value === undefined ? undefined : { v: value }
                );
              });
            });
            const rows = mutation.rows.map((values) => [...values]);
            return prepared(
              { type: "insert-rows", start: mutation.start, rows },
              { type: "delete-rows", start: mutation.start, count },
              // A blank inserted row occupies no cells; only the values the
              // caller supplied are chargeable.
              BASE_BYTES + rows.flat().filter((value) => value !== undefined).length * RECORD_BYTES,
              () => this.insertRows(mutation.start, rows),
              () => {
                this.#permute(deleteOrder(nextRowCount, mutation.start, count), nextRowCount - count);
              }
            );
          }
          case "insert-row-objects":
            // FINDING: RowData is a field-name-keyed caller object. A record
            // matrix is index-addressed and has no field names, so this
            // variant has no meaning here.
            throw new Error("spike: insert-row-objects is not expressible on a record matrix");
          case "delete-rows": {
            const captured: RecordSnapshot[] = [];
            for (let offset = 0; offset < mutation.count; offset += 1) {
              for (let column = 0; column < stagedColumnCount; column += 1) {
                const record = recordAt(mutation.start + offset, column);
                if (record !== undefined) {
                  captured.push({ row: offset, column, record });
                }
              }
            }
            const beforeRowCount = rowMap.length;
            restageRows(deleteOrder(beforeRowCount, mutation.start, mutation.count));
            // FINDING (Q2): the inverse has to be ONE DataMutation, and the
            // only row-inserting variants carry bare values or RowData. The
            // records' `t` / `s` are dropped here; rollback below keeps them
            // because it closes over the records instead.
            const restoredValues: unknown[][] = Array.from({ length: mutation.count }, () =>
              Array.from({ length: stagedColumnCount }, () => undefined as unknown)
            );
            for (const cell of captured) {
              (restoredValues[cell.row] as unknown[])[cell.column] = cell.record?.v;
            }
            return prepared(
              mutation,
              { type: "insert-rows", start: mutation.start, rows: restoredValues },
              restoreBytes(captured),
              () => this.deleteRows(mutation.start, mutation.count),
              () => {
                this.#permute(
                  insertOrder(beforeRowCount - mutation.count, mutation.start, mutation.count),
                  beforeRowCount
                );
                this.#applyRecords(
                  captured.map((cell) => ({ ...cell, row: mutation.start + cell.row }))
                );
              }
            );
          }
          case "move-rows": {
            restageRows(moveOrder(rowMap.length, mutation.start, mutation.count, mutation.target));
            const adjustedTarget = mutation.target > mutation.start ? mutation.target - mutation.count : mutation.target;
            const undoTarget = mutation.start > adjustedTarget ? mutation.start + mutation.count : mutation.start;
            return prepared(
              mutation,
              { type: "move-rows", start: adjustedTarget, count: mutation.count, target: undoTarget },
              BASE_BYTES,
              () => this.moveRows(mutation.start, mutation.count, mutation.target),
              () => this.moveRows(adjustedTarget, mutation.count, undoTarget)
            );
          }
          case "reorder-rows": {
            const order = Int32Array.from(mutation.order);
            const inverse = invertRowOrder(order);
            restageRows(order);
            return prepared(
              { type: "reorder-rows", order },
              { type: "reorder-rows", order: inverse },
              // The inverse is a permutation, not the moved records: O(rows),
              // independent of how many cells are occupied.
              BASE_BYTES + order.length * 4,
              () => this.reorderRows(order),
              () => this.reorderRows(inverse)
            );
          }
          case "insert-columns": {
            // FINDING: only `columns.length` is meaningful here. The
            // `{id, field}` mapping is ArrayDataSource's, not the matrix's.
            const count = mutation.columns.length;
            stagedColumnCount += count;
            return prepared(
              mutation,
              { type: "delete-columns", start: mutation.start, count },
              BASE_BYTES,
              () => this.insertColumns(mutation.start, mutation.columns),
              () => this.deleteColumns(mutation.start, count)
            );
          }
          case "delete-columns": {
            const captured: RecordSnapshot[] = [];
            for (let row = 0; row < rowMap.length; row += 1) {
              for (let offset = 0; offset < mutation.count; offset += 1) {
                const record = recordAt(row, mutation.start + offset);
                if (record !== undefined) {
                  captured.push({ row, column: offset, record });
                }
              }
            }
            stagedColumnCount -= mutation.count;
            // FINDING: same shape problem as delete-rows, plus the inverse
            // must invent `{id, field}` mappings the matrix does not own.
            const columns = Array.from({ length: mutation.count }, (_unused, offset) => ({
              id: `spike-column-${mutation.start + offset}`,
              field: `spike-column-${mutation.start + offset}`
            }));
            return prepared(
              mutation,
              { type: "insert-columns", start: mutation.start, columns },
              restoreBytes(captured),
              () => this.deleteColumns(mutation.start, mutation.count),
              () => {
                this.insertColumns(mutation.start, columns);
                this.#applyRecords(
                  captured.map((cell) => ({ ...cell, column: mutation.start + cell.column }))
                );
              }
            );
          }
          case "move-columns": {
            const adjustedTarget = mutation.target > mutation.start ? mutation.target - mutation.count : mutation.target;
            const undoTarget = mutation.start > adjustedTarget ? mutation.start + mutation.count : mutation.start;
            return prepared(
              mutation,
              { type: "move-columns", start: adjustedTarget, count: mutation.count, target: undoTarget },
              BASE_BYTES,
              () => this.moveColumns(mutation.start, mutation.count, mutation.target),
              () => this.moveColumns(adjustedTarget, mutation.count, undoTarget)
            );
          }
        }
      },
      dispose: () => {
        overlay.clear();
        rowMap = [];
      }
    };
  }

  #permute(order: Int32Array, nextRowCount: number): void {
    this.#matrix.reorderRows(order);
    this.#rowCount = nextRowCount;
  }

  #applyRecords(cells: readonly RecordSnapshot[]): void {
    for (const cell of cells) {
      this.#writeRecord(cell.row, cell.column, cell.record);
    }
  }

  #writeRecord(row: number, column: number, record: CellRecord | undefined): void {
    if (record === undefined) {
      this.#matrix.delete(row, column);
    } else {
      this.#matrix.set(row, column, record);
    }
  }

  #shiftColumns(start: number, delta: number): void {
    const captured: (CellRecord | undefined)[][] = [];
    for (let row = 0; row < this.#rowCount; row += 1) {
      const cells: (CellRecord | undefined)[] = [];
      for (let column = 0; column < this.#columnCount; column += 1) {
        cells.push(this.#matrix.get(row, column));
      }
      if (delta > 0) {
        cells.splice(start, 0, ...Array.from({ length: delta }, () => undefined));
      } else {
        cells.splice(start, -delta);
      }
      captured.push(cells);
    }
    const width = Math.max(this.#columnCount, this.#columnCount + delta);
    captured.forEach((cells, row) => {
      for (let column = 0; column < width; column += 1) {
        this.#writeRecord(row, column, cells[column]);
      }
    });
  }
}

/**
 * Setting a value must not throw away the rest of the record - that is the
 * whole point of ADR §3.1's "attributes travel with the cell".
 */
function mergeValue(existing: CellRecord | undefined, value: unknown): CellRecord | undefined {
  if (value === undefined && existing === undefined) {
    return undefined;
  }
  return existing === undefined ? { v: value } : { ...existing, v: value };
}
