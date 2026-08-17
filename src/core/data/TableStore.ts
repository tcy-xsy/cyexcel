import { CyExcelError, CyExcelErrorCode } from "../errors";
import { validateAxisChange } from "../layout/structuralChange";
import {
  MutationCoordinator,
  type CommitReceipt,
  type Mutation,
  type MutationPrepareContext,
  type MutationPrepareSession,
  type PreparedMutationParticipant
} from "../mutations";
import type { CellRange, ColumnFieldMapping, DataMutation, DataSource, Patch } from "../types";
import { assertValidCoordinate, assertValidRange } from "../validation";
import type { ChangeSet } from "./changes";

export interface TableStoreCommitResult {
  readonly patches: readonly Patch[];
  readonly changeSet: ChangeSet;
  readonly version: number;
}

class TableDataMutation implements Mutation {
  readonly type: string;
  readonly changes: readonly Patch[];
  readonly #source: DataSource;
  readonly #mutation: DataMutation;

  constructor(source: DataSource, mutation: DataMutation, changes: readonly Patch[]) {
    this.#source = source;
    this.#mutation = mutation;
    this.type = `data:${mutation.type}`;
    this.changes = Object.freeze([...changes]);
  }

  prepare(context: MutationPrepareContext): PreparedMutationParticipant {
    const prepared = context.prepareData(this.#source, this.#mutation);
    const forward = new TableDataMutation(
      this.#source,
      prepared.forward,
      changesForDataMutation(prepared.forward, this.changes)
    );
    const inverse = new TableDataMutation(
      this.#source,
      prepared.inverse,
      changesForDataMutation(prepared.inverse, this.changes)
    );
    return {
      hasEffect: true,
      forward,
      inverse,
      estimatedBytes: prepared.estimatedBytes,
      commit: () => prepared.commit(),
      rollback: () => prepared.rollback(),
      dispose: () => prepared.dispose()
    };
  }
}

function changesForDataMutation(mutation: DataMutation, fallback: readonly Patch[]): readonly Patch[] {
  switch (mutation.type) {
    case "set-cell":
      return [{ type: "cell-value", row: mutation.row, column: mutation.column }];
    case "set-range":
      return [{ type: "range-value", range: mutation.range }];
    case "restore-cells":
      return fallback;
    case "insert-rows":
    case "insert-row-objects":
      return [{ type: "row-structure", operation: "insert", start: mutation.start, count: mutation.rows.length }];
    case "delete-rows":
      return [{ type: "row-structure", operation: "delete", start: mutation.start, count: mutation.count }];
    case "move-rows":
      return [{ type: "row-structure", operation: "move", start: mutation.start, count: mutation.count, target: mutation.target }];
    case "insert-columns":
      return [{ type: "column-structure", operation: "insert", start: mutation.start, count: mutation.columns.length }];
    case "delete-columns":
      return [{ type: "column-structure", operation: "delete", start: mutation.start, count: mutation.count }];
    case "move-columns":
      return [{ type: "column-structure", operation: "move", start: mutation.start, count: mutation.count, target: mutation.target }];
  }
}

/** Mutable table model whose writes are finalized by one shared coordinator. */
export class TableStore {
  readonly #dataSource: DataSource;
  readonly #coordinator: MutationCoordinator;
  #nextDirectTransactionId = 1;

  constructor(dataSource: DataSource, coordinator = new MutationCoordinator()) {
    this.#dataSource = dataSource;
    this.#coordinator = coordinator;
  }

  get version(): number {
    return this.#coordinator.revisionGate.revision;
  }

  get coordinator(): MutationCoordinator {
    return this.#coordinator;
  }

  beginMutationSession(transactionId: string): MutationPrepareSession {
    return this.#coordinator.begin(transactionId, (baseRevision) => {
      const sourceTransaction = this.#dataSource.beginTransaction?.(baseRevision);
      if (!sourceTransaction) {
        throw new CyExcelError(
          CyExcelErrorCode.MutationNotSupported,
          "This DataSource does not implement the staged transaction contract"
        );
      }
      if (sourceTransaction.baseRevision !== baseRevision) {
        sourceTransaction.dispose();
        throw new CyExcelError(CyExcelErrorCode.MutationConflict, "DataSource transaction revision mismatch");
      }
      return {
        model: sourceTransaction,
        prepareData: (source, mutation) => {
          if (source !== this.#dataSource) {
            throw new CyExcelError(CyExcelErrorCode.MutationConflict, "Mutation targets another DataSource");
          }
          return sourceTransaction.prepareMutation(mutation);
        },
        dispose: () => sourceTransaction.dispose()
      };
    });
  }

  getRowCount(): number {
    return this.#dataSource.getRowCount();
  }

  getColumnCount(): number {
    return this.#dataSource.getColumnCount();
  }

  getCell(row: number, column: number): unknown {
    this.#assertCellExists(row, column);
    return this.#dataSource.getCell(row, column);
  }

  createSetCellMutation(row: number, column: number, value: unknown): Mutation {
    assertValidCoordinate(row, column);
    return new TableDataMutation(
      this.#dataSource,
      { type: "set-cell", row, column, value },
      [{ type: "cell-value", row, column }]
    );
  }

  createSetRangeMutation(range: CellRange, values: readonly (readonly unknown[])[]): Mutation {
    assertValidRange(range);
    const expectedRows = range.endRow - range.startRow + 1;
    const expectedColumns = range.endColumn - range.startColumn + 1;
    if (values.length !== expectedRows || values.some((row) => row.length !== expectedColumns)) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidRange,
        `Range values must have shape ${expectedRows}x${expectedColumns}`
      );
    }
    return new TableDataMutation(
      this.#dataSource,
      { type: "set-range", range, values },
      [{ type: "range-value", range }]
    );
  }

  createInsertRowsMutation(start: number, rows: readonly (readonly unknown[])[]): Mutation {
    this.#assertTransactionalStructuralCapability("insertRows");
    this.#assertStructuralShape("row", "insert", start, rows.length);
    return new TableDataMutation(
      this.#dataSource,
      { type: "insert-rows", start, rows },
      [{ type: "row-structure", operation: "insert", start, count: rows.length }]
    );
  }

  createDeleteRowsMutation(start: number, count: number): Mutation {
    this.#assertTransactionalStructuralCapability("deleteRows");
    this.#assertStructuralShape("row", "delete", start, count);
    return new TableDataMutation(
      this.#dataSource,
      { type: "delete-rows", start, count },
      [{ type: "row-structure", operation: "delete", start, count }]
    );
  }

  createMoveRowsMutation(start: number, count: number, target: number): Mutation {
    this.#assertTransactionalStructuralCapability("moveRows");
    this.#assertStructuralShape("row", "move", start, count, target);
    return new TableDataMutation(
      this.#dataSource,
      { type: "move-rows", start, count, target },
      [{ type: "row-structure", operation: "move", start, count, target }]
    );
  }

  createInsertColumnsMutation(start: number, columns: readonly ColumnFieldMapping[]): Mutation {
    this.#assertTransactionalStructuralCapability("insertColumns");
    this.#assertStructuralShape("column", "insert", start, columns.length);
    return new TableDataMutation(
      this.#dataSource,
      { type: "insert-columns", start, columns },
      [{ type: "column-structure", operation: "insert", start, count: columns.length }]
    );
  }

  createDeleteColumnsMutation(start: number, count: number): Mutation {
    this.#assertTransactionalStructuralCapability("deleteColumns");
    this.#assertStructuralShape("column", "delete", start, count);
    return new TableDataMutation(
      this.#dataSource,
      { type: "delete-columns", start, count },
      [{ type: "column-structure", operation: "delete", start, count }]
    );
  }

  createMoveColumnsMutation(start: number, count: number, target: number): Mutation {
    this.#assertTransactionalStructuralCapability("moveColumns");
    this.#assertStructuralShape("column", "move", start, count, target);
    return new TableDataMutation(
      this.#dataSource,
      { type: "move-columns", start, count, target },
      [{ type: "column-structure", operation: "move", start, count, target }]
    );
  }

  setCell(row: number, column: number, value: unknown): TableStoreCommitResult {
    return this.#execute(this.createSetCellMutation(row, column, value));
  }

  setRange(range: CellRange, values: readonly (readonly unknown[])[]): TableStoreCommitResult {
    return this.#execute(this.createSetRangeMutation(range, values));
  }

  updateCell(row: number, column: number, updater: (current: unknown) => unknown): TableStoreCommitResult {
    return this.setCell(row, column, updater(this.getCell(row, column)));
  }

  insertRows(start: number, rows: readonly (readonly unknown[])[]): TableStoreCommitResult {
    return this.#execute(this.createInsertRowsMutation(start, rows));
  }

  deleteRows(start: number, count: number): TableStoreCommitResult {
    return this.#execute(this.createDeleteRowsMutation(start, count));
  }

  moveRows(start: number, count: number, target: number): TableStoreCommitResult {
    return this.#execute(this.createMoveRowsMutation(start, count, target));
  }

  insertColumns(start: number, columns: readonly ColumnFieldMapping[]): TableStoreCommitResult {
    return this.#execute(this.createInsertColumnsMutation(start, columns));
  }

  deleteColumns(start: number, count: number): TableStoreCommitResult {
    return this.#execute(this.createDeleteColumnsMutation(start, count));
  }

  moveColumns(start: number, count: number, target: number): TableStoreCommitResult {
    return this.#execute(this.createMoveColumnsMutation(start, count, target));
  }

  #execute(mutation: Mutation): TableStoreCommitResult {
    const session = this.beginMutationSession(`store-${this.#nextDirectTransactionId++}`);
    let receipt: CommitReceipt;
    try {
      session.stage(mutation);
      const prepared = session.seal();
      try {
        receipt = this.#coordinator.commit(prepared);
      } catch (error) {
        this.#coordinator.discard(prepared);
        throw error;
      }
    } catch (error) {
      session.discard();
      throw error;
    }
    return this.#toCommitResult(receipt);
  }

  #toCommitResult(receipt: CommitReceipt): TableStoreCommitResult {
    return { patches: receipt.patches, changeSet: receipt.changeSet, version: receipt.afterRevision };
  }

  #assertCellExists(row: number, column: number): void {
    assertValidCoordinate(row, column);
    if (row >= this.getRowCount()) {
      throw new CyExcelError(CyExcelErrorCode.InvalidCoordinate, `Row ${row} does not exist`);
    }
    if (column >= this.getColumnCount()) {
      throw new CyExcelError(CyExcelErrorCode.ColumnNotFound, `Column ${column} does not exist`);
    }
  }

  #assertTransactionalStructuralCapability(method: keyof DataSource): void {
    if (!this.#dataSource[method] || !this.#dataSource.beginTransaction) {
      throw new CyExcelError(
        CyExcelErrorCode.StructuralModificationNotSupported,
        `This DataSource does not support transactional ${String(method)}()`
      );
    }
  }

  #assertStructuralShape(
    axis: "row" | "column",
    operation: "insert" | "delete" | "move",
    start: number,
    count: number,
    target?: number
  ): void {
    const minimumLength = Math.max(start + (operation === "insert" ? 0 : count), target ?? 0);
    validateAxisChange({ axis, operation, start, count, ...(target === undefined ? {} : { target }) }, minimumLength);
  }
}
