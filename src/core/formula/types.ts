import type { CellRecord, CellValue } from "../sheet/types";

export interface FormulaCell {
  readonly row: number;
  readonly column: number;
}

export interface FormulaModelAccess {
  getRecord(row: number, column: number): CellRecord | undefined;
}

/** Calculation is derived state: invalidation never creates history entries. */
export interface FormulaEngine {
  getValue(row: number, column: number): CellValue | undefined;
  getError(row: number, column: number): string | undefined;
  invalidate(row: number, column: number): readonly FormulaCell[];
  dispose(): void;
}

export type FormulaEngineFactory = (model: FormulaModelAccess) => FormulaEngine;
