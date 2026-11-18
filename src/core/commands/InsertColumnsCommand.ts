import { snapshotColumnDefinitions } from "../cells/columnSnapshot.js";
import type { ColumnDefinition } from "../cells/types.js";
import type { Command } from "../types.js";

export interface InsertColumnsCommandParams {
  readonly start: number;
  readonly columns: readonly ColumnDefinition[];
}

/** A column intent carries its full definition so data and structure participants derive from one snapshot. */
export class InsertColumnsCommand implements Command<"insert-columns", InsertColumnsCommandParams> {
  readonly type = "insert-columns";
  readonly params: InsertColumnsCommandParams;

  constructor(start: number, columns: readonly ColumnDefinition[]) {
    this.params = Object.freeze({ start, columns: snapshotColumnDefinitions(columns) });
    Object.freeze(this);
  }
}
