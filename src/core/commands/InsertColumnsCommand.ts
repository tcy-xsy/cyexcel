import { snapshotColumnDefinitions } from "../cells/columnSnapshot";
import type { ColumnDefinition } from "../cells/types";
import type { Command } from "../types";

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
