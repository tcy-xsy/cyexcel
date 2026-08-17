import type { Command } from "../types";

export interface InsertRowsCommandParams {
  readonly start: number;
  readonly rows: readonly (readonly unknown[])[];
}

/** `rows` is column-ordered values, one array per new row (same convention as SetRangeCommand) — DataSource.insertRows() owns the field-name mapping. */
export class InsertRowsCommand implements Command<"insert-rows", InsertRowsCommandParams> {
  readonly type = "insert-rows";
  readonly params: InsertRowsCommandParams;

  constructor(start: number, rows: readonly (readonly unknown[])[]) {
    this.params = Object.freeze({
      start,
      rows: Object.freeze(rows.map((row) => Object.freeze([...row])))
    });
    Object.freeze(this);
  }
}
