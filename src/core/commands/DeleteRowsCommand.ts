import type { Command } from "../types";

export type DeleteRowsCommandParams =
  | { readonly start: number; readonly count: number }
  | { readonly rows: readonly number[] };

function normalizeRows(rows: readonly number[]): readonly number[] {
  const unique = [...new Set(rows)];
  for (const row of unique) {
    if (!Number.isSafeInteger(row) || row < 0) {
      throw new RangeError(`row must be a non-negative safe integer: ${row}`);
    }
  }
  unique.sort((a, b) => a - b);
  return Object.freeze(unique);
}

/**
 * 连续区间形态（`start`/`count`）是常规删除；非连续行形态（`rows`）是
 * ADR 0003 §3.5 说的"唯一的非连续行集合"——过滤态下选中一段再删除，只应删掉
 * 其中可见的那些行。`rows` 在构造时就升序去重；是否在 `rowCount` 范围内由
 * planner 的 `guards.checkRows` 校验，这里不知道表有多少行。
 */
export class DeleteRowsCommand implements Command<"delete-rows", DeleteRowsCommandParams> {
  readonly type = "delete-rows";
  readonly params: DeleteRowsCommandParams;

  constructor(startOrRows: number | readonly number[], count?: number) {
    if (typeof startOrRows === "number") {
      if (count === undefined) {
        throw new RangeError("DeleteRowsCommand(start, count) requires count");
      }
      this.params = Object.freeze({ start: startOrRows, count });
    } else {
      this.params = Object.freeze({ rows: normalizeRows(startOrRows) });
    }
    Object.freeze(this);
  }
}
