import type { ColumnDefinition } from "../cells";
import { CyExcelError, CyExcelErrorCode } from "../errors";
import { columnIndexToLetter } from "./columnLetter";
import type { RuntimeHeader } from "./types";

export interface HeaderManagerOptions {
  readonly columns: readonly ColumnDefinition[];
  /**
   * When a column has no explicit `header`, show its spreadsheet-style
   * letter (A, B, ..., Z, AA, ...) instead of falling back to its `id`.
   * Default `false` — unset behavior (fall back to `id`) is unchanged. An
   * explicit `header` always overrides the letter for that one column —
   * letter and custom title are mutually exclusive per column, matching
   * Univer's `customizeColumnHeader()` model (one row, not a stacked
   * second row), not two simultaneous header rows.
   */
  readonly columnLetterFallback?: boolean;
}

/** Immutable, fully validated header state that can be prepared off to the side. */
export interface HeaderManagerRoot {
  readonly headers: readonly RuntimeHeader[];
}

/**
 * Resolves declared column `header`s once, at construction — mirrors
 * ColumnManager's cached-merge design (Step 8). Pane splitting belongs to
 * Scene/HeaderCanvasRenderer; this class only owns the logical header row.
 */
export class HeaderManager {
  #root: HeaderManagerRoot;
  /** Fixed at construction — there's no runtime API to change it, same as any other CyExcelCoreOptions field. */
  readonly #columnLetterFallback: boolean;

  constructor(options: HeaderManagerOptions) {
    // 默认 A/B/C（ADR 0004 §3.8.1）——列没有 `id` 之后没有第二个可当默认的答案。
    this.#columnLetterFallback = options.columnLetterFallback ?? true;
    this.#root = this.prepareRoot(options.columns);
  }

  get root(): HeaderManagerRoot {
    return this.#root;
  }

  /**
   * Resolves and validates a complete root without touching the live manager.
   * Callers may discard the result if any later participant fails.
   */
  prepareRoot(columns: readonly ColumnDefinition[]): HeaderManagerRoot {
    const headers = HeaderManager.#resolveHeaders(columns, this.#columnLetterFallback);
    return Object.freeze({ headers });
  }

  /** Swaps a previously prepared root; no resolution, allocation, or hooks run here. */
  replaceRoot(root: HeaderManagerRoot): void {
    this.#root = root;
  }

  static #resolveHeaders(
    columns: readonly ColumnDefinition[],
    columnLetterFallback: boolean
  ): readonly RuntimeHeader[] {
    return Object.freeze(columns.map((column, index) => {
      const declared = column.header;
      /*
       * 兜底原来是 `column.id`，而列已经没有 `id` 了（ADR 0004 §3.8.2）。剩下
       * 的两种答案分别对应这个开关的两个值：打开是电子表格字母（A/B/C，默认，
       * §3.8.1 把默认从 `false` 翻成了 `true`），关掉是**没有标题**——不是回退
       * 到别的什么名字，而是这一列就不显示标题。关掉它的语义因此从"用 id 当
       * 标题"变成了"只显示我自己声明的标题"。
       */
      const fallback = columnLetterFallback ? columnIndexToLetter(index) : "";
      if (typeof declared === "string") {
        return Object.freeze({ column: index, title: declared });
      }
      const title = declared?.title ?? fallback;
      return declared?.renderer !== undefined
        ? Object.freeze({ column: index, title, renderer: declared.renderer })
        : Object.freeze({ column: index, title });
    }));
  }

  getHeader(column: number): RuntimeHeader {
    const header = this.#root.headers[column];
    if (!header) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidCoordinate,
        `Header column index ${column} is out of range`
      );
    }
    return header;
  }

  getHeaders(): readonly RuntimeHeader[] {
    return this.#root.headers;
  }

}
