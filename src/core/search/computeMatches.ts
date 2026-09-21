import type { CellValue } from "../sheet/types";

/**
 * Scans occupied cells for a text match, yielding to the event loop between
 * batches so a scan over millions of occupied cells never shows up as a
 * single long task (Step 23 exit criteria) — unlike Step 22's `filterByColumn`,
 * which is a single synchronous loop with no such guarantee.
 */
export interface SearchMatch {
  readonly row: number;
  readonly column: number;
}

export interface SearchCell {
  readonly row: number;
  readonly column: number;
  readonly value: CellValue | undefined;
}

export interface SearchRequest {
  readonly occupiedCells: Iterable<SearchCell>;
  readonly getDisplayText: (row: number, column: number, value: CellValue | undefined) => string;
  readonly query: string;
  readonly matchCase?: boolean;
  readonly useRegex?: boolean;
  readonly signal?: AbortSignal;
  readonly onProgress?: (scanned: number) => void;
  /** Cells processed per batch before yielding to the event loop. Defaults to 2000. */
  readonly batchSize?: number;
}

/** The caller-facing subset of SearchRequest — everything except the sheet-specific data feed, which CyExcelCore supplies itself. */
export type SearchOptions = Pick<SearchRequest, "matchCase" | "useRegex" | "signal" | "onProgress">;

export interface ReplaceMatchesOptions {
  readonly matchCase?: boolean;
  readonly useRegex?: boolean;
  /** 替换文本里的全部命中；默认只替换第一处。 */
  readonly all?: boolean;
}

function yieldToEventLoop(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * 按查找条件替换文本；没找到时返回 `null`，让调用方区分"没命中"和"替换成空串"。
 * 替换串按字面量写入——这里不展开 `$1`，避免用户输入的普通文本被当成正则反向引用。
 */
export function replaceMatches(
  text: string,
  query: string,
  replacement: string,
  options: ReplaceMatchesOptions = {}
): string | null {
  if (query.length === 0) {
    return null;
  }

  const { matchCase = false, useRegex = false, all = false } = options;
  const pattern = useRegex
    ? new RegExp(query, `${matchCase ? "" : "i"}${all ? "g" : ""}`)
    : new RegExp(escapeRegExp(query), `${matchCase ? "" : "i"}${all ? "g" : ""}`);

  if (!pattern.test(text)) {
    return null;
  }
  pattern.lastIndex = 0;
  const nextText = text.replace(pattern, () => replacement);
  return nextText === text ? null : nextText;
}

/**
 * Returns matches in the order `occupiedCells` yields them, or `null` if
 * cancelled via `signal` — a cancelled scan writes nothing and reports
 * nothing, same convention as `computeSortOrder`.
 */
export async function computeMatches(request: SearchRequest): Promise<readonly SearchMatch[] | null> {
  const { occupiedCells, getDisplayText, query, matchCase = false, useRegex = false, signal, onProgress } = request;
  const batchSize = request.batchSize ?? 2000;

  if (query.length === 0) {
    return [];
  }

  const isMatch = useRegex
    ? ((pattern) => (text: string) => pattern.test(text))(new RegExp(query, matchCase ? "" : "i"))
    : matchCase
      ? (text: string) => text.includes(query)
      : ((needle) => (text: string) => text.toLowerCase().includes(needle))(query.toLowerCase());

  if (signal?.aborted) {
    return null;
  }

  const matches: SearchMatch[] = [];
  let scanned = 0;
  let sinceYield = 0;
  for (const cell of occupiedCells) {
    if (isMatch(getDisplayText(cell.row, cell.column, cell.value))) {
      matches.push({ row: cell.row, column: cell.column });
    }
    scanned += 1;
    onProgress?.(scanned);
    sinceYield += 1;

    if (signal?.aborted) {
      return null;
    }

    if (sinceYield >= batchSize) {
      sinceYield = 0;
      await yieldToEventLoop();
      if (signal?.aborted) {
        return null;
      }
    }
  }

  return matches;
}
