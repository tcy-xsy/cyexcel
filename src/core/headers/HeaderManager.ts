import type { ColumnDefinition } from "../cells";
import { CyExcelError, CyExcelErrorCode } from "../errors";
import { columnIndexToLetter } from "./columnLetter";
import type { HeaderGroupDefinition, RuntimeHeader, RuntimeHeaderGroup } from "./types";

export interface HeaderManagerOptions {
  readonly columns: readonly ColumnDefinition[];
  /** Arbitrary-depth nesting via `children` — see HeaderGroupDefinition's own doc comment. */
  readonly groups?: readonly HeaderGroupDefinition[];
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

function invalidGroup(group: HeaderGroupDefinition, reason: string): CyExcelError {
  return new CyExcelError(
    CyExcelErrorCode.InvalidOptions,
    `Invalid header group "${group.id}": ${reason}`
  );
}

/** Validates and resolves a leaf node's own `columns` — the exact Step 9 logic, just decoupled from the recursive tree walk below. */
function resolveLeafColumns(
  group: HeaderGroupDefinition,
  columns: readonly number[],
  columnCount: number
): { readonly startColumn: number; readonly endColumn: number } {
  if (columns.length === 0) {
    throw invalidGroup(group, "columns must not be empty");
  }
  const sorted = [...columns].sort((a, b) => a - b);
  const startColumn = sorted[0] as number;
  const endColumn = sorted[sorted.length - 1] as number;

  if (endColumn - startColumn + 1 !== sorted.length) {
    throw invalidGroup(group, "columns must be contiguous");
  }
  for (let index = 0; index < sorted.length; index += 1) {
    if (sorted[index] !== startColumn + index) {
      throw invalidGroup(group, "columns must not repeat");
    }
  }
  if (startColumn < 0 || endColumn >= columnCount) {
    throw invalidGroup(group, `columns must be within [0, ${columnCount - 1}]`);
  }

  return { startColumn, endColumn };
}

interface ResolvedNode {
  readonly group: RuntimeHeaderGroup;
  /** 0 = leaf-adjacent (declares `columns` directly); a branch is `1 + max(child levels)`. */
  readonly level: number;
}

/** Immutable, fully validated header state that can be prepared off to the side. */
export interface HeaderManagerRoot {
  readonly headers: readonly RuntimeHeader[];
  readonly groupsByLevel: readonly (readonly RuntimeHeaderGroup[])[];
}

/**
 * Recursively resolves one tree node, pushing the resulting RuntimeHeaderGroup
 * into `groupsByLevel[level]` as a side effect (simplest way to accumulate
 * across an arbitrary number of independent trees in the top-level array).
 */
function resolveNode(
  node: HeaderGroupDefinition,
  columnCount: number,
  groupsByLevel: RuntimeHeaderGroup[][]
): ResolvedNode {
  const hasColumns = node.columns !== undefined;
  const hasChildren = node.children !== undefined;
  if (hasColumns && hasChildren) {
    throw invalidGroup(node, "must not specify both columns and children");
  }
  if (!hasColumns && !hasChildren) {
    throw invalidGroup(node, "must specify either columns or children");
  }

  if (hasChildren) {
    const children = node.children as readonly HeaderGroupDefinition[];
    if (children.length === 0) {
      throw invalidGroup(node, "children must not be empty");
    }
    const resolvedChildren = children.map((child) => resolveNode(child, columnCount, groupsByLevel));
    const sorted = [...resolvedChildren].sort((a, b) => a.group.startColumn - b.group.startColumn);
    for (let index = 0; index < sorted.length; index += 1) {
      const current = sorted[index] as ResolvedNode;
      const previous = sorted[index - 1];
      if (previous && current.group.startColumn !== previous.group.endColumn + 1) {
        throw invalidGroup(node, "children must be contiguous, without gaps or overlaps");
      }
    }
    const startColumn = (sorted[0] as ResolvedNode).group.startColumn;
    const endColumn = (sorted[sorted.length - 1] as ResolvedNode).group.endColumn;
    const level = 1 + Math.max(...resolvedChildren.map((child) => child.level));
    const group: RuntimeHeaderGroup = { id: node.id, title: node.title, startColumn, endColumn };
    (groupsByLevel[level] ??= []).push(group);
    return { group, level };
  }

  const { startColumn, endColumn } = resolveLeafColumns(node, node.columns as readonly number[], columnCount);
  const group: RuntimeHeaderGroup = { id: node.id, title: node.title, startColumn, endColumn };
  (groupsByLevel[0] ??= []).push(group);
  return { group, level: 0 };
}

/**
 * Resolves declared column `header`s and (optional) arbitrary-depth groups
 * once, at construction — mirrors ColumnManager's cached-merge design
 * (Step 8). Pane splitting belongs to Scene/HeaderCanvasRenderer; this
 * class only owns the logical group spans.
 */
export class HeaderManager {
  #root: HeaderManagerRoot;
  /** Fixed at construction — there's no runtime API to change it, same as any other CyExcelCoreOptions field. */
  readonly #columnLetterFallback: boolean;

  constructor(options: HeaderManagerOptions) {
    // 默认 A/B/C（ADR 0004 §3.8.1）——列没有 `id` 之后没有第二个可当默认的答案。
    this.#columnLetterFallback = options.columnLetterFallback ?? true;
    this.#root = this.prepareRoot(options.columns, options.groups);
  }

  get root(): HeaderManagerRoot {
    return this.#root;
  }

  /**
   * Resolves and validates a complete root without touching the live manager.
   * Callers may discard the result if any later participant fails.
   */
  prepareRoot(columns: readonly ColumnDefinition[], groups?: readonly HeaderGroupDefinition[]): HeaderManagerRoot {
    const headers = HeaderManager.#resolveHeaders(columns, this.#columnLetterFallback);
    const groupsByLevel = HeaderManager.#resolveGroups(columns, groups);
    return Object.freeze({ headers, groupsByLevel });
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

  static #resolveGroups(
    columns: readonly ColumnDefinition[],
    groups: readonly HeaderGroupDefinition[] | undefined
  ): readonly (readonly RuntimeHeaderGroup[])[] {
    const groupsByLevel: RuntimeHeaderGroup[][] = [];
    for (const node of groups ?? []) {
      resolveNode(node, columns.length, groupsByLevel);
    }
    return Object.freeze(groupsByLevel.map((groups) => Object.freeze(groups.map((group) => Object.freeze(group)))));
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

  /** Number of group rows — 0 if `groups` was never configured, 1 for Step 9's single-level case, N for an N-deep tree (the deepest branch wins). */
  get groupDepth(): number {
    return this.#root.groupsByLevel.length;
  }

  /** `level` counts outward from the leaves: 0 is the leaf-adjacent row (Step 9's only level and the default), `groupDepth - 1` is the outermost. Out-of-range levels return an empty array, not an error. */
  getGroups(level = 0): readonly RuntimeHeaderGroup[] {
    return this.#root.groupsByLevel[level] ?? [];
  }

}
