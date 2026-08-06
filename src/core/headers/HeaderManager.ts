import type { ColumnDefinition } from "../cells";
import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { IndexRange } from "../layout";
import { splitByFrozenBoundary } from "../layout";
import type { HeaderGroupDefinition, HeaderGroupFragment, RuntimeHeader, RuntimeHeaderGroup } from "./types";

export interface HeaderManagerOptions {
  readonly columns: readonly ColumnDefinition[];
  /** Two-level grouping only — a group cannot contain another group. */
  readonly groups?: readonly HeaderGroupDefinition[];
}

function invalidGroup(group: HeaderGroupDefinition, reason: string): CyExcelError {
  return new CyExcelError(
    CyExcelErrorCode.InvalidOptions,
    `Invalid header group "${group.id}": ${reason}`
  );
}

function resolveGroup(group: HeaderGroupDefinition, columnCount: number): RuntimeHeaderGroup {
  if (group.columns.length === 0) {
    throw invalidGroup(group, "columns must not be empty");
  }
  const sorted = [...group.columns].sort((a, b) => a - b);
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

  return { id: group.id, title: group.title, startColumn, endColumn };
}

/**
 * Resolves declared column `header`s and (optional) two-level groups once,
 * at construction — mirrors ColumnManager's cached-merge design (Step 8).
 * `getGroupFragments()` is the shared geometry both HeaderCanvasRenderer's
 * painting and this class's own tests use to verify Step 9's 4th
 * pre-written test ("分组跨冻结边界时拆成绘制片段但保留同一逻辑 ID").
 */
export class HeaderManager {
  readonly #headers: readonly RuntimeHeader[];
  readonly #groups: readonly RuntimeHeaderGroup[];

  constructor(options: HeaderManagerOptions) {
    this.#headers = options.columns.map((column, index) => {
      const declared = column.header;
      if (typeof declared === "string") {
        return { column: index, title: declared };
      }
      const title = declared?.title ?? column.id;
      return declared?.renderer !== undefined
        ? { column: index, title, renderer: declared.renderer }
        : { column: index, title };
    });

    this.#groups = (options.groups ?? []).map((group) =>
      resolveGroup(group, options.columns.length)
    );
  }

  getHeader(column: number): RuntimeHeader {
    const header = this.#headers[column];
    if (!header) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidCoordinate,
        `Header column index ${column} is out of range`
      );
    }
    return header;
  }

  getHeaders(): readonly RuntimeHeader[] {
    return this.#headers;
  }

  getGroups(): readonly RuntimeHeaderGroup[] {
    return this.#groups;
  }

  /** Splits every group's column span at the frozen boundary; a group entirely on one side yields exactly one fragment. */
  getGroupFragments(frozenColumnCount: number): readonly HeaderGroupFragment[] {
    const fragments: HeaderGroupFragment[] = [];
    for (const group of this.#groups) {
      const columnRange: IndexRange = {
        start: group.startColumn,
        count: group.endColumn - group.startColumn + 1
      };
      const split = splitByFrozenBoundary(columnRange, frozenColumnCount);

      if (split.frozen.count > 0) {
        fragments.push({
          group,
          startColumn: split.frozen.start,
          endColumn: split.frozen.start + split.frozen.count - 1,
          frozenColumn: true
        });
      }
      if (split.scrollable.count > 0) {
        fragments.push({
          group,
          startColumn: split.scrollable.start,
          endColumn: split.scrollable.start + split.scrollable.count - 1,
          frozenColumn: false
        });
      }
    }
    return fragments;
  }
}
