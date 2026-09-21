import { buildMergeClearGrid, rangesIntersect, resolveMergeCells } from "../sheet/merges";
import type { CellRange, CommandContext } from "../types";
import { assertValidRange } from "../validation";
import type { ClearRangeCommand } from "./ClearRangeCommand";
import { CommandHandlerRegistry } from "./CommandHandlerRegistry";
import type { DeleteColumnsCommand } from "./DeleteColumnsCommand";
import type { DeleteRowsCommand } from "./DeleteRowsCommand";
import type { FilterCommand } from "./FilterCommand";
import type { HideColumnCommand } from "./HideColumnCommand";
import type { InsertColumnsCommand } from "./InsertColumnsCommand";
import type { InsertRowsCommand } from "./InsertRowsCommand";
import type { MergeCellsCommand } from "./MergeCellsCommand";
import type { MoveColumnsCommand } from "./MoveColumnsCommand";
import type { MoveRowsCommand } from "./MoveRowsCommand";
import type { ReorderRowsCommand } from "./ReorderRowsCommand";
import type { ResizeColumnCommand } from "./ResizeColumnCommand";
import type { ResizeRowCommand } from "./ResizeRowCommand";
import type { SetCellCommand } from "./SetCellCommand";
import type { SetRangeCommand } from "./SetRangeCommand";
import type { SetStyleCommand } from "./SetStyleCommand";
import type { SetRowHeightsCommand } from "./SetRowHeightsCommand";
import type { ShowColumnCommand } from "./ShowColumnCommand";
import type { UnmergeCellsCommand } from "./UnmergeCellsCommand";
import type {
  DataCommandMutationFactory,
  LayoutCommandMutationFactory,
  StructuralCommandGuards,
  StructureCommandMutationFactory
} from "./capabilities";

export interface BuiltinCommandPlannerOptions {
  readonly data: DataCommandMutationFactory;
  readonly layout?: LayoutCommandMutationFactory;
  readonly structure?: StructureCommandMutationFactory;
  readonly guards?: StructuralCommandGuards;
}

/** `rows` must already be ascending (guaranteed by `DeleteRowsCommand`'s own normalization). */
function toContiguousRuns(rows: readonly number[]): readonly { readonly start: number; readonly count: number }[] {
  const runs: { start: number; count: number }[] = [];
  for (const row of rows) {
    const last = runs.at(-1);
    if (last !== undefined && row === last.start + last.count) {
      last.count += 1;
    } else {
      runs.push({ start: row, count: 1 });
    }
  }
  return runs;
}

function createEmptyGrid(range: CellRange): readonly (readonly undefined[])[] {
  const rowCount = range.endRow - range.startRow + 1;
  const columnCount = range.endColumn - range.startColumn + 1;
  return Object.freeze(
    Array.from({ length: rowCount }, () =>
      Object.freeze(Array.from({ length: columnCount }, () => undefined))
    )
  );
}

/**
 * `mergeCells(range)` 清空非锚点格：`buildMergeClearGrid` 用 `undefined` 表示锚点格
 * "跳过、保留原值"，但 `valueToRecord`/`writeCell`（`sheet/import.ts`、`sheet/apply.ts`）
 * 把网格里的 `undefined` 一律当成"删掉这一格"，不是"跳过"——同一份网格整体交给
 * `createSetRange` 会连锚点格一起清空。于是这里不发一条覆盖整个 `range` 的写入，
 * 而是拆成两条**都不包含锚点格**的矩形子写入：锚点所在行的其余列一条、锚点行以下
 * 的所有行一条；`range` 只有一行一列（锚点是唯一的格子）时两条都不发。
 */
function stageMergeClear(
  range: CellRange,
  data: DataCommandMutationFactory,
  context: CommandContext
): void {
  const grid = buildMergeClearGrid(range);
  const anchorRow = grid[0];
  if (range.endColumn > range.startColumn && anchorRow !== undefined) {
    context.stage(
      data.createSetRange(
        { ...range, startColumn: range.startColumn + 1, endRow: range.startRow },
        [anchorRow.slice(1)]
      )
    );
  }
  if (range.endRow > range.startRow) {
    context.stage(
      data.createSetRange({ ...range, startRow: range.startRow + 1 }, grid.slice(1))
    );
  }
}

export function createBuiltinCommandPlanner(options: BuiltinCommandPlannerOptions): CommandHandlerRegistry {
  const registry = new CommandHandlerRegistry();

  registry.register<SetCellCommand>({
    type: "set-cell",
    plan: ({ row, column, value }, context) => context.stage(options.data.createSetCell(row, column, value))
  });
  registry.register<SetRangeCommand>({
    type: "set-range",
    plan: ({ range, values }, context) => context.stage(options.data.createSetRange(range, values))
  });
  registry.register<SetStyleCommand>({
    type: "set-style",
    plan: ({ range, patch }, context) => {
      assertValidRange(range);
      context.stage(options.data.createSetStyle(range, patch));
    }
  });
  registry.register<MergeCellsCommand>({
    type: "merge-cells",
    plan: ({ range }, context) => {
      assertValidRange(range);
      // 重叠决策唯一入口是 `resolveMergeCells`：部分相交拒绝、完全包含吸收
      // （设计稿 §0.2/§2），不是自己拼 `[...merges, range]`。
      const merges = resolveMergeCells(context.model.merges(), range);
      context.stage(options.data.createSetMerges(merges));
      stageMergeClear(range, options.data, context);
    }
  });
  registry.register<UnmergeCellsCommand>({
    type: "unmerge-cells",
    plan: ({ range }, context) => {
      assertValidRange(range);
      const remaining = context.model.merges().filter((merge) => !rangesIntersect(merge, range));
      context.stage(options.data.createSetMerges(remaining));
    }
  });
  registry.register<ClearRangeCommand>({
    type: "clear-range",
    plan: ({ range }, context) => context.stage(options.data.createSetRange(range, createEmptyGrid(range)))
  });
  registry.register<InsertRowsCommand>({
    type: "insert-rows",
    plan: ({ start, rows }, context) => {
      const change = { operation: "insert" as const, start, count: rows.length };
      options.guards?.checkRows(change);
      context.stage(options.data.createInsertRows(start, rows));
      const structure = options.structure?.createInsertRows(start, rows.length);
      if (structure) {
        context.stage(structure);
      }
    }
  });
  registry.register<DeleteRowsCommand>({
    type: "delete-rows",
    plan: (params, context) => {
      const runs = "rows" in params ? toContiguousRuns(params.rows) : [params];
      // Bottom-up: a run's own start/count is only meaningful against the
      // *not-yet-shifted* row indices above it, so the highest-start run must
      // be staged (and its data physically removed) before any lower run's
      // indices are touched.
      for (let index = runs.length - 1; index >= 0; index -= 1) {
        const { start, count } = runs[index] as { start: number; count: number };
        const change = { operation: "delete" as const, start, count };
        options.guards?.checkRows(change);
        context.stage(options.data.createDeleteRows(start, count));
        const structure = options.structure?.createDeleteRows(start, count);
        if (structure) {
          context.stage(structure);
        }
      }
    }
  });
  registry.register<MoveRowsCommand>({
    type: "move-rows",
    plan: ({ start, count, target }, context) => {
      const change = { operation: "move" as const, start, count, target };
      options.guards?.checkRows(change);
      context.stage(options.data.createMoveRows(start, count, target));
      const structure = options.structure?.createMoveRows(start, count, target);
      if (structure) {
        context.stage(structure);
      }
    }
  });
  registry.register<ReorderRowsCommand>({
    type: "reorder-rows",
    plan: ({ order }, context) => {
      // Staged into the same context as the data mutation, exactly like
      // move-rows stages its structure participant: one transaction, so no
      // published frame can show reordered data against unreordered row
      // geometry. The layout side is a LayoutRootAction rather than a
      // structure participant because a permutation has no
      // start/count/target to put in an AxisChangeDescriptor.
      //
      // No `options.guards?.checkRows(...)` here: beforeRowsChange speaks
      // insert/delete/move with a start and a count, and inventing one for a
      // whole-axis permutation would hand every hook a lie about which rows
      // moved.
      context.stage(options.data.createReorderRows(order));
      const layout = options.layout?.createLayoutMutation({ type: "reorder-rows", order });
      if (layout) {
        context.stage(layout);
      }
    }
  });
  registry.register<InsertColumnsCommand>({
    type: "insert-columns",
    plan: ({ start, columns }, context) => {
      const change = { operation: "insert" as const, start, count: columns.length };
      options.guards?.checkColumns(change);
      // 只把个数交给数据层：列在存储里没有身份（ADR 0004 §3.8.2），`columns` 里剩下
      // 的是宽度、类型这些展示属性，它们的家在 `structure` 那一半。
      context.stage(options.data.createInsertColumns(start, columns.length));
      const structure = options.structure?.createInsertColumns(start, columns);
      if (structure) {
        context.stage(structure);
      }
      // 显式宽度还要写进权威（`columnData[c].w`）：`structure` 那一条只补了几何索引，
      // 而列宽的权威在模型里（ADR 0004 阶段 4B T8b，设计稿 §4.4）。缺省宽的列一条都不
      // 发——「没有 `w`」正是它要表达的状态。
      //
      // 排在 `structure` 之后：这些下标是**插入之后**的坐标空间里的。用
      // `set-column-meta` 而不是 `set-column-width`，是因为它同时也把几何宽度重算一遍
      // 写回索引，而那一步与 `structure` 刚写的值相同，幂等；撤销按相反顺序退，先撤
      // 元数据再撤结构，每一步两边都对得上。
      const layout = options.layout;
      if (layout) {
        columns.forEach((column, offset) => {
          if (column.width === undefined) {
            return;
          }
          context.stage(layout.createLayoutMutation({
            type: "set-column-meta",
            column: start + offset,
            meta: { w: column.width }
          }));
        });
      }
    }
  });
  registry.register<DeleteColumnsCommand>({
    type: "delete-columns",
    plan: ({ start, count }, context) => {
      const change = { operation: "delete" as const, start, count };
      options.guards?.checkColumns(change);
      context.stage(options.data.createDeleteColumns(start, count));
      const structure = options.structure?.createDeleteColumns(start, count);
      if (structure) {
        context.stage(structure);
      }
    }
  });
  registry.register<MoveColumnsCommand>({
    type: "move-columns",
    plan: ({ start, count, target }, context) => {
      const change = { operation: "move" as const, start, count, target };
      options.guards?.checkColumns(change);
      context.stage(options.data.createMoveColumns(start, count, target));
      const structure = options.structure?.createMoveColumns(start, count, target);
      if (structure) {
        context.stage(structure);
      }
    }
  });

  if (options.layout) {
    const layout = options.layout;
    registry.register<ResizeRowCommand>({
      type: "resize-row",
      plan: ({ row, height }, context) =>
        context.stage(layout.createRowHeightMutation([{ row, height }], "manual"))
    });
    registry.register<SetRowHeightsCommand>({
      type: "set-row-heights",
      plan: ({ updates, provenance }, context) =>
        context.stage(layout.createRowHeightMutation(updates, provenance))
    });
    registry.register<ResizeColumnCommand>({
      type: "resize-column",
      plan: ({ column, width }, context) =>
        context.stage(layout.createLayoutMutation({ type: "set-column-width", column, width }))
    });
    registry.register<HideColumnCommand>({
      type: "hide-column",
      plan: ({ column }, context) =>
        context.stage(layout.createLayoutMutation({ type: "hide-column", column }))
    });
    registry.register<ShowColumnCommand>({
      type: "show-column",
      plan: ({ column }, context) =>
        context.stage(layout.createLayoutMutation({ type: "show-column", column }))
    });
    registry.register<FilterCommand>({
      type: "filter-rows",
      // Stages only a layout mutation — no `options.data.*` call at all.
      // That absence *is* ADR 0003 §3.2's "过滤没有写数据路径" made
      // structural, not just documented; no `options.guards?.checkRows(...)`
      // either, since nothing here is a row insert/delete/move.
      plan: ({ filteredRows }, context) =>
        context.stage(layout.createLayoutMutation({ type: "set-filtered-rows", filteredRows }))
    });
  }

  return registry;
}
