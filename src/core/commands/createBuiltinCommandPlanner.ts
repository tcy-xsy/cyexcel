import type { CellRange } from "../types";
import type { ClearRangeCommand } from "./ClearRangeCommand";
import { CommandHandlerRegistry } from "./CommandHandlerRegistry";
import type { DeleteColumnsCommand } from "./DeleteColumnsCommand";
import type { DeleteRowsCommand } from "./DeleteRowsCommand";
import type { HideColumnCommand } from "./HideColumnCommand";
import type { InsertColumnsCommand } from "./InsertColumnsCommand";
import type { InsertRowsCommand } from "./InsertRowsCommand";
import type { MoveColumnsCommand } from "./MoveColumnsCommand";
import type { MoveRowsCommand } from "./MoveRowsCommand";
import type { ReorderRowsCommand } from "./ReorderRowsCommand";
import type { ResizeColumnCommand } from "./ResizeColumnCommand";
import type { ResizeRowCommand } from "./ResizeRowCommand";
import type { SetCellCommand } from "./SetCellCommand";
import type { SetRangeCommand } from "./SetRangeCommand";
import type { SetRowHeightsCommand } from "./SetRowHeightsCommand";
import type { ShowColumnCommand } from "./ShowColumnCommand";
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

function createEmptyGrid(range: CellRange): readonly (readonly undefined[])[] {
  const rowCount = range.endRow - range.startRow + 1;
  const columnCount = range.endColumn - range.startColumn + 1;
  return Object.freeze(
    Array.from({ length: rowCount }, () =>
      Object.freeze(Array.from({ length: columnCount }, () => undefined))
    )
  );
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
    plan: ({ start, count }, context) => {
      const change = { operation: "delete" as const, start, count };
      options.guards?.checkRows(change);
      context.stage(options.data.createDeleteRows(start, count));
      const structure = options.structure?.createDeleteRows(start, count);
      if (structure) {
        context.stage(structure);
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
      const mappings = Object.freeze(columns.map((column) => Object.freeze({
        id: column.id,
        field: column.field ?? column.id
      })));
      context.stage(options.data.createInsertColumns(start, mappings));
      const structure = options.structure?.createInsertColumns(start, columns);
      if (structure) {
        context.stage(structure);
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
  }

  return registry;
}
