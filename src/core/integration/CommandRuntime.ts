import {
  CommandManager,
  type CommandHandlerRegistry,
  createBuiltinCommandPlanner,
  type CommandManagerOptions,
  type CommandTransactionHost,
  type DataCommandMutationFactory,
  type LayoutCommandMutationFactory,
  type StructuralCommandGuards,
  type StructureCommandMutationFactory
} from "../commands";
import type { TableStore } from "../data/TableStore";
import {
  BEFORE_COLUMNS_CHANGE_HOOK,
  BEFORE_COMMAND_HOOK,
  BEFORE_ROWS_CHANGE_HOOK,
  BEFORE_TRANSACTION_COMMIT_HOOK,
  HookManager,
  type HookRegistry
} from "../hooks";
import type { LayoutEngine } from "../layout/LayoutEngine";
import { LayoutRootMutation, type LayoutRootAction } from "../mutations/LayoutRootMutation";
import { RowHeightMutation } from "../mutations/RowHeightMutation";
import { StructureIntegrationMutation } from "../mutations/StructureIntegrationMutation";
import type { CommitReceipt, Mutation, MutationPrepareSession, PreparedTransaction } from "../mutations/types";
import type { CellRange } from "../types";
import { checkBeforeColumnsChange } from "../commands/checkBeforeColumnsChange";
import { checkBeforeRowsChange } from "../commands/checkBeforeRowsChange";
import type { ColumnDefinition } from "../cells/types";
import { DEFAULT_COLUMN_WIDTH } from "../cells/ColumnManager";
import { createGridStructurePayload } from "./GridStructureOwner";
import type { GridStructureOwner } from "./GridStructureOwner";

const STANDALONE_COMMAND_HOOKS: HookRegistry = {
  beforeCommand: BEFORE_COMMAND_HOOK,
  beforeTransactionCommit: BEFORE_TRANSACTION_COMMIT_HOOK,
  beforeRowsChange: BEFORE_ROWS_CHANGE_HOOK,
  beforeColumnsChange: BEFORE_COLUMNS_CHANGE_HOOK
};

export class TableCommandTransactionHost implements CommandTransactionHost {
  readonly #store: TableStore;

  constructor(store: TableStore) {
    this.#store = store;
  }

  begin(transactionId: string): MutationPrepareSession {
    return this.#store.beginMutationSession(transactionId);
  }

  commit(prepared: PreparedTransaction): CommitReceipt {
    return this.#store.coordinator.commit(prepared);
  }

  discard(prepared: PreparedTransaction): void {
    this.#store.coordinator.discard(prepared);
  }
}

export class TableDataCommandMutationFactory implements DataCommandMutationFactory {
  readonly #store: TableStore;

  constructor(store: TableStore) {
    this.#store = store;
  }

  createSetCell(row: number, column: number, value: unknown): Mutation {
    return this.#store.createSetCellMutation(row, column, value);
  }

  createSetRange(range: CellRange, values: readonly (readonly unknown[])[]): Mutation {
    return this.#store.createSetRangeMutation(range, values);
  }

  createInsertRows(start: number, rows: readonly (readonly unknown[])[]): Mutation {
    return this.#store.createInsertRowsMutation(start, rows);
  }

  createDeleteRows(start: number, count: number): Mutation {
    return this.#store.createDeleteRowsMutation(start, count);
  }

  createMoveRows(start: number, count: number, target: number): Mutation {
    return this.#store.createMoveRowsMutation(start, count, target);
  }

  createReorderRows(order: Int32Array): Mutation {
    return this.#store.createReorderRowsMutation(order);
  }

  createInsertColumns(start: number, count: number): Mutation {
    return this.#store.createInsertColumnsMutation(start, count);
  }

  createDeleteColumns(start: number, count: number): Mutation {
    return this.#store.createDeleteColumnsMutation(start, count);
  }

  createMoveColumns(start: number, count: number, target: number): Mutation {
    return this.#store.createMoveColumnsMutation(start, count, target);
  }
}

export class GridLayoutCommandMutationFactory implements LayoutCommandMutationFactory {
  readonly #layout: LayoutEngine;

  constructor(layout: LayoutEngine) {
    this.#layout = layout;
  }

  createLayoutMutation(action: LayoutRootAction): Mutation {
    return new LayoutRootMutation(this.#layout, action);
  }

  createRowHeightMutation(
    updates: Parameters<LayoutCommandMutationFactory["createRowHeightMutation"]>[0],
    provenance: Parameters<LayoutCommandMutationFactory["createRowHeightMutation"]>[1]
  ): Mutation {
    return new RowHeightMutation(
      this.#layout,
      updates.map((update) => ({ ...update, provenance })),
      provenance === "auto" ? "preserve-manual" : "overwrite"
    );
  }
}

export class GridStructureCommandMutationFactory implements StructureCommandMutationFactory {
  readonly #owner: GridStructureOwner;

  constructor(owner: GridStructureOwner) {
    this.#owner = owner;
  }

  createInsertRows(start: number, count: number): Mutation {
    return this.#create({ axis: "row", operation: "insert", start, count });
  }

  createDeleteRows(start: number, count: number): Mutation {
    return this.#create({ axis: "row", operation: "delete", start, count });
  }

  createMoveRows(start: number, count: number, target: number): Mutation {
    return this.#create({ axis: "row", operation: "move", start, count, target });
  }

  createInsertColumns(start: number, columns: readonly ColumnDefinition[]): Mutation {
    return this.#create(
      { axis: "column", operation: "insert", start, count: columns.length },
      createGridStructurePayload({
        columns,
        columnWidths: columns.map((column) => column.width ?? DEFAULT_COLUMN_WIDTH)
      })
    );
  }

  createDeleteColumns(start: number, count: number): Mutation {
    return this.#create({ axis: "column", operation: "delete", start, count });
  }

  createMoveColumns(start: number, count: number, target: number): Mutation {
    return this.#create({ axis: "column", operation: "move", start, count, target });
  }

  #create(
    change: Parameters<GridStructureOwner["prepare"]>[1],
    payload = createGridStructurePayload()
  ): Mutation {
    return new StructureIntegrationMutation(this.#owner, change, payload);
  }
}

export class HookStructuralCommandGuards implements StructuralCommandGuards {
  readonly #hooks: HookManager;

  constructor(hooks: HookManager) {
    this.#hooks = hooks;
  }

  checkRows(change: Parameters<StructuralCommandGuards["checkRows"]>[0]): void {
    checkBeforeRowsChange(this.#hooks, change);
  }

  checkColumns(change: Parameters<StructuralCommandGuards["checkColumns"]>[0]): void {
    checkBeforeColumnsChange(this.#hooks, change);
  }
}

function commandHooks(options: CommandManagerOptions): HookManager {
  return options.hooks ?? new HookManager({ registry: STANDALONE_COMMAND_HOOKS });
}

export interface TableCommandManagerOptions extends CommandManagerOptions {
  readonly layout?: LayoutEngine;
  readonly configureHandlers?: (handlers: CommandHandlerRegistry) => void;
}

export class TableCommandManager extends CommandManager {
  constructor(store: TableStore, options: TableCommandManagerOptions = {}) {
    const hooks = commandHooks(options);
    const planner = createBuiltinCommandPlanner({
      data: new TableDataCommandMutationFactory(store),
      ...(options.layout !== undefined
        ? { layout: new GridLayoutCommandMutationFactory(options.layout) }
        : {}),
      guards: new HookStructuralCommandGuards(hooks)
    });
    options.configureHandlers?.(planner);
    super(new TableCommandTransactionHost(store), planner, { ...options, hooks });
  }
}

export function createTableCommandManager(
  store: TableStore,
  options: TableCommandManagerOptions = {}
): CommandManager {
  return new TableCommandManager(store, options);
}

export interface GridCommandManagerOptions extends CommandManagerOptions {
  readonly store: TableStore;
  readonly layout: LayoutEngine;
  readonly structureOwner: GridStructureOwner;
}

export function createGridCommandManager(options: GridCommandManagerOptions): CommandManager {
  const hooks = commandHooks(options);
  const planner = createBuiltinCommandPlanner({
    data: new TableDataCommandMutationFactory(options.store),
    layout: new GridLayoutCommandMutationFactory(options.layout),
    structure: new GridStructureCommandMutationFactory(options.structureOwner),
    guards: new HookStructuralCommandGuards(hooks)
  });
  return new CommandManager(
    new TableCommandTransactionHost(options.store),
    planner,
    { ...options, hooks }
  );
}
