import { HookManager, RESOLVE_COLUMN_HOOK, type HookRegistry } from "../hooks";
import type { CellTypeRegistry } from "./CellTypeRegistry";
import type { ColumnDefinition, RuntimeColumn } from "./types";

const DEFAULT_WIDTH = 120;
const DEFAULT_TYPE = "text";
const RESOLVED_VERSION = 1;

/** Context for the `resolveColumn` Hook — which declared column this waterfall run is resolving. */
export interface ResolveColumnContext {
  readonly definition: ColumnDefinition;
  readonly index: number;
}

/** Used when no shared HookManager is injected — just this module's own hook name (see CommandManager's identical fallback pattern). */
const STANDALONE_HOOK_REGISTRY: HookRegistry = {
  resolveColumn: RESOLVE_COLUMN_HOOK
};

export interface ColumnManagerOptions {
  readonly columns: readonly ColumnDefinition[];
  readonly registry: CellTypeRegistry;
  /**
   * Shares a HookManager with the rest of a CyExcelCore instance so
   * `sheet.hooks.tap("resolveColumn", ...)` actually participates in real
   * column resolution. Defaults to a private HookManager scoped to just
   * `resolveColumn`.
   */
  readonly hooks?: HookManager;
  /**
   * Sugar for `hooks.tap("resolveColumn", ...)`, kept for backward
   * compatibility with Step 8's original shape — run once per column right
   * after the type-chain merge.
   */
  readonly resolveColumn?: (column: RuntimeColumn) => RuntimeColumn;
}

/**
 * Resolves declared ColumnDefinitions into cached RuntimeColumns once, at
 * construction — merge order is: registry type chain (built-in defaults ->
 * extends chain -> the column's own registered type) -> this column's own
 * `field`/`width`/`renderer` -> the `resolveColumn` Hook (waterfall, so
 * every registered tap runs in priority order, each seeing the previous
 * tap's output). getColumn()/getColumns() only ever read the cache;
 * painting never re-merges.
 */
export class ColumnManager {
  readonly #columns: readonly RuntimeColumn[];

  constructor(options: ColumnManagerOptions) {
    const hooks = options.hooks ?? new HookManager({ registry: STANDALONE_HOOK_REGISTRY });

    if (options.resolveColumn) {
      const resolveColumn = options.resolveColumn;
      hooks.tap("resolveColumn", { name: "legacy-resolveColumn-option" }, (column: RuntimeColumn) =>
        resolveColumn(column)
      );
    }

    this.#columns = options.columns.map((definition, index) =>
      ColumnManager.#resolveOne(definition, index, options.registry, hooks)
    );
  }

  getColumn(index: number): RuntimeColumn {
    const column = this.#columns[index];
    if (!column) {
      throw new RangeError(`Column index ${index} is out of range`);
    }
    return column;
  }

  getColumns(): readonly RuntimeColumn[] {
    return this.#columns;
  }

  get columnCount(): number {
    return this.#columns.length;
  }

  static #resolveOne(
    definition: ColumnDefinition,
    index: number,
    registry: CellTypeRegistry,
    hooks: HookManager
  ): RuntimeColumn {
    const typeName = definition.type ?? DEFAULT_TYPE;
    const cellType = registry.resolve(typeName);

    const merged: RuntimeColumn = {
      ...cellType,
      id: definition.id,
      field: definition.field ?? definition.id,
      width: definition.width ?? DEFAULT_WIDTH,
      typeName,
      ...(definition.renderer !== undefined ? { renderer: definition.renderer } : {}),
      ...(definition.options !== undefined ? { options: definition.options } : {}),
      ...(definition.createEditor !== undefined ? { createEditor: definition.createEditor } : {}),
      clipboard: {
        copy: definition.clipboard?.copy ?? true,
        paste: definition.clipboard?.paste ?? true
      },
      version: RESOLVED_VERSION
    };

    return hooks.runWaterfall<RuntimeColumn, ResolveColumnContext>("resolveColumn", merged, { definition, index });
  }
}
