import { HookManager, RESOLVE_COLUMN_HOOK, type HookRegistry } from "../hooks";
import { DEFAULT_COLUMN_WIDTH } from "../layout";
import type { CellTypeRegistry } from "./CellTypeRegistry";
import type { ColumnDefinition, RuntimeColumn } from "./types";

/**
 * 缺省列宽只有一个定义处，在布局层（`LayoutEngine.ts`）。
 *
 * ADR 0004 阶段 4B T8b 之后，「这一列没有显式宽度时算多宽」这个问题由布局配置回答
 * （`LayoutOptions.defaultColumnWidth`），因为它正是那条不变量里的缺省项；这里只是
 * 转出去，好让 `ColumnDefinition.width` 缺省时与布局算的是同一个数。方向是
 * cells → layout（`cells/types.ts` 早就从 `../layout` 取 `CellRect`），反过来会给
 * 布局层加一条它不该有的依赖。
 */
export { DEFAULT_COLUMN_WIDTH };
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
}

/** Fully resolved column state prepared before it becomes the live manager root. */
export type RuntimeColumnRoot = readonly RuntimeColumn[];

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
  #columns: readonly RuntimeColumn[];
  readonly #registry: CellTypeRegistry;
  readonly #hooks: HookManager;

  constructor(options: ColumnManagerOptions) {
    this.#registry = options.registry;
    this.#hooks = options.hooks ?? new HookManager({ registry: STANDALONE_HOOK_REGISTRY });

    this.#columns = this.prepareRoot(options.columns);
  }

  /**
   * Resolves a complete replacement root without changing the columns visible
   * through getColumn()/getColumns(). If resolution throws, the current root
   * is untouched and the partially built local array becomes unreachable.
   */
  prepareRoot(columns: readonly ColumnDefinition[]): RuntimeColumnRoot {
    return ColumnManager.#resolveAll(columns, this.#registry, this.#hooks);
  }

  /** Installs an already-resolved root. This narrow commit step never runs hooks. */
  replaceRoot(root: RuntimeColumnRoot): void {
    this.#columns = root;
  }

  static #resolveAll(
    columns: readonly ColumnDefinition[],
    registry: CellTypeRegistry,
    hooks: HookManager
  ): readonly RuntimeColumn[] {
    return Object.freeze(
      columns.map((definition, index) =>
        ColumnManager.#freezeRuntimeColumn(ColumnManager.#resolveOne(definition, index, registry, hooks))
      )
    );
  }

  static #freezeRuntimeColumn(column: RuntimeColumn): RuntimeColumn {
    const options = column.options?.map((option) => Object.freeze({ ...option }));
    return Object.freeze({
      ...column,
      clipboard: Object.freeze({ ...column.clipboard }),
      ...(options !== undefined ? { options: Object.freeze(options) } : {})
    });
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
      width: definition.width ?? DEFAULT_COLUMN_WIDTH,
      typeName,
      ...(definition.renderer !== undefined ? { renderer: definition.renderer } : {}),
      ...(definition.options !== undefined ? { options: definition.options } : {}),
      ...(definition.createEditor !== undefined ? { createEditor: definition.createEditor } : {}),
      clipboard: {
        copy: definition.clipboard?.copy ?? true,
        paste: definition.clipboard?.paste ?? true
      },
      wrap: definition.wrap ?? false,
      // Three layers, narrowest first: the column decides, else the cell type
      // decides (a type with no total order opts itself out — see
      // `builtins.ts`), else sortable. Same shape as `compare`: the type
      // declares the capability, the column sets the policy.
      sortable: definition.sortable ?? cellType.sortable ?? true,
      filterable: definition.filterable ?? cellType.filterable ?? true,
      version: RESOLVED_VERSION
    };

    return hooks.runWaterfall<RuntimeColumn, ResolveColumnContext>("resolveColumn", merged, { definition, index });
  }
}
