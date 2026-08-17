import type { HookDefinition, HookRegistry } from "./types";

/**
 * One exported constant per hook name this project actually fires
 * somewhere — not the full doc/customization-and-hooks.md §5 catalog.
 * `CommandManager`/`ColumnManager` import the specific constants they need
 * for their own standalone fallback registry (when no shared HookManager
 * is injected), and `PRODUCTION_HOOK_REGISTRY` below composes the same
 * constants for CyExcelCore's shared instance — one definition per name,
 * never two registries disagreeing about a hook's semantics.
 */
export const BEFORE_INIT_HOOK: HookDefinition = { mode: "series", async: false, isolate: false };
export const BEFORE_DESTROY_HOOK: HookDefinition = { mode: "series", async: false, isolate: true };
export const BEFORE_COMMAND_HOOK: HookDefinition = { mode: "bail", async: false, isolate: false };
export const BEFORE_TRANSACTION_COMMIT_HOOK: HookDefinition = { mode: "bail", async: false, isolate: false };
export const BEFORE_ROWS_CHANGE_HOOK: HookDefinition = { mode: "bail", async: false, isolate: false };
export const BEFORE_COLUMNS_CHANGE_HOOK: HookDefinition = { mode: "bail", async: false, isolate: false };
export const RESOLVE_COLUMN_HOOK: HookDefinition = { mode: "waterfall", async: false, isolate: true };
export const BEFORE_CONTEXT_MENU_OPEN_HOOK: HookDefinition = { mode: "bail", async: false, isolate: false };
export const EXTEND_CONTEXT_MENU_HOOK: HookDefinition = { mode: "waterfall", async: false, isolate: true };
/** Reports other hooks' isolated errors. HookManager special-cases this name so a failure in it never re-invokes itself. */
export const ON_HOOK_ERROR_HOOK: HookDefinition = { mode: "series", async: false, isolate: true };

export const PRODUCTION_HOOK_REGISTRY: HookRegistry = {
  beforeInit: BEFORE_INIT_HOOK,
  beforeDestroy: BEFORE_DESTROY_HOOK,
  beforeCommand: BEFORE_COMMAND_HOOK,
  beforeTransactionCommit: BEFORE_TRANSACTION_COMMIT_HOOK,
  beforeRowsChange: BEFORE_ROWS_CHANGE_HOOK,
  beforeColumnsChange: BEFORE_COLUMNS_CHANGE_HOOK,
  resolveColumn: RESOLVE_COLUMN_HOOK,
  beforeContextMenuOpen: BEFORE_CONTEXT_MENU_OPEN_HOOK,
  extendContextMenu: EXTEND_CONTEXT_MENU_HOOK,
  onHookError: ON_HOOK_ERROR_HOOK
};
