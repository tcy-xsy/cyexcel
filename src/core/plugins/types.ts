import type { ColumnDefinition as CyExcelColumnConfig } from "../cells";
import type { ClipboardCopyOptions, ClipboardCopyResult, ClipboardPasteOptions } from "../clipboard";
import type { HookHandlerMeta } from "../hooks";
import type { ContextMenuProvider, ContextMenuProviderMeta } from "../menu";
import type { SheetMenuCommands } from "../menu/types";
import type { ShortcutHandler } from "../shortcuts";
import type { Disposer } from "../types";

/**
 * The capability surface handed to a plugin's `setup()` — intentionally
 * not the real `HookManager`/`EventBus`/`ContextMenuService` instances
 * (doc/customization-and-hooks.md §11.2: "插件通过公开能力扩展 cyexcel，
 * 不允许访问任意内部私有对象"). Every registration method here is
 * PluginManager's own wrapper that also tracks the returned disposer for
 * that plugin (see PluginManager.use()), not a passthrough.
 */
export interface PluginContext {
  readonly sheetMenu?: SheetMenuCommands;
  readonly hooks: {
    tap<TContext = unknown, TValue = unknown>(
      hookName: string,
      meta: HookHandlerMeta,
      handler: ((context: TContext) => void) | ((value: TValue, context: TContext) => TValue)
    ): Disposer;
    tapAsync<TContext = unknown, TValue = unknown>(
      hookName: string,
      meta: HookHandlerMeta,
      handler:
        | ((context: TContext, signal: AbortSignal) => Promise<void>)
        | ((value: TValue, context: TContext, signal: AbortSignal) => Promise<TValue>)
    ): Disposer;
  };
  readonly events: {
    on<TPayload = unknown>(eventName: string, listener: (payload: TPayload) => void): Disposer;
  };
  readonly contextMenu: {
    registerProvider(meta: ContextMenuProviderMeta, provider: ContextMenuProvider): Disposer;
  };
  readonly getCell: (row: number, column: number) => unknown;
  readonly copy: (options?: ClipboardCopyOptions) => Promise<ClipboardCopyResult>;
  readonly paste: (options?: ClipboardPasteOptions) => Promise<void>;
  /** Inserts `count` blank rows at `start` — the same public Command path `sheet.insertRows()` uses, not direct Store access (doc/customization-and-hooks.md §11.2/re-实现 #4). */
  readonly insertRows: (start: number, count?: number) => void;
  readonly deleteRows: (start: number, count?: number) => void;
  readonly moveRows: (start: number, count: number, target: number) => void;
  /** Same public Command path `sheet.insertColumns()` uses — see `insertRows` above; unlike it, there's no count-only overload (Step 17: a column needs at least an `id`). */
  readonly insertColumns: (start: number, columns: readonly CyExcelColumnConfig[]) => void;
  readonly deleteColumns: (start: number, count?: number) => void;
  readonly moveColumns: (start: number, count: number, target: number) => void;
  readonly hideColumn: (column: number) => void;
  /** Same public shortcut registry `sheet.registerShortcut()` uses — see doc/progress.md's Step 18 entry for the combo syntax and conflict policy. */
  readonly registerShortcut: (combo: string, handler: ShortcutHandler) => Disposer;
}

export interface CyExcelPlugin {
  readonly name: string;
  /** A returned Disposer is called on uninstall, in addition to (not instead of) auto-disposing every tap/listener/provider registered through `context` — see PluginManager. */
  setup(context: PluginContext): void | Disposer;
}
