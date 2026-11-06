import type { ColumnDefinition as CyExcelColumnConfig } from "../cells";
import type { ClipboardCopyOptions, ClipboardCopyResult, ClipboardPasteOptions } from "../clipboard";
import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { EventBus } from "../events";
import type { HookManager } from "../hooks";
import { LifetimeScope } from "../lifecycle";
import type { ContextMenuService } from "../menu";
import type { SheetMenuCommands } from "../menu/types";
import type { ShortcutHandler } from "../shortcuts";
import type { Disposer } from "../types";
import type { CyExcelPlugin, PluginContext } from "./types";

export interface PluginManagerOptions {
  readonly sheetMenu?: SheetMenuCommands;
  readonly hooks: HookManager;
  readonly events: EventBus;
  readonly contextMenu: ContextMenuService;
  readonly getCell: (row: number, column: number) => unknown;
  readonly copy: (options?: ClipboardCopyOptions) => Promise<ClipboardCopyResult>;
  readonly paste: (options?: ClipboardPasteOptions) => Promise<void>;
  readonly insertRows: (start: number, count?: number) => void;
  readonly deleteRows: (start: number, count?: number) => void;
  readonly moveRows: (start: number, count: number, target: number) => void;
  readonly insertColumns: (start: number, columns: readonly CyExcelColumnConfig[]) => void;
  readonly deleteColumns: (start: number, count?: number) => void;
  readonly moveColumns: (start: number, count: number, target: number) => void;
  readonly hideColumn: (column: number) => void;
  readonly registerShortcut: (combo: string, handler: ShortcutHandler) => Disposer;
}

interface InstalledPlugin {
  readonly name: string;
  readonly scope: LifetimeScope;
}

type PluginManagerState = "registering" | "active" | "disposed";
type RuntimeCapabilityName =
  | "sheetMenu"
  | "getCell"
  | "copy"
  | "paste"
  | "insertRows"
  | "deleteRows"
  | "moveRows"
  | "insertColumns"
  | "deleteColumns"
  | "moveColumns"
  | "hideColumn";

/**
 * Installs plugins and guarantees uniform teardown (doc/customization-
 * and-hooks.md §10/§11.2 + Step 11's exit criteria): every `hooks.tap`/
 * `events.on`/`contextMenu.registerProvider` call made through the
 * `PluginContext` handed to a plugin's `setup()` is recorded automatically,
 * so uninstalling a plugin cleans all of it up even if the plugin's own
 * returned cleanup function forgets something.
 */
export class PluginManager {
  readonly #options: PluginManagerOptions;
  readonly #installed: InstalledPlugin[] = [];
  #state: PluginManagerState = "registering";

  constructor(options: PluginManagerOptions) {
    this.#options = options;
  }

  get state(): PluginManagerState {
    return this.#state;
  }

  use(plugin: CyExcelPlugin): Disposer {
    if (this.#state === "disposed") {
      throw new CyExcelError(CyExcelErrorCode.PluginDisposed, "Cannot install a plugin after PluginManager is disposed");
    }
    if (this.#state !== "registering") {
      throw new CyExcelError(
        CyExcelErrorCode.PluginRuntimeNotReady,
        "Plugins must be installed before PluginManager.activate()"
      );
    }

    const {
      hooks,
      events,
      contextMenu,
      getCell,
      copy,
      paste,
      insertRows,
      deleteRows,
      moveRows,
      insertColumns,
      deleteColumns,
      moveColumns,
      hideColumn,
      registerShortcut
    } = this.#options;
    const scope = new LifetimeScope({
      onDisposeError: error => events.emit("error", { error, eventName: `plugin:${plugin.name}:dispose` })
    });
    let pluginDisposed = false;
    const context: PluginContext = {
      ...(this.#options.sheetMenu === undefined ? {} : { sheetMenu: {
        getState: this.#runtime("sheetMenu", this.#options.sheetMenu.getState, () => pluginDisposed),
        run: this.#runtime("sheetMenu", this.#options.sheetMenu.run, () => pluginDisposed)
      } }),
      hooks: {
        tap(hookName, meta, handler) {
          // Pure passthrough to the shared HookManager — the plugin author
          // already got checked against PluginContext's own `tap` signature
          // at their call site, so re-unifying two independent sets of
          // generics here has nothing left to verify.
          const dispose = hooks.tap(hookName, meta, handler as never);
          return scope.defer(dispose);
        },
        tapAsync(hookName, meta, handler) {
          const dispose = hooks.tapAsync(hookName, meta, handler as never);
          return scope.defer(dispose);
        }
      },
      events: {
        on(eventName, listener) {
          const dispose = events.on(eventName, listener);
          return scope.defer(dispose);
        }
      },
      contextMenu: {
        registerProvider(meta, provider) {
          const dispose = contextMenu.registerProvider(meta, provider);
          return scope.defer(dispose);
        }
      },
      getCell: this.#runtime("getCell", getCell, () => pluginDisposed),
      copy: this.#runtime("copy", copy, () => pluginDisposed),
      paste: this.#runtime("paste", paste, () => pluginDisposed),
      insertRows: this.#runtime("insertRows", insertRows, () => pluginDisposed),
      deleteRows: this.#runtime("deleteRows", deleteRows, () => pluginDisposed),
      moveRows: this.#runtime("moveRows", moveRows, () => pluginDisposed),
      insertColumns: this.#runtime("insertColumns", insertColumns, () => pluginDisposed),
      deleteColumns: this.#runtime("deleteColumns", deleteColumns, () => pluginDisposed),
      moveColumns: this.#runtime("moveColumns", moveColumns, () => pluginDisposed),
      hideColumn: this.#runtime("hideColumn", hideColumn, () => pluginDisposed),
      registerShortcut: (combo, handler) => {
        if (this.#state === "disposed" || pluginDisposed) {
          throw new CyExcelError(
            CyExcelErrorCode.PluginDisposed,
            "Plugin registration capability registerShortcut() cannot be used after plugin disposal"
          );
        }
        const dispose = registerShortcut(combo, handler);
        return scope.defer(dispose);
      }
    };

    let cleanup: Disposer | undefined;
    try {
      cleanup = plugin.setup(context) ?? undefined;
    } catch (error) {
      scope.dispose();
      throw error;
    }
    if (cleanup) {
      scope.defer(cleanup);
    }
    const installed: InstalledPlugin = { name: plugin.name, scope };
    this.#installed.push(installed);

    let disposed = false;
    return () => {
      if (disposed) {
        return;
      }
      disposed = true;
      pluginDisposed = true;
      this.#teardown(installed);
      const index = this.#installed.indexOf(installed);
      if (index >= 0) {
        this.#installed.splice(index, 1);
      }
    };
  }

  activate(): void {
    if (this.#state === "disposed") {
      throw new CyExcelError(CyExcelErrorCode.PluginDisposed, "Cannot activate a disposed PluginManager");
    }
    this.#state = "active";
  }

  /** Uninstalls every currently-installed plugin. Safe to call repeatedly. */
  dispose(): void {
    if (this.#state === "disposed") {
      return;
    }

    this.#state = "disposed";
    for (const installed of [...this.#installed]) {
      this.#teardown(installed);
    }
    this.#installed.length = 0;
  }

  #runtime<TArgs extends unknown[], TResult>(
    name: RuntimeCapabilityName,
    capability: (...args: TArgs) => TResult,
    isPluginDisposed: () => boolean
  ): (...args: TArgs) => TResult {
    return (...args) => {
      this.#assertRuntimeReady(name, isPluginDisposed());
      return capability(...args);
    };
  }

  #assertRuntimeReady(name: RuntimeCapabilityName, pluginDisposed: boolean): void {
    if (this.#state === "registering") {
      throw new CyExcelError(
        CyExcelErrorCode.PluginRuntimeNotReady,
        `Plugin runtime capability ${name}() is not ready until PluginManager.activate()`
      );
    }
    if (this.#state === "disposed" || pluginDisposed) {
      throw new CyExcelError(
        CyExcelErrorCode.PluginDisposed,
        `Plugin runtime capability ${name}() cannot be used after plugin disposal`
      );
    }
  }

  #teardown(installed: InstalledPlugin): void {
    installed.scope.dispose();
  }
}
