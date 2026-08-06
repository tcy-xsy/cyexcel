import type { ClipboardCopyOptions, ClipboardCopyResult, ClipboardPasteOptions } from "../clipboard";
import type { EventBus } from "../events";
import type { HookManager } from "../hooks";
import type { ContextMenuService } from "../menu";
import type { Disposer } from "../types";
import type { CyExcelPlugin, PluginContext } from "./types";

export interface PluginManagerOptions {
  readonly hooks: HookManager;
  readonly events: EventBus;
  readonly contextMenu: ContextMenuService;
  readonly getCell: (row: number, column: number) => unknown;
  readonly copy: (options?: ClipboardCopyOptions) => Promise<ClipboardCopyResult>;
  readonly paste: (options?: ClipboardPasteOptions) => Promise<void>;
  readonly insertRows: (start: number, count?: number) => void;
  readonly deleteRows: (start: number, count?: number) => void;
  readonly moveRows: (start: number, count: number, target: number) => void;
}

interface InstalledPlugin {
  readonly disposers: Disposer[];
  readonly cleanup: Disposer | undefined;
}

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

  constructor(options: PluginManagerOptions) {
    this.#options = options;
  }

  use(plugin: CyExcelPlugin): Disposer {
    const { hooks, events, contextMenu, getCell, copy, paste, insertRows, deleteRows, moveRows } = this.#options;
    const disposers: Disposer[] = [];
    const context: PluginContext = {
      hooks: {
        tap(hookName, meta, handler) {
          // Pure passthrough to the shared HookManager — the plugin author
          // already got checked against PluginContext's own `tap` signature
          // at their call site, so re-unifying two independent sets of
          // generics here has nothing left to verify.
          const dispose = hooks.tap(hookName, meta, handler as never);
          disposers.push(dispose);
          return dispose;
        },
        tapAsync(hookName, meta, handler) {
          const dispose = hooks.tapAsync(hookName, meta, handler as never);
          disposers.push(dispose);
          return dispose;
        }
      },
      events: {
        on(eventName, listener) {
          const dispose = events.on(eventName, listener);
          disposers.push(dispose);
          return dispose;
        }
      },
      contextMenu: {
        registerProvider(meta, provider) {
          const dispose = contextMenu.registerProvider(meta, provider);
          disposers.push(dispose);
          return dispose;
        }
      },
      getCell,
      copy,
      paste,
      insertRows,
      deleteRows,
      moveRows
    };

    const cleanup = plugin.setup(context) ?? undefined;
    const installed: InstalledPlugin = { disposers, cleanup };
    this.#installed.push(installed);

    let disposed = false;
    return () => {
      if (disposed) {
        return;
      }
      disposed = true;
      this.#teardown(installed);
      const index = this.#installed.indexOf(installed);
      if (index >= 0) {
        this.#installed.splice(index, 1);
      }
    };
  }

  /** Uninstalls every currently-installed plugin. Safe to call repeatedly. */
  dispose(): void {
    for (const installed of [...this.#installed]) {
      this.#teardown(installed);
    }
    this.#installed.length = 0;
  }

  #teardown(installed: InstalledPlugin): void {
    installed.cleanup?.();
    for (const dispose of installed.disposers) {
      dispose();
    }
  }
}
