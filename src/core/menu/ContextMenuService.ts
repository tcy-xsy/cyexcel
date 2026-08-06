import {
  BEFORE_CONTEXT_MENU_OPEN_HOOK,
  EXTEND_CONTEXT_MENU_HOOK,
  HookManager,
  createCancellableContext,
  type HookRegistry
} from "../hooks";
import type { Disposer } from "../types";
import { DefaultMenuRenderer } from "./DefaultMenuRenderer";
import type {
  ContextMenuContext,
  ContextMenuItem,
  ContextMenuProvider,
  ContextMenuProviderMeta,
  ContextMenuRenderContext,
  ContextMenuRenderer,
  ContextMenuRendererFactory
} from "./types";

/** Used when no shared HookManager is injected (see CommandManager/ColumnManager's identical fallback pattern). */
const STANDALONE_HOOK_REGISTRY: HookRegistry = {
  beforeContextMenuOpen: BEFORE_CONTEXT_MENU_OPEN_HOOK,
  extendContextMenu: EXTEND_CONTEXT_MENU_HOOK
};

type ActionItem = Extract<ContextMenuItem, { readonly type: "action" }>;

function findAction(items: readonly ContextMenuItem[], id: string): ActionItem | undefined {
  for (const item of items) {
    if (item.type === "action" && item.id === id) {
      return item;
    }
    if (item.type === "submenu") {
      const found = findAction(item.children, id);
      if (found) {
        return found;
      }
    }
  }
  return undefined;
}

export interface ContextMenuServiceOptions {
  /** Where the active renderer mounts — same convention as EditorManager/DOMCellManager's `container` option. */
  readonly container: HTMLElement;
  readonly hooks?: HookManager;
  /** Defaults to DefaultMenuRenderer. */
  readonly rendererFactory?: ContextMenuRendererFactory;
  /** Called once whenever the menu closes, for any reason — CyExcelCore passes `() => this.#contentCanvas.focus()`. */
  readonly restoreFocus?: () => void;
}

/**
 * Providers only ever *append* items (see registerProvider); deleting,
 * replacing, or disabling items already in the merged list is the
 * `extendContextMenu` waterfall Hook's job — two extension mechanisms,
 * each responsible for one direction, instead of providers needing to
 * cooperate with each other (doc/customization-and-hooks.md §3.6).
 */
export class ContextMenuService {
  readonly #container: HTMLElement;
  readonly #hooks: HookManager;
  readonly #rendererFactory: ContextMenuRendererFactory;
  readonly #restoreFocus: (() => void) | undefined;
  readonly #providers: Array<{ readonly seq: number; readonly meta: ContextMenuProviderMeta; readonly provider: ContextMenuProvider }> =
    [];
  #providerSeq = 0;

  #activeRenderer: ContextMenuRenderer | null = null;
  #activeItems: readonly ContextMenuItem[] = [];
  #activeContext: ContextMenuContext | null = null;
  #busyItemId: string | null = null;
  #activeAbortController: AbortController | null = null;

  constructor(options: ContextMenuServiceOptions) {
    this.#container = options.container;
    this.#hooks = options.hooks ?? new HookManager({ registry: STANDALONE_HOOK_REGISTRY });
    this.#rendererFactory = options.rendererFactory ?? { create: () => new DefaultMenuRenderer() };
    this.#restoreFocus = options.restoreFocus;
  }

  registerProvider(meta: ContextMenuProviderMeta, provider: ContextMenuProvider): Disposer {
    const entry = { seq: this.#providerSeq, meta, provider };
    this.#providerSeq += 1;
    this.#providers.push(entry);

    let disposed = false;
    return () => {
      if (disposed) {
        return;
      }
      disposed = true;
      const index = this.#providers.indexOf(entry);
      if (index >= 0) {
        this.#providers.splice(index, 1);
      }
    };
  }

  /** Closes any currently-open menu (replace), runs `beforeContextMenuOpen`/`extendContextMenu`, and mounts a fresh renderer. Returns false without opening anything if cancelled or the resulting menu is empty. */
  open(context: ContextMenuContext): boolean {
    this.close();

    const bailContext = createCancellableContext({ context });
    this.#hooks.runBail("beforeContextMenuOpen", bailContext);
    if (bailContext.cancelled) {
      return false;
    }

    const items = this.#buildMenu(context);
    if (items.length === 0) {
      return false;
    }

    this.#activeContext = context;
    this.#activeItems = items;
    this.#busyItemId = null;

    const renderer = this.#rendererFactory.create();
    this.#activeRenderer = renderer;
    renderer.mount(this.#container, this.#renderContext());
    renderer.focus();
    return true;
  }

  /** Aborts any in-flight action, destroys the current renderer, and restores focus. Safe to call when nothing is open. */
  close(): void {
    if (!this.#activeRenderer) {
      return;
    }
    this.#activeAbortController?.abort();
    this.#activeAbortController = null;

    const renderer = this.#activeRenderer;
    this.#activeRenderer = null;
    this.#activeItems = [];
    this.#activeContext = null;
    this.#busyItemId = null;

    renderer.destroy();
    this.#restoreFocus?.();
  }

  isOpen(): boolean {
    return this.#activeRenderer !== null;
  }

  dispose(): void {
    this.close();
  }

  #buildMenu(context: ContextMenuContext): readonly ContextMenuItem[] {
    const applicable = this.#providers
      .filter((entry) => !entry.meta.target || entry.meta.target === context.target)
      .slice()
      .sort((a, b) => (b.meta.priority ?? 0) - (a.meta.priority ?? 0) || a.seq - b.seq);

    const merged = applicable.flatMap((entry) => entry.provider(context));
    return this.#hooks.runWaterfall("extendContextMenu", merged, context);
  }

  #renderContext(): ContextMenuRenderContext {
    return {
      items: this.#activeItems,
      busyItemId: this.#busyItemId,
      anchorPoint: this.#activeContext?.point ?? { x: 0, y: 0 },
      actions: {
        run: (itemId) => this.#runAction(itemId),
        close: () => this.close()
      }
    };
  }

  #runAction(itemId: string): void {
    const item = findAction(this.#activeItems, itemId);
    const context = this.#activeContext;
    if (!item || !context || item.disabled) {
      return;
    }

    const controller = new AbortController();
    this.#activeAbortController = controller;
    const result = item.run(context, controller.signal);

    if (!isPromiseLike(result)) {
      this.close();
      return;
    }

    this.#busyItemId = itemId;
    this.#activeRenderer?.update(this.#renderContext());

    result.then(
      () => {
        // The menu may already have been closed/replaced (which aborts this
        // controller) while the action was in flight — don't commit
        // anything further in that case, per the "AbortSignal 取消后不再
        // 提交结果" exit criterion.
        if (controller.signal.aborted) {
          return;
        }
        this.close();
      },
      () => {
        if (controller.signal.aborted) {
          return;
        }
        this.#busyItemId = null;
        this.#activeRenderer?.update(this.#renderContext());
      }
    );
  }
}

function isPromiseLike(value: unknown): value is Promise<void> {
  return typeof value === "object" && value !== null && typeof (value as Promise<void>).then === "function";
}
