import type { SelectionSnapshot } from "../interaction";

export type ContextMenuTarget = "cell" | "row-header" | "column-header" | "corner" | "blank";

/**
 * Read-only description of what a right-click (or programmatic open())
 * landed on. `row`/`column` are only present when the target actually
 * resolves to one — e.g. "corner" never has either, "blank" never has
 * either (click past the data extent), "row-header" has `row` but not
 * `column`.
 */
export interface ContextMenuContext {
  readonly target: ContextMenuTarget;
  readonly row?: number;
  readonly column?: number;
  readonly selection: SelectionSnapshot;
  readonly point: { readonly x: number; readonly y: number };
}

export type ContextMenuItem =
  | { readonly type: "separator"; readonly id: string }
  | {
      readonly type: "action";
      readonly id: string;
      readonly label: string;
      readonly disabled?: boolean;
      readonly checked?: boolean;
      readonly hidden?: boolean;
      run(context: ContextMenuContext, signal: AbortSignal): void | Promise<void>;
    }
  | {
      readonly type: "submenu";
      readonly id: string;
      readonly label: string;
      readonly children: readonly ContextMenuItem[];
    };

export interface ContextMenuProviderMeta {
  readonly name: string;
  /** Higher runs first when providers are merged; same priority preserves registration order. Defaults to 0. */
  readonly priority?: number;
  /** Only invoked when the open()'d context's target matches. Omit to run for every target. */
  readonly target?: ContextMenuTarget;
}

export type ContextMenuProvider = (context: ContextMenuContext) => readonly ContextMenuItem[];

export interface ContextMenuRenderContext {
  readonly items: readonly ContextMenuItem[];
  /** The `id` of the action item currently awaiting its async run(), or null when nothing is in flight. */
  readonly busyItemId: string | null;
  /** Where the menu was opened — the same point classifyContextMenuTarget() was given. */
  readonly anchorPoint: { readonly x: number; readonly y: number };
  readonly actions: {
    readonly run: (itemId: string) => void;
    readonly close: () => void;
  };
}

/**
 * Shape matches `DomCellRenderer`/`DomHeaderRenderer`'s established
 * mount/update/destroy lifecycle (src/core/cells/types.ts), plus `focus()`
 * so the menu can immediately receive keyboard navigation once opened.
 */
export interface ContextMenuRenderer {
  mount(container: HTMLElement, context: ContextMenuRenderContext): void;
  update(context: ContextMenuRenderContext): void;
  focus(): void;
  destroy(): void;
}

export interface ContextMenuRendererFactory {
  create(): ContextMenuRenderer;
}
