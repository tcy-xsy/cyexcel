import type { ContextMenuItem, ContextMenuRenderContext, ContextMenuRenderer } from "./types";

type ActionItem = Extract<ContextMenuItem, { readonly type: "action" }>;

function isNavigable(item: ContextMenuItem): item is ActionItem {
  return item.type === "action" && !item.disabled && !item.hidden;
}

const HIGHLIGHT_BACKGROUND = "#eef2ff";

/**
 * The built-in DOM renderer — a plain, unstyled-beyond-basics popup with
 * mouse and keyboard navigation. Not installed by default (see
 * ContextMenuService's `rendererFactory` option); a custom renderer can
 * replace it entirely, following the same mount/update/focus/destroy
 * contract.
 *
 * Highlight changes (arrow-key nav, mouse hover) only ever toggle a style
 * on the existing element via `#setHighlighted()` — they never rebuild the
 * DOM. A real browser only fires the synthetic `click` event when the same
 * node receives both `mousedown` and `mouseup`; an earlier version of this
 * renderer called a full rebuild from the `mouseenter` handler, which
 * removed and replaced the item the mouse was already over, silently
 * swallowing every real (Playwright or human) click on menu items — see
 * doc/progress.md's Step 11 entry.
 */
export class DefaultMenuRenderer implements ContextMenuRenderer {
  #root: HTMLDivElement | null = null;
  #context: ContextMenuRenderContext | null = null;
  #highlightedId: string | null = null;
  readonly #itemElements = new Map<string, HTMLElement>();

  mount(container: HTMLElement, context: ContextMenuRenderContext): void {
    this.#context = context;
    const root = document.createElement("div");
    root.setAttribute("role", "menu");
    root.tabIndex = -1;
    root.style.position = "absolute";
    root.style.left = `${context.anchorPoint.x}px`;
    root.style.top = `${context.anchorPoint.y}px`;
    root.style.background = "#ffffff";
    root.style.border = "1px solid #d0d0d0";
    root.style.boxShadow = "0 2px 8px rgba(0, 0, 0, 0.15)";
    root.style.padding = "4px 0";
    root.style.minWidth = "160px";
    root.style.fontSize = "13px";
    root.style.outline = "none";
    root.addEventListener("keydown", this.#handleKeyDown);
    container.appendChild(root);
    this.#root = root;
    this.#build();
    this.#clampToBounds(context.bounds);
  }

  update(context: ContextMenuRenderContext): void {
    this.#context = context;
    this.#build();
  }

  focus(): void {
    this.#root?.focus();
  }

  destroy(): void {
    this.#root?.removeEventListener("keydown", this.#handleKeyDown);
    this.#root?.remove();
    this.#root = null;
    this.#context = null;
    this.#highlightedId = null;
    this.#itemElements.clear();
  }

  /**
   * Shifts the just-mounted (now real-layout, measurable) root left/up by
   * just enough to stay fully within `bounds`, if it was given one — a
   * plain anchor-point placement with no awareness of available space
   * would otherwise overflow past whatever ancestor is doing the actual
   * clipping (typically the scroller's `overflow: auto`), and the
   * overflowing portion isn't just invisible, it's also unclickable (a
   * real user report: "在右侧插入列" was silently doing nothing — the
   * click landed past the clip boundary, on nothing). Only ever shifts
   * toward the anchor point's own side, never past `bounds`'s opposite
   * edge — a menu wider/taller than `bounds` itself still clips at
   * `bounds.left`/`bounds.top` rather than being pushed further off
   * in the other direction.
   */
  #clampToBounds(bounds: DOMRect | null): void {
    if (!bounds || !this.#root) {
      return;
    }
    const rect = this.#root.getBoundingClientRect();
    let deltaX = Math.min(0, bounds.right - rect.right);
    deltaX = Math.max(deltaX, bounds.left - rect.left);
    let deltaY = Math.min(0, bounds.bottom - rect.bottom);
    deltaY = Math.max(deltaY, bounds.top - rect.top);
    if (deltaX === 0 && deltaY === 0) {
      return;
    }
    const currentLeft = Number.parseFloat(this.#root.style.left) || 0;
    const currentTop = Number.parseFloat(this.#root.style.top) || 0;
    this.#root.style.left = `${currentLeft + deltaX}px`;
    this.#root.style.top = `${currentTop + deltaY}px`;
  }

  #handleKeyDown = (event: KeyboardEvent): void => {
    const context = this.#context;
    if (!context) {
      return;
    }

    if (event.key === "Escape") {
      event.preventDefault();
      context.actions.close();
      return;
    }

    if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      event.preventDefault();
      const navigable = context.items.filter(isNavigable);
      if (navigable.length === 0) {
        return;
      }
      const currentIndex = navigable.findIndex((item) => item.id === this.#highlightedId);
      const delta = event.key === "ArrowDown" ? 1 : -1;
      const nextIndex = (currentIndex + delta + navigable.length) % navigable.length;
      this.#setHighlighted(navigable[nextIndex]?.id ?? null);
      return;
    }

    if (event.key === "Enter" && this.#highlightedId) {
      event.preventDefault();
      context.actions.run(this.#highlightedId);
    }
  };

  /** Full (re)build — only called from mount()/update(), never from a hover/keyboard highlight change (see the class doc comment). */
  #build(): void {
    const context = this.#context;
    if (!this.#root || !context) {
      return;
    }
    this.#root.replaceChildren();
    this.#itemElements.clear();
    for (const item of context.items) {
      this.#root.appendChild(this.#createItemElement(item, context));
    }
  }

  #createItemElement(item: ContextMenuItem, context: ContextMenuRenderContext): HTMLElement {
    if (item.type === "separator") {
      const el = document.createElement("div");
      el.setAttribute("role", "separator");
      el.style.height = "1px";
      el.style.margin = "4px 0";
      el.style.background = "#e5e5e5";
      return el;
    }

    if (item.type === "submenu") {
      // Nested items render inline under their group label — a full flyout
      // submenu (hover-to-open, its own positioning) is out of scope for
      // this step; see doc/progress.md's Step 11 remaining-risk note.
      const el = document.createElement("div");
      el.setAttribute("role", "group");
      const label = document.createElement("div");
      label.textContent = item.label;
      label.style.padding = "6px 12px";
      label.style.fontWeight = "600";
      label.style.color = "#6b7280";
      el.appendChild(label);
      for (const child of item.children) {
        el.appendChild(this.#createItemElement(child, context));
      }
      return el;
    }

    if (item.hidden) {
      const el = document.createElement("div");
      el.style.display = "none";
      return el;
    }

    const busy = context.busyItemId === item.id;
    const menuIsBusy = context.busyItemId !== null;
    const interactive = !item.disabled && !menuIsBusy;

    const el = document.createElement("div");
    el.setAttribute("role", "menuitem");
    el.dataset["itemId"] = item.id;
    el.textContent = busy ? `${item.label}…` : item.checked ? `✓ ${item.label}` : item.label;
    el.style.padding = "6px 12px";
    el.style.cursor = interactive ? "pointer" : "default";
    el.style.opacity = item.disabled || (menuIsBusy && !busy) ? "0.5" : "1";
    el.style.pointerEvents = interactive ? "auto" : "none";
    el.style.background = this.#highlightedId === item.id ? HIGHLIGHT_BACKGROUND : "transparent";

    if (interactive) {
      this.#itemElements.set(item.id, el);
      el.addEventListener("mouseenter", () => this.#setHighlighted(item.id));
      el.addEventListener("click", () => {
        context.actions.run(item.id);
      });
    }

    return el;
  }

  /** Toggles background on the previously/newly highlighted elements directly — never rebuilds the DOM (see the class doc comment for why that matters). */
  #setHighlighted(id: string | null): void {
    if (this.#highlightedId) {
      const previous = this.#itemElements.get(this.#highlightedId);
      if (previous) {
        previous.style.background = "transparent";
      }
    }
    this.#highlightedId = id;
    if (id) {
      const next = this.#itemElements.get(id);
      if (next) {
        next.style.background = HIGHLIGHT_BACKGROUND;
      }
    }
  }
}
