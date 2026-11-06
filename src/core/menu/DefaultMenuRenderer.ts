import type { ContextMenuItem, ContextMenuRenderContext, ContextMenuRenderer } from "./types";

type Entry = Exclude<ContextMenuItem, { type: "separator" }>;
interface Panel {
  element: HTMLDivElement;
  items: readonly ContextMenuItem[];
  entries: Map<string, HTMLElement>;
  highlighted: string | null;
  parentId?: string;
}
const visibleItems = (items: readonly ContextMenuItem[]): ContextMenuItem[] => {
  const result: ContextMenuItem[] = [];
  for (const item of items) {
    if (item.type !== "separator" && item.hidden) continue;
    if (item.type === "separator" && (!result.length || result.at(-1)?.type === "separator")) continue;
    result.push(item);
  }
  if (result.at(-1)?.type === "separator") result.pop();
  return result;
};
let sequence = 0;

/** Independent flyout panels avoid clipping submenus inside the scrollable root menu. */
export class DefaultMenuRenderer implements ContextMenuRenderer {
  #container: HTMLElement | undefined;
  #context: ContextMenuRenderContext | undefined;
  #panels: Panel[] = [];
  #hoverTimer: ReturnType<typeof setTimeout> | undefined;

  mount(container: HTMLElement, context: ContextMenuRenderContext): void {
    this.#container = container;
    this.#context = context;
    this.#buildRoot();
  }

  update(context: ContextMenuRenderContext): void {
    const path = this.#panels.slice(1).map(panel => panel.parentId!);
    this.#context = context;
    this.#removePanels(0);
    this.#buildRoot();
    for (let depth = 0; depth < path.length; depth++) this.#openSubmenu(depth, path[depth]!, false);
    this.focus();
  }

  focus(): void { this.#panels.at(-1)?.element.focus({ preventScroll: true }); }
  destroy(): void {
    this.#cancelHover();
    this.#removePanels(0);
    this.#container = undefined;
    this.#context = undefined;
  }

  #bounds(): { left: number; top: number; right: number; bottom: number; width: number; height: number } {
    const viewport = { left: 0, top: 0, right: window.innerWidth, bottom: window.innerHeight };
    const bounds = this.#context?.bounds;
    const left = Math.max(viewport.left, bounds?.left ?? 0), top = Math.max(viewport.top, bounds?.top ?? 0);
    const right = Math.min(viewport.right, bounds?.right ?? viewport.right), bottom = Math.min(viewport.bottom, bounds?.bottom ?? viewport.bottom);
    return { left, top, right, bottom, width: Math.max(0, right - left), height: Math.max(0, bottom - top) };
  }

  #buildRoot(): void {
    const context = this.#context;
    if (!context) return;
    const panel = this.#createPanel(context.items, 0);
    panel.element.style.left = `${context.anchorPoint.x}px`;
    panel.element.style.top = `${context.anchorPoint.y}px`;
    if (context.error) {
      const error = document.createElement("div");
      error.setAttribute("role", "alert");
      error.textContent = context.error;
      error.style.cssText = "padding:8px 12px;color:#b42318;white-space:normal;line-height:1.5;border-top:1px solid #eee";
      panel.element.append(error);
    }
    this.#clamp(panel.element);
  }

  #createPanel(items: readonly ContextMenuItem[], depth: number, parentId?: string): Panel {
    const element = document.createElement("div");
    element.className = "cyx-context-menu";
    element.setAttribute("role", "menu");
    element.setAttribute("aria-label", depth === 0 ? "表格操作" : "更多操作");
    element.setAttribute("aria-busy", String(this.#context?.busyItemId !== null));
    element.tabIndex = -1;
    const bounds = this.#bounds();
    element.style.cssText = `position:absolute;background:#fff;color:#263238;border:1px solid #dce1e6;border-radius:6px;box-shadow:0 5px 20px #172b4d26;padding:4px;box-sizing:border-box;min-width:${Math.min(216, bounds.width)}px;max-width:${Math.max(0, bounds.width)}px;max-height:${Math.max(0, bounds.height)}px;overflow-y:auto;overflow-x:hidden;font-size:13px;line-height:20px;outline:none;z-index:${100 + depth}`;
    const panel: Panel = { element, items: visibleItems(items), entries: new Map(), highlighted: null, ...(parentId === undefined ? {} : { parentId }) };
    this.#panels.push(panel);
    for (const item of panel.items) {
      if (item.type === "separator") {
        const separator = document.createElement("div");
        separator.setAttribute("role", "separator");
        separator.style.cssText = "height:1px;margin:4px 7px;background:#e8ebee";
        element.append(separator);
        continue;
      }
      const row = document.createElement("div");
      row.id = `cyx-menu-${++sequence}`;
      row.dataset["itemId"] = item.id;
      row.setAttribute("role", item.type === "action" && item.checked !== undefined ? "menuitemcheckbox" : "menuitem");
      if (item.type === "action" && item.checked !== undefined) row.setAttribute("aria-checked", String(item.checked));
      row.setAttribute("aria-label", item.label);
      row.style.cssText = "display:flex;align-items:center;gap:8px;min-height:30px;padding:3px 8px;box-sizing:border-box;border-radius:3px;white-space:nowrap;cursor:default";
      const disabled = Boolean(item.disabled || this.#context?.busyItemId);
      row.setAttribute("aria-disabled", String(disabled));
      if (item.description) row.title = item.description;
      if (disabled) row.style.color = "#a1a8b0";
      const check = document.createElement("span");
      check.textContent = item.type === "action" && item.checked ? "✓" : "";
      check.style.cssText = "width:14px;flex-shrink:0;color:inherit";
      check.setAttribute("aria-hidden", "true");
      const label = document.createElement("span");
      label.textContent = item.label + (this.#context?.busyItemId === item.id ? "…" : "");
      label.style.cssText = "flex:1;overflow:hidden;text-overflow:ellipsis";
      row.append(check, label);
      if (item.type === "submenu") {
        row.setAttribute("aria-haspopup", "menu");
        row.setAttribute("aria-expanded", "false");
        const arrow = document.createElement("span");
        arrow.textContent = "›";
        arrow.setAttribute("aria-hidden", "true");
        row.append(arrow);
      } else if (item.shortcut) {
        const shortcut = document.createElement("span");
        shortcut.textContent = item.shortcut;
        shortcut.style.cssText = "margin-left:18px;font-size:11px;color:#86909c";
        shortcut.setAttribute("aria-hidden", "true");
        row.append(shortcut);
      }
      panel.entries.set(item.id, row);
      row.addEventListener("mouseenter", () => {
        this.#cancelHover();
        this.#highlight(depth, item.id);
        if (this.#panels[depth + 1]?.parentId !== item.id) this.#removePanels(depth + 1);
        if (item.type === "submenu" && !disabled) this.#hoverTimer = setTimeout(() => this.#openSubmenu(depth, item.id, false), 140);
      });
      row.addEventListener("mouseleave", () => this.#cancelHover());
      row.addEventListener("click", event => {
        event.stopPropagation();
        this.#cancelHover();
        if (disabled) return;
        if (item.type === "submenu") this.#openSubmenu(depth, item.id, true);
        else this.#context?.actions.run(item.id);
      });
      element.append(row);
    }
    element.addEventListener("keydown", event => this.#keyDown(event, depth));
    element.addEventListener("contextmenu", event => { event.preventDefault(); event.stopPropagation(); });
    element.addEventListener("mouseenter", () => element.focus({ preventScroll: true }));
    element.addEventListener("scroll", () => { this.#cancelHover(); this.#removePanels(depth + 1); });
    this.#container?.append(element);
    return panel;
  }

  #clamp(element: HTMLElement): void {
    const bounds = this.#bounds(), rect = element.getBoundingClientRect();
    const dx = Math.max(bounds.left - rect.left, Math.min(0, bounds.right - rect.right));
    const dy = Math.max(bounds.top - rect.top, Math.min(0, bounds.bottom - rect.bottom));
    element.style.left = `${(parseFloat(element.style.left) || 0) + dx}px`;
    element.style.top = `${(parseFloat(element.style.top) || 0) + dy}px`;
  }

  #openSubmenu(depth: number, id: string, focus: boolean): void {
    const parent = this.#panels[depth];
    const item = parent?.items.find(entry => entry.id === id);
    const row = parent?.entries.get(id);
    if (!parent || !row || item?.type !== "submenu" || item.disabled || this.#context?.busyItemId) return;
    this.#removePanels(depth + 1);
    const panel = this.#createPanel(item.children, depth + 1, id);
    row.setAttribute("aria-expanded", "true");
    const rect = row.getBoundingClientRect(), origin = this.#container!.getBoundingClientRect();
    const bounds = this.#bounds(), width = panel.element.getBoundingClientRect().width;
    const left = rect.right + width + 4 > bounds.right ? rect.left - width - 4 : rect.right + 4;
    panel.element.style.left = `${left - origin.left}px`;
    panel.element.style.top = `${rect.top - origin.top}px`;
    this.#clamp(panel.element);
    if (focus) {
      this.#highlight(depth + 1, this.#navigable(panel)[0]?.id ?? null);
      panel.element.focus({ preventScroll: true });
    }
  }

  #removePanels(depth: number): void {
    for (const panel of this.#panels.splice(depth)) {
      panel.element.remove();
      if (panel.parentId) this.#panels.at(-1)?.entries.get(panel.parentId)?.setAttribute("aria-expanded", "false");
    }
  }
  #cancelHover(): void { clearTimeout(this.#hoverTimer); this.#hoverTimer = undefined; }
  #navigable(panel: Panel): Entry[] {
    return this.#context?.busyItemId ? [] : panel.items.filter((item): item is Entry => item.type !== "separator" && !item.disabled);
  }
  #highlight(depth: number, id: string | null): void {
    const panel = this.#panels[depth];
    if (!panel) return;
    if (panel.highlighted) panel.entries.get(panel.highlighted)?.style.setProperty("background", "transparent");
    panel.highlighted = id;
    const entry = id ? panel.entries.get(id) : undefined;
    if (entry) {
      entry.style.background = "#edf5f2";
      panel.element.setAttribute("aria-activedescendant", entry.id);
    } else panel.element.removeAttribute("aria-activedescendant");
  }

  #keyDown(event: KeyboardEvent, depth: number): void {
    const panel = this.#panels[depth];
    if (!panel) return;
    const entries = this.#navigable(panel);
    const index = entries.findIndex(item => item.id === panel.highlighted);
    if (!["ArrowDown", "ArrowUp", "ArrowRight", "ArrowLeft", "Home", "End", "Enter", " ", "Escape", "Tab"].includes(event.key)) return;
    event.preventDefault();
    event.stopPropagation();
    this.#cancelHover();
    if (event.key === "Tab") { this.#context?.actions.close(); return; }
    if (event.key === "Escape" || event.key === "ArrowLeft") {
      if (depth > 0) { this.#removePanels(depth); this.#panels[depth - 1]?.element.focus({ preventScroll: true }); }
      else if (event.key === "Escape") this.#context?.actions.close();
      return;
    }
    if (["ArrowDown", "ArrowUp", "Home", "End"].includes(event.key)) {
      this.#removePanels(depth + 1);
      const next = event.key === "Home" ? 0 : event.key === "End" ? entries.length - 1 : index < 0 ? (event.key === "ArrowUp" ? entries.length - 1 : 0) : (index + (event.key === "ArrowDown" ? 1 : -1) + entries.length) % entries.length;
      const entry = entries[next];
      this.#highlight(depth, entry?.id ?? null);
      if (entry) panel.entries.get(entry.id)?.scrollIntoView({ block: "nearest" });
      return;
    }
    const item = entries[index];
    if (item?.type === "submenu") this.#openSubmenu(depth, item.id, true);
    else if (item && event.key !== "ArrowRight") this.#context?.actions.run(item.id);
  }
}
