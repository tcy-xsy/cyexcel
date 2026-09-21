import type { CyExcelCore } from "../CyExcelCore";
import { computeOverflow } from "./overflow";
import { defaultToolbarSchema } from "./schema";
import { resolveItemState } from "./state";
import * as icons from "./icons";
import { TOOLBAR_GROUPS, type ToolbarContext, type ToolbarItem, type ToolbarOptions } from "./types";
import * as styles from "./styles";

export interface ToolbarHandle {
  /** 立刻重算一遍所有项的状态。事件驱动之外的手动出口，宿主一般用不到。 */
  refresh(): void;
  dispose(): void;
}

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className !== undefined) {
    node.className = className;
  }
  return node;
}

function sortItems(items: readonly ToolbarItem[]): readonly ToolbarItem[] {
  return [...items].sort((a, b) => {
    const groupDelta = TOOLBAR_GROUPS.indexOf(a.group) - TOOLBAR_GROUPS.indexOf(b.group);
    return groupDelta !== 0 ? groupDelta : a.order - b.order;
  });
}

/**
 * 一条工具栏。
 *
 * **显式引入的可选模块**（`cyexcel/toolbar`），不 import 就完全不进产物——与
 * `createBasicContextMenuPlugin` 同一个模式，理由也一样：业务 UI 不该焊进 core。
 *
 * 项从 schema 来（数据，不是代码），所以宿主可以整体替换或追加，不必改我们的源码。
 * 状态（高亮 / 禁用 / 当前值）在每次选区或数据变化后重算一遍——照 Univer 的
 * `activated$`/`disabled$` 那套语义，但用函数而不是 Observable，因为我们没有 rxjs，
 * 而几十个按钮重算一遍的成本远低于引入一套响应式。
 */
export function createToolbar(
  sheet: CyExcelCore,
  container: HTMLElement,
  options: ToolbarOptions = {}
): ToolbarHandle {
  const items = sortItems([...(options.schema ?? defaultToolbarSchema), ...(options.items ?? [])]);

  const root = el("div", "cyx-toolbar");
  root.setAttribute("role", "toolbar");
  styles.apply(root, styles.toolbarRoot);
  const moreButton = el("button", "cyx-toolbar-button cyx-toolbar-more");
  moreButton.type = "button";
  moreButton.innerHTML = icons.more;
  styles.apply(moreButton, styles.button);
  moreButton.setAttribute("aria-label", "更多");
  moreButton.hidden = true;
  const overflowHost = el("div", "cyx-toolbar-overflow");
  overflowHost.hidden = true;

  /** 当前打开的弹层（面板或溢出菜单）的清理函数。同一时刻只允许开一个。 */
  let closeOpenLayer: (() => void) | null = null;
  const closeLayer = (): void => {
    closeOpenLayer?.();
    closeOpenLayer = null;
  };

  const context = (): ToolbarContext => ({
    sheet,
    range: sheet.selection.range,
    styleOf: (row, column) => sheet.getCellStyle(row, column)
  });

  const rendered: {
    item: ToolbarItem;
    node: HTMLElement;
    control: HTMLElement;
    repaint?: () => void;
  }[] = [];
  const stateDisposers: (() => void)[] = [];

  function openPanel(item: ToolbarItem, anchor: HTMLElement): void {
    closeLayer();
    const layer = el("div", "cyx-toolbar-panel");
    styles.apply(layer, styles.panel);
    const rect = anchor.getBoundingClientRect();
    const rootRect = root.getBoundingClientRect();
    layer.style.left = `${rect.left - rootRect.left}px`;
    layer.style.top = `${rect.bottom - rootRect.top}px`;
    const disposePanel = item.panel?.(context(), layer) ?? (() => {});
    root.append(layer);

    const onPointerDown = (event: MouseEvent): void => {
      if (!layer.contains(event.target as Node) && !anchor.contains(event.target as Node)) {
        closeLayer();
      }
    };
    const onKeyDown = (event: KeyboardEvent): void => {
      if (event.key === "Escape") {
        closeLayer();
        anchor.focus();
      }
    };
    // 捕获阶段监听：面板里的按钮自己也会 stopPropagation，冒泡阶段可能收不到。
    document.addEventListener("mousedown", onPointerDown, true);
    document.addEventListener("keydown", onKeyDown, true);

    closeOpenLayer = () => {
      document.removeEventListener("mousedown", onPointerDown, true);
      document.removeEventListener("keydown", onKeyDown, true);
      disposePanel();
      layer.remove();
    };
  }

  function buildControl(item: ToolbarItem): HTMLElement {
    if (item.type === "select") {
      const select = el("select", "cyx-toolbar-select");
      styles.apply(select, styles.select);
      select.setAttribute("aria-label", item.label);
      for (const option of item.options ?? []) {
        const node = document.createElement("option");
        node.value = String(option.value);
        node.textContent = option.label;
        select.append(node);
      }
      select.addEventListener("change", () => {
        void item.apply?.(context(), select.value);
      });
      return select;
    }

    const button = el("button", "cyx-toolbar-button");
    button.type = "button";
    styles.apply(button, styles.button);
    button.setAttribute("aria-label", item.label);
    if (item.tooltip !== undefined) {
      button.title = item.tooltip;
    }
    if (item.icon !== undefined) {
      // icons.ts 全是字面量常量，没有任何外部输入参与拼接——这是能用 innerHTML 的前提。
      button.innerHTML = item.icon;
    } else {
      button.textContent = item.label;
    }
    button.addEventListener("click", () => {
      if (item.panel !== undefined) {
        openPanel(item, button);
        return;
      }
      closeLayer();
      void item.run?.(context());
    });
    return button;
  }

  let currentGroup: string | null = null;
  for (const item of items) {
    if (currentGroup !== null && item.group !== currentGroup) {
      const line = el("span", "cyx-toolbar-separator");
      styles.apply(line, styles.separator);
      root.append(line);
    }
    currentGroup = item.group;
    const slot = el("span", "cyx-toolbar-slot");
    styles.apply(slot, styles.slot);
    const control = buildControl(item);
    slot.append(control);
    root.append(slot);
    const entry: (typeof rendered)[number] = { item, node: slot, control };
    if (control instanceof HTMLButtonElement) {
      let activated = false;
      stateDisposers.push(styles.attachButtonStates(control, () => activated));
      entry.repaint = () => {
        activated = entry.node.dataset.activated === "true";
        control.dispatchEvent(new Event("mouseleave"));
      };
    }
    rendered.push(entry);
  }
  root.append(moreButton, overflowHost);
  container.append(root);

  function refresh(): void {
    const ctx = context();
    for (const entry of rendered) {
      const state = resolveItemState(entry.item, ctx);
      entry.node.dataset.activated = state.activated ? "true" : "false";
      if (entry.control instanceof HTMLButtonElement || entry.control instanceof HTMLSelectElement) {
        entry.control.disabled = state.disabled;
      }
      if (entry.control instanceof HTMLSelectElement) {
        // 选区里的值不一致时 value 是 undefined——这时显示空白，而不是随便挑一格
        // 的值糊弄（见 state.ts 的 commonValue）。
        entry.control.value = state.value === undefined ? "" : String(state.value);
      }
      entry.repaint?.();
    }
  }

  function applyOverflow(): void {
    const { visible, overflowed } = computeOverflow({
      itemWidths: rendered.map((entry) => entry.node.offsetWidth),
      available: root.clientWidth,
      moreButtonWidth: moreButton.offsetWidth || 32
    });
    for (const index of visible) {
      rendered[index]?.node.removeAttribute("hidden");
    }
    for (const index of overflowed) {
      rendered[index]?.node.setAttribute("hidden", "");
    }
    moreButton.hidden = overflowed.length === 0;
  }

  /** 一次事务可能连发多条 change，几十个按钮没必要各算几遍——合并到下一帧。 */
  let frame = 0;
  const scheduleRefresh = (): void => {
    if (frame !== 0) {
      return;
    }
    frame = requestAnimationFrame(() => {
      frame = 0;
      refresh();
      applyOverflow();
    });
  };

  const disposers = [
    sheet.events.on("selection-change", scheduleRefresh),
    sheet.events.on("change", scheduleRefresh),
    sheet.events.on("layout-change", scheduleRefresh)
  ];

  moreButton.addEventListener("click", () => {
    closeLayer();
    const layer = el("div", "cyx-toolbar-panel");
    styles.apply(layer, styles.panel);
    layer.style.right = "0px";
    layer.style.top = `${moreButton.offsetHeight}px`;
    for (const entry of rendered) {
      if (!entry.node.hasAttribute("hidden")) {
        continue;
      }
      const clone = el("button", "cyx-toolbar-panel-action");
      clone.type = "button";
      styles.apply(clone, styles.panelAction);
      clone.textContent = entry.item.label;
      clone.addEventListener("click", () => {
        closeLayer();
        void entry.item.run?.(context());
      });
      layer.append(clone);
    }
    root.append(layer);
    closeOpenLayer = () => {
      layer.remove();
    };
  });

  const resizeObserver =
    typeof ResizeObserver === "undefined" ? null : new ResizeObserver(() => applyOverflow());
  resizeObserver?.observe(root);

  refresh();
  applyOverflow();

  return {
    refresh,
    dispose: () => {
      closeLayer();
      if (frame !== 0) {
        cancelAnimationFrame(frame);
        frame = 0;
      }
      resizeObserver?.disconnect();
      for (const dispose of [...disposers, ...stateDisposers]) {
        dispose();
      }
      root.remove();
    }
  };
}
