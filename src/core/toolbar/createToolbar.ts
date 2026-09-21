import type { CyExcelCore } from "../CyExcelCore";
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
/** 作为 `CyExcelCoreOptions.topWidgets` 的一项使用。 */
export function toolbarWidget(options: ToolbarOptions = {}) {
  return (sheet: CyExcelCore, host: HTMLElement) => createToolbar(sheet, host, options);
}

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
  const disposeMoreFocus = styles.attachFocusStyle(moreButton);
  const disposeMoreTooltip = styles.attachTooltip(moreButton, "更多工具");
  styles.setHidden(moreButton, true);
  const overflowHost = el("div", "cyx-toolbar-overflow");
  styles.setHidden(overflowHost, true);

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

  /**
   * 分段按钮"上次选过的值"。点主区套用它，不必每次都展开面板重选——Excel 与
   * Univer 的颜色/边框按钮都是这个行为（Univer 的边框图标甚至会变成上次那一种）。
   */
  const remembered = new Map<string, string>();

  const rendered: {
    item: ToolbarItem;
    node: HTMLElement;
    control: HTMLElement;
    repaint?: () => void;
  }[] = [];
  const stateDisposers: (() => void)[] = [];

  /**
   * 开一个弹层（面板或溢出菜单）。
   *
   * 定位、外部点击关闭、Esc 关闭对两者是同一件事，所以收在这里，面板自己只管内容。
   *
   * **`anchorRect` 是传进来的、不是现场量的**：从溢出菜单里打开一个面板时，作为锚点
   * 的那一行按钮会随着溢出菜单一起被移除，等到这时再 `getBoundingClientRect()` 只能
   * 得到一个全零的矩形，面板会飞到左上角去。所以调用方在关掉旧层之前先量好。
   */
  function openLayer(options: {
    anchorRect: DOMRect;
    /** Esc 关闭后把焦点还给谁。锚点可能已经不在文档里了，所以单独给。 */
    focusTarget: HTMLElement;
    fill: (layer: HTMLElement) => () => void;
  }): void {
    closeLayer();
    const layer = el("div", "cyx-toolbar-panel");
    styles.apply(layer, styles.panel);
    // 视口坐标 + 挂 body：工具栏是 overflow:hidden，挂进去会被整个裁掉（见 styles.panel）。
    layer.style.left = `${options.anchorRect.left}px`;
    layer.style.top = `${options.anchorRect.bottom}px`;
    const disposeContent = options.fill(layer);
    document.body.append(layer);
    // 量一次再往回收：靠右的按钮（"更多"就在最右）弹出的面板会越过视口边缘；
    // 项多的面板（溢出菜单）也可能顶破视口底边。与 tooltip 同一套收边做法。
    const layerRect = layer.getBoundingClientRect();
    if (layerRect.right > window.innerWidth - 4) {
      layer.style.left = `${Math.max(4, window.innerWidth - 4 - layerRect.width)}px`;
    }
    if (layerRect.bottom > window.innerHeight - 4) {
      // 优先翻到锚点上方；上方也放不下就贴着视口底边。
      const above = options.anchorRect.top - layerRect.height;
      layer.style.top = `${above >= 4 ? above : Math.max(4, window.innerHeight - 4 - layerRect.height)}px`;
    }

    const onPointerDown = (event: MouseEvent): void => {
      if (!layer.contains(event.target as Node)) {
        closeLayer();
      }
    };
    const onKeyDown = (event: KeyboardEvent): void => {
      if (event.key === "Escape") {
        closeLayer();
        options.focusTarget.focus();
      }
    };
    // 捕获阶段监听：面板里的按钮自己也会 stopPropagation，冒泡阶段可能收不到。
    document.addEventListener("mousedown", onPointerDown, true);
    document.addEventListener("keydown", onKeyDown, true);

    closeOpenLayer = () => {
      document.removeEventListener("mousedown", onPointerDown, true);
      document.removeEventListener("keydown", onKeyDown, true);
      disposeContent();
      layer.remove();
    };
  }

  function openPanel(item: ToolbarItem, anchor: HTMLElement, onChosen?: (value: string) => void): void {
    // 先量后关：见 openLayer 的 anchorRect 说明。
    const anchorRect = anchor.getBoundingClientRect();
    // 锚点若本身在某个弹层里（从溢出菜单点进来的），它马上就会随那个层消失，
    // 焦点得还给"更多"按钮，否则 Esc 之后焦点掉到 body 上。
    const focusTarget = anchor.closest(".cyx-toolbar-panel") === null ? anchor : moreButton;
    openLayer({
      anchorRect,
      focusTarget,
      fill: (layer) => item.panel?.(context(), layer, onChosen) ?? (() => {})
    });
  }

  /** 一颗图标按钮的公共部分：外观、无障碍名、tooltip。 */
  function makeButton(item: ToolbarItem, label: string, icon: string | undefined): HTMLButtonElement {
    const button = el("button", "cyx-toolbar-button");
    button.type = "button";
    styles.apply(button, styles.button);
    button.setAttribute("aria-label", label);
    if (icon !== undefined) {
      // icons.ts 全是字面量常量，没有任何外部输入参与拼接——这是能用 innerHTML 的前提。
      button.innerHTML = icon;
    } else {
      button.textContent = label;
    }
    return button;
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

    // 分段按钮：同时有 panel 和 apply 的项（颜色、边框）。左边图标直接套用上次选过的
    // 值，右边箭头展开面板——最常见的操作是"再来一次刚才那个"。
    if (item.panel !== undefined && item.apply !== undefined) {
      const group = el("span", "cyx-toolbar-split");
      styles.apply(group, styles.split);

      const main = makeButton(item, item.label, item.icon);
      styles.apply(main, styles.splitMain);
      const swatch = el("span", "cyx-toolbar-split-swatch");
      styles.apply(swatch, styles.splitSwatch);
      main.append(swatch);
      const paintSwatch = (): void => {
        // 只有颜色类的记忆值是个颜色串；边框记的是 "all"/"top" 这类预设名，
        // 拿去当背景色只会得到透明，所以那种项干脆不画色带。
        const value = item.type === "color" ? remembered.get(item.id) ?? item.defaultValue : undefined;
        swatch.style.background = value ?? "transparent";
        styles.setHidden(swatch, value === undefined);
      };
      paintSwatch();
      main.addEventListener("click", () => {
        closeLayer();
        const value = remembered.get(item.id) ?? item.defaultValue;
        if (value !== undefined) {
          void item.apply?.(context(), value);
        }
      });

      const caret = makeButton(item, `${item.label}（更多）`, icons.caretDown);
      styles.apply(caret, styles.splitCaret);
      caret.addEventListener("click", () => {
        openPanel(item, group, (value) => {
          remembered.set(item.id, value);
          paintSwatch();
        });
      });

      group.append(main, caret);
      stateDisposers.push(
        styles.attachButtonStates(main, () => false),
        styles.attachButtonStates(caret, () => false),
        styles.attachFocusStyle(main),
        styles.attachFocusStyle(caret),
        styles.attachTooltip(main, item.tooltip ?? item.label),
        styles.attachTooltip(caret, `${item.label}（更多选项）`)
      );
      return group;
    }

    const button = makeButton(item, item.label, item.icon);
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

  /** 分隔线与"它前面那一项的下标"——那一项被收起时分隔线也要跟着消失。 */
  const separators: { node: HTMLElement; precedingIndex: number }[] = [];

  let currentGroup: string | null = null;
  for (const item of items) {
    if (currentGroup !== null && item.group !== currentGroup) {
      const line = el("span", "cyx-toolbar-separator");
      styles.apply(line, styles.separator);
      root.append(line);
      separators.push({ node: line, precedingIndex: rendered.length - 1 });
    }
    currentGroup = item.group;
    const slot = el("span", "cyx-toolbar-slot");
    styles.apply(slot, styles.slot);
    const control = buildControl(item);
    slot.append(control);
    root.append(slot);
    const entry: (typeof rendered)[number] = { item, node: slot, control };
    if (control instanceof HTMLButtonElement || control instanceof HTMLSelectElement) {
      stateDisposers.push(
        styles.attachFocusStyle(control),
        styles.attachTooltip(control, item.tooltip ?? item.label)
      );
    }
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

  /**
   * 哪些项收进"更多"。
   *
   * **用浏览器的测量，不自己做算术**：早先的版本把各项 offsetWidth 加起来跟容器宽
   * 比，结果漏掉了 flex 的 gap（2px × 三十几项）和分隔线的 margin，总宽被低估上百
   * 像素，临界区就判断成"放得下"——工具栏被裁掉一截却不出现"更多"。往里补 gap、
   * margin、padding 是补不完的，而 `scrollWidth > clientWidth` 天然把它们都算进去了。
   *
   * 代价是逐个隐藏时会多次触发 reflow。项只有三十几个，且这只在容器宽度真的变了
   * 时才跑（见 lastWidth），不在滚动或数据变化的路径上。
   */
  let lastWidth = -1;

  function applyOverflow(force = false): void {
    const available = root.clientWidth;
    if (!force && available === lastWidth) {
      return;
    }
    lastWidth = available;

    for (const entry of rendered) {
      styles.setHidden(entry.node, false);
    }
    syncSeparators();
    styles.setHidden(moreButton, true);
    if (available === 0 || root.scrollWidth <= available) {
      return;
    }

    // 放不下：先把"更多"亮出来（它自己也占位），再从右往左收，收到塞得下为止。
    styles.setHidden(moreButton, false);
    for (let index = rendered.length - 1; index >= 0; index -= 1) {
      const node = rendered[index]?.node;
      if (node !== undefined) {
        styles.setHidden(node, true);
      }
      syncSeparators();
      if (root.scrollWidth <= available) {
        return;
      }
    }
  }

  /** 一整组被收起后，它前面的分隔线也要消失，否则末尾会悬着一条竖线。 */
  function syncSeparators(): void {
    for (const separator of separators) {
      const preceding = rendered[separator.precedingIndex];
      // `HTMLElement.hidden` 的类型是 `boolean | "until-found"`，这里只会是前者。
      styles.setHidden(separator.node, preceding === undefined || preceding.node.hidden === true);
    }
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
      applyOverflow(true);
    });
  };

  const disposers = [
    sheet.events.on("selection-change", scheduleRefresh),
    sheet.events.on("change", scheduleRefresh),
    sheet.events.on("layout-change", scheduleRefresh)
  ];

  /**
   * 溢出菜单：被收起来的项在这里仍然**完整可用**。
   *
   * 早先这里只把每一项克隆成一个按钮、点了调 `run`，于是面板类（颜色、边框、过滤、
   * 查找替换）和下拉类（字体、字号、数字格式）一旦被收起来就点不开了——窄屏下正好
   * 是最需要它们的时候。现在按项的类型分三种渲染：
   *
   * - `select`：直接在菜单里放一个真的下拉，改完就生效，菜单不关（可以连着改字体和字号）；
   * - 有 `panel` 的：点击把菜单换成那个面板，锚点用这一行自己的位置；
   * - 其余：点击执行 `run` 并关闭菜单。
   *
   * 禁用状态照样生效——否则单格选区下"合并"在菜单里还是可点的，点了会抛。
   */
  moreButton.addEventListener("click", () => {
    const anchorRect = moreButton.getBoundingClientRect();
    openLayer({
      anchorRect,
      focusTarget: moreButton,
      fill: (layer) => {
        const ctx = context();
        const disposers: (() => void)[] = [];
        for (const entry of rendered) {
          if (!entry.node.hidden) {
            continue;
          }
          const item = entry.item;
          const state = resolveItemState(item, ctx);

          if (item.type === "select") {
            const row = el("label", "cyx-toolbar-panel-row");
            styles.apply(row, styles.panelRow);
            row.append(document.createTextNode(item.label));
            const select = el("select", "cyx-toolbar-select");
            styles.apply(select, styles.select);
            select.setAttribute("aria-label", item.label);
            for (const option of item.options ?? []) {
              const node = document.createElement("option");
              node.value = String(option.value);
              node.textContent = option.label;
              select.append(node);
            }
            select.value = state.value === undefined ? "" : String(state.value);
            select.disabled = state.disabled;
            const onChange = (): void => {
              void item.apply?.(context(), select.value);
            };
            select.addEventListener("change", onChange);
            disposers.push(() => select.removeEventListener("change", onChange));
            row.append(select);
            layer.append(row);
            continue;
          }

          const button = el("button", "cyx-toolbar-panel-action");
          button.type = "button";
          styles.apply(button, styles.panelAction);
          button.textContent = item.label;
          button.disabled = state.disabled;
          if (state.disabled) {
            button.style.opacity = "0.4";
            button.style.cursor = "default";
          }
          if (state.activated) {
            button.style.color = styles.COLORS.activeText;
          }
          button.addEventListener("click", () => {
            if (item.panel !== undefined) {
              // 先量后关：这一行马上就要随菜单一起消失了。
              openPanel(item, button);
              return;
            }
            closeLayer();
            void item.run?.(context());
          });
          layer.append(button);
        }
        return () => {
          for (const dispose of disposers) {
            dispose();
          }
        };
      }
    });
  });

  const resizeObserver =
    typeof ResizeObserver === "undefined" ? null : new ResizeObserver(() => applyOverflow());
  resizeObserver?.observe(root);

  refresh();
  applyOverflow(true);

  return {
    refresh,
    dispose: () => {
      closeLayer();
      if (frame !== 0) {
        cancelAnimationFrame(frame);
        frame = 0;
      }
      resizeObserver?.disconnect();
      disposeMoreFocus();
      disposeMoreTooltip();
      for (const dispose of [...disposers, ...stateDisposers]) {
        dispose();
      }
      root.remove();
    }
  };
}
