import type { BorderSide } from "../style/types";
import type { ToolbarContext } from "./types";
import * as styles from "./styles";
import { el } from "./dom";

/**
 * 四个弹出面板，都是纯 DOM。
 *
 * 统一的契约：`(host, …) => disposer`——把内容塞进调用方给的 `host` 里，返回一个
 * 清理函数。弹出层的定位、外部点击关闭、Esc 关闭由 `createToolbar` 统一处理，
 * 面板自己只管内容，所以这些函数彼此之间没有重复的壳。
 *
 * 这一层在 node 里测不了（没有 DOM，项目不引 jsdom——见 step25-toolbar-plan.md §5），
 * 所以刻意写得薄：没有分支逻辑，所有判断都在 `state.ts` 那边。
 */

/** 常用主题色按行分组，保证同一行是同一色相的深浅变化。 */
const THEME_COLORS: readonly { readonly label: string; readonly color: string }[] = Object.freeze([
  { label: "墨黑", color: "#111827" },
  { label: "灰", color: "#4b5563" },
  { label: "浅灰", color: "#9ca3af" },
  { label: "纸白", color: "#f8fafc" },
  { label: "白", color: "#ffffff" },
  { label: "红", color: "#dc2626" },
  { label: "橙", color: "#ea580c" },
  { label: "琥珀", color: "#d97706" },
  { label: "绿", color: "#16a34a" },
  { label: "青", color: "#0891b2" },
  { label: "蓝", color: "#2563eb" },
  { label: "靛蓝", color: "#4f46e5" },
  { label: "紫", color: "#7c3aed" },
  { label: "粉", color: "#db2777" },
  { label: "棕", color: "#92400e" },
  { label: "浅红", color: "#fca5a5" },
  { label: "浅橙", color: "#fdba74" },
  { label: "浅琥珀", color: "#fcd34d" },
  { label: "浅绿", color: "#86efac" },
  { label: "浅蓝", color: "#93c5fd" }
]);

const RECENT_COLOR_LIMIT = 10;
let recentColors: string[] = [];

function rememberColor(color: string): void {
  recentColors = [color, ...recentColors.filter(item => item !== color)]
    .slice(0, RECENT_COLOR_LIMIT);
}

function cssColorToHex(value: string): string | undefined {
  const source = value.trim();
  if (/^#[0-9a-f]{6}$/i.test(source)) return source.toLowerCase();
  if (/^#[0-9a-f]{3}$/i.test(source)) {
    return `#${[...source.slice(1)].map(char => char + char).join("")}`.toLowerCase();
  }

  const probe = document.createElement("span");
  probe.style.position = "absolute";
  probe.style.left = "-9999px";
  probe.style.color = source;
  if (!probe.style.color) return undefined;

  document.body.append(probe);
  const computed = getComputedStyle(probe).color;
  probe.remove();
  const match = /^rgba?\((\d+), (\d+), (\d+)(?:, 1)?\)$/i.exec(computed);
  if (!match) return undefined;

  return `#${match.slice(1, 4)
    .map(channel => Number(channel).toString(16).padStart(2, "0"))
    .join("")}`;
}

function createSwatch(
  color: string,
  label: string,
  current: string | undefined,
  onPick: (color: string) => void
): HTMLButtonElement {
  const button = el("button", "cyx-toolbar-swatch");
  button.type = "button";
  styles.apply(button, styles.swatch);
  button.style.background = color;
  button.title = `${label} ${color}`;
  button.setAttribute("aria-label", `${label} ${color}`);
  if (current !== undefined && cssColorToHex(current) === color) {
    button.dataset.current = "true";
    styles.apply(button, styles.swatchCurrent);
  }
  const paintState = (hovered: boolean): void => {
    if (button.dataset.current === "true") {
      button.style.boxShadow = "";
      styles.apply(button, styles.swatchCurrent);
      return;
    }
    if (!hovered) button.style.outline = "";
    button.style.boxShadow = hovered ? `0 0 0 2px ${styles.COLORS.hover}` : "";
  };
  button.addEventListener("mouseenter", () => paintState(true));
  button.addEventListener("mouseleave", () => paintState(false));
  button.addEventListener("focus", () => paintState(true));
  button.addEventListener("blur", () => paintState(false));
  button.addEventListener("click", () => onPick(color));
  return button;
}

function closePanel(host: HTMLElement): void {
  window.setTimeout(() => {
    host.dispatchEvent(new CustomEvent("cyx-toolbar-close"));
  }, 0);
}

export function colorPanel(
  host: HTMLElement,
  options: {
    current: string | undefined;
    onPick: (color: string | null) => void;
    /** 回报给工具栏：分段按钮据此记住"上次选的颜色"，下次点图标直接套用。 */
    onChosen?: (value: string) => void;
  }
): () => void {
  const current = options.current === undefined ? undefined : cssColorToHex(options.current);
  const pick = (color: string): void => {
    rememberColor(color);
    options.onChosen?.(color);
    options.onPick(color);
    closePanel(host);
  };

  const root = el("div", "cyx-toolbar-color-panel");
  styles.apply(root, styles.colorPanel);
  const sectionLabel = (text: string): HTMLElement => {
    const label = el("div", "cyx-toolbar-panel-label", text);
    styles.apply(label, styles.panelSectionLabel);
    return label;
  };

  if (recentColors.length > 0) {
    root.append(sectionLabel("最近使用"));
    const recentGrid = el("div", "cyx-toolbar-swatches");
    styles.apply(recentGrid, styles.swatchGrid);
    for (const color of recentColors) {
      recentGrid.append(createSwatch(color, "最近颜色", current, pick));
    }
    root.append(recentGrid);
  }

  root.append(sectionLabel("常用颜色"));
  const grid = el("div", "cyx-toolbar-swatches");
  styles.apply(grid, styles.swatchGrid);
  for (const themeColor of THEME_COLORS) {
    grid.append(createSwatch(themeColor.color, themeColor.label, current, pick));
  }

  root.append(sectionLabel("自定义"), createCustomColor(current, pick));

  const clear = el("button", "cyx-toolbar-panel-action", "无颜色");
  clear.type = "button";
  styles.apply(clear, styles.panelAction);
  // "没设颜色"也是一种当前状态，同样要标出来，否则整片选区没有颜色时调色板上
  // 一个标记都没有，看起来像是坏了。
  if (options.current === undefined) {
    clear.dataset.current = "true";
    styles.apply(clear, styles.panelActionCurrent);
  }
  clear.addEventListener("click", () => {
    options.onPick(null);
    closePanel(host);
  });

  root.append(clear);
  host.append(root);
  return () => {
    host.replaceChildren();
  };
}

function createCustomColor(current: string | undefined, onPick: (color: string) => void): HTMLElement {
  const row = el("div", "cyx-toolbar-custom-color");
  styles.apply(row, styles.customColorRow);

  const picker = document.createElement("input");
  picker.type = "color";
  picker.value = current ?? "#111827";
  picker.setAttribute("aria-label", "打开取色器");
  styles.apply(picker, styles.customColorPicker);
  picker.addEventListener("change", () => onPick(picker.value));

  const input = el("input", "cyx-toolbar-panel-input");
  styles.apply(input, styles.customColorInput);
  input.type = "text";
  input.spellcheck = false;
  input.placeholder = "#1D4ED8";
  input.setAttribute("aria-label", "十六进制颜色");

  const apply = el("button", "cyx-toolbar-panel-action", "应用");
  apply.type = "button";
  styles.apply(apply, styles.customColorApply);
  const submit = (): void => {
    const color = cssColorToHex(input.value);
    if (color === undefined) {
      input.setAttribute("aria-invalid", "true");
      input.style.borderColor = "#dc2626";
      input.focus();
      return;
    }
    input.removeAttribute("aria-invalid");
    input.style.borderColor = "";
    onPick(color);
  };
  apply.addEventListener("click", submit);
  input.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      submit();
    }
  });

  row.append(picker, input, apply);
  return row;
}

const BORDER_SIDE: BorderSide = { color: "#111827", width: 1 };

/** 边框预设：key 同时是分段按钮记住的那个值（见 applyBorderPreset）。 */
export const BORDER_PRESETS: readonly { key: string; label: string }[] = Object.freeze([
  { key: "all", label: "四周" },
  { key: "top", label: "上边框" },
  { key: "bottom", label: "下边框" },
  { key: "left", label: "左边框" },
  { key: "right", label: "右边框" },
  { key: "none", label: "无边框" }
]);

/** 按预设 key 套用边框。分段按钮的主区与面板走的是同一条路径。 */
export function applyBorderPreset(ctx: ToolbarContext, key: string): void {
  const patch =
    key === "all"
      ? { border: { top: BORDER_SIDE, right: BORDER_SIDE, bottom: BORDER_SIDE, left: BORDER_SIDE } }
      : key === "none"
        ? { border: null }
        : { border: { [key]: BORDER_SIDE } };
  ctx.sheet.setRangeStyle(ctx.range, patch);
}

export function borderPanel(
  host: HTMLElement,
  ctx: ToolbarContext,
  onChosen?: (value: string) => void
): () => void {
  const list = el("div", "cyx-toolbar-panel-list");
  for (const preset of BORDER_PRESETS) {
    const button = el("button", "cyx-toolbar-panel-action", preset.label);
    button.type = "button";
    styles.apply(button, styles.panelAction);
    button.addEventListener("click", () => {
      onChosen?.(preset.key);
      applyBorderPreset(ctx, preset.key);
    });
    list.append(button);
  }
  host.append(list);
  return () => {
    host.replaceChildren();
  };
}

/**
 * 过滤面板。
 *
 * **谓词由这里（宿主侧的工具栏）构造，core 不猜**——这与 Step 22 Part B 定下的
 * "`toggle-filter` 永远是宿主驱动的"是同一条：core 不知道该用什么条件，只提供
 * `filterByColumn(column, predicate)`。这里给的是三个最常用的条件形态。
 */
export function filterPanel(host: HTMLElement, ctx: ToolbarContext): () => void {
  const column = ctx.range.startColumn;
  const form = el("div", "cyx-toolbar-panel-list");

  const input = el("input", "cyx-toolbar-panel-input");
  styles.apply(input, styles.panelInput);
  input.type = "text";
  input.placeholder = "包含文本…";
  input.setAttribute("aria-label", "过滤条件");

  const apply = el("button", "cyx-toolbar-panel-action", "应用");
  apply.type = "button";
  styles.apply(apply, styles.panelAction);
  apply.addEventListener("click", () => {
    const needle = input.value.trim().toLowerCase();
    ctx.sheet.filterByColumn(column, (value) =>
      needle === "" ? true : String(value ?? "").toLowerCase().includes(needle)
    );
  });

  const nonEmpty = el("button", "cyx-toolbar-panel-action", "只看非空");
  nonEmpty.type = "button";
  styles.apply(nonEmpty, styles.panelAction);
  nonEmpty.addEventListener("click", () => {
    ctx.sheet.filterByColumn(column, (value) => value !== null && value !== undefined && value !== "");
  });

  const clear = el("button", "cyx-toolbar-panel-action", "清除过滤");
  clear.type = "button";
  styles.apply(clear, styles.panelAction);
  clear.addEventListener("click", () => ctx.sheet.clearFilter());

  form.append(input, apply, nonEmpty, clear);
  host.append(form);
  return () => {
    host.replaceChildren();
  };
}
