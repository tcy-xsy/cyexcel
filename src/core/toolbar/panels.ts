import { defaultTranslator, type Translator } from "../locale/index.js";
import type { BorderSide } from "../style/types.js";
import type { ToolbarContext } from "./types.js";
import * as styles from "./styles.js";
import { el } from "./dom.js";

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
function themeColors(t: Translator) {
  return Object.freeze([
    { label: t("Ink"), color: "#111827" },
    { label: t("Gray"), color: "#4b5563" },
    { label: t("Light gray"), color: "#9ca3af" },
    { label: t("Off-white"), color: "#f8fafc" },
    { label: t("White"), color: "#ffffff" },
    { label: t("Red"), color: "#dc2626" },
    { label: t("Orange"), color: "#ea580c" },
    { label: t("Amber"), color: "#d97706" },
    { label: t("Green"), color: "#16a34a" },
    { label: t("Cyan"), color: "#0891b2" },
    { label: t("Blue"), color: "#2563eb" },
    { label: t("Indigo"), color: "#4f46e5" },
    { label: t("Purple"), color: "#7c3aed" },
    { label: t("Pink"), color: "#db2777" },
    { label: t("Brown"), color: "#92400e" },
    { label: t("Light red"), color: "#fca5a5" },
    { label: t("Light orange"), color: "#fdba74" },
    { label: t("Light amber"), color: "#fcd34d" },
    { label: t("Light green"), color: "#86efac" },
    { label: t("Light blue"), color: "#93c5fd" }
  ]);

}

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
    t?: Translator;
    current: string | undefined;
    onPick: (color: string | null) => void;
    /** 回报给工具栏：分段按钮据此记住"上次选的颜色"，下次点图标直接套用。 */
    onChosen?: (value: string) => void;
  }
): () => void {
  const t = options.t ?? defaultTranslator;
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
    root.append(sectionLabel(t("Recently used")));
    const recentGrid = el("div", "cyx-toolbar-swatches");
    styles.apply(recentGrid, styles.swatchGrid);
    for (const color of recentColors) {
      recentGrid.append(createSwatch(color, t("Recent color"), current, pick));
    }
    root.append(recentGrid);
  }

  root.append(sectionLabel(t("Theme colors")));
  const grid = el("div", "cyx-toolbar-swatches");
  styles.apply(grid, styles.swatchGrid);
  for (const themeColor of themeColors(t)) {
    grid.append(createSwatch(themeColor.color, themeColor.label, current, pick));
  }
  root.append(grid);

  root.append(sectionLabel(t("Custom")), createCustomColor(current, pick, t));

  const clear = el("button", "cyx-toolbar-panel-action", t("No color"));
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

function createCustomColor(current: string | undefined, onPick: (color: string) => void, t: Translator): HTMLElement {
  const row = el("div", "cyx-toolbar-custom-color");
  styles.apply(row, styles.customColorRow);

  const picker = document.createElement("input");
  picker.type = "color";
  picker.value = current ?? "#111827";
  picker.setAttribute("aria-label", t("Open color picker"));
  styles.apply(picker, styles.customColorPicker);
  picker.addEventListener("change", () => onPick(picker.value));

  const input = el("input", "cyx-toolbar-panel-input");
  styles.apply(input, styles.customColorInput);
  input.type = "text";
  input.spellcheck = false;
  input.placeholder = "#1D4ED8";
  input.setAttribute("aria-label", t("Hex color"));

  const apply = el("button", "cyx-toolbar-panel-action", t("Apply"));
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
function borderPresets(t: Translator): readonly { key: string; label: string }[] {
  return Object.freeze([
    { key: "all", label: t("All borders") },
    { key: "top", label: t("Top border") },
    { key: "bottom", label: t("Bottom border") },
    { key: "left", label: t("Left border") },
    { key: "right", label: t("Right border") },
    { key: "none", label: t("No borders") }
  ]);

}

export const BORDER_PRESETS = borderPresets(defaultTranslator);

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
  const t = ctx.sheet.t;
  const list = el("div", "cyx-toolbar-panel-list");
  for (const preset of borderPresets(t)) {
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
  const t = ctx.sheet.t;
  const column = ctx.range.startColumn;
  const form = el("div", "cyx-toolbar-panel-list");

  const input = el("input", "cyx-toolbar-panel-input");
  styles.apply(input, styles.panelInput);
  input.type = "text";
  input.placeholder = t("Contains text...");
  input.setAttribute("aria-label", t("Filter condition"));

  const apply = el("button", "cyx-toolbar-panel-action", t("Apply"));
  apply.type = "button";
  styles.apply(apply, styles.panelAction);
  apply.addEventListener("click", () => {
    const needle = input.value.trim().toLowerCase();
    ctx.sheet.filterByColumn(column, (value) =>
      needle === "" ? true : String(value ?? "").toLowerCase().includes(needle)
    );
  });

  const nonEmpty = el("button", "cyx-toolbar-panel-action", t("Non-empty only"));
  nonEmpty.type = "button";
  styles.apply(nonEmpty, styles.panelAction);
  nonEmpty.addEventListener("click", () => {
    ctx.sheet.filterByColumn(column, (value) => value !== null && value !== undefined && value !== "");
  });

  const clear = el("button", "cyx-toolbar-panel-action", t("Clear filter condition"));
  clear.type = "button";
  styles.apply(clear, styles.panelAction);
  clear.addEventListener("click", () => ctx.sheet.clearFilter());

  form.append(input, apply, nonEmpty, clear);
  host.append(form);
  return () => {
    host.replaceChildren();
  };
}
