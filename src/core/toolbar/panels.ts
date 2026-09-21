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

/** 十个常用色 + 一个"无"。不做取色器——那是宿主的活，`setRangeStyle` 接受任意 CSS 颜色串。 */
const SWATCHES = [
  "#111827", "#ef4444", "#f97316", "#eab308", "#22c55e",
  "#06b6d4", "#3b82f6", "#8b5cf6", "#ec4899", "#ffffff"
];

export function colorPanel(
  host: HTMLElement,
  options: {
    current: string | undefined;
    onPick: (color: string | null) => void;
    /** 回报给工具栏：分段按钮据此记住"上次选的颜色"，下次点图标直接套用。 */
    onChosen?: (value: string) => void;
  }
): () => void {
  const grid = el("div", "cyx-toolbar-swatches");
  styles.apply(grid, styles.swatchGrid);
  for (const color of SWATCHES) {
    const button = el("button", "cyx-toolbar-swatch");
    button.type = "button";
    styles.apply(button, styles.swatch);
    button.style.background = color;
    button.setAttribute("aria-label", color);
    if (options.current === color) {
      button.dataset.current = "true";
      // 必须连内联样式一起写：这个模块没有 stylesheet（见 styles.ts 顶部），
      // `[data-current]` 选不中任何东西，只设 dataset 的话高亮根本不出现——
      // "当前这片选区是什么颜色"于是永远看不见。dataset 留着给宿主换皮肤用。
      styles.apply(button, styles.swatchCurrent);
    }
    button.addEventListener("click", () => {
      options.onChosen?.(color);
      options.onPick(color);
    });
    grid.append(button);
  }
  const clear = el("button", "cyx-toolbar-panel-action", "无颜色");
  clear.type = "button";
  styles.apply(clear, styles.panelAction);
  // "没设颜色"也是一种当前状态，同样要标出来，否则整片选区没有颜色时调色板上
  // 一个标记都没有，看起来像是坏了。
  if (options.current === undefined) {
    clear.dataset.current = "true";
    styles.apply(clear, styles.panelActionCurrent);
  }
  clear.addEventListener("click", () => options.onPick(null));

  host.append(grid, clear);
  return () => {
    host.replaceChildren();
  };
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

