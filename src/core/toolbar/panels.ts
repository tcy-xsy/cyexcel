import type { BorderSide } from "../style/types";
import type { ToolbarContext } from "./types";

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

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
  text?: string
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className !== undefined) {
    node.className = className;
  }
  if (text !== undefined) {
    node.textContent = text;
  }
  return node;
}

/** 十个常用色 + 一个"无"。不做取色器——那是宿主的活，`setRangeStyle` 接受任意 CSS 颜色串。 */
const SWATCHES = [
  "#111827", "#ef4444", "#f97316", "#eab308", "#22c55e",
  "#06b6d4", "#3b82f6", "#8b5cf6", "#ec4899", "#ffffff"
];

export function colorPanel(
  host: HTMLElement,
  options: { current: string | undefined; onPick: (color: string | null) => void }
): () => void {
  const grid = el("div", "cyx-toolbar-swatches");
  for (const color of SWATCHES) {
    const button = el("button", "cyx-toolbar-swatch");
    button.type = "button";
    button.style.background = color;
    button.setAttribute("aria-label", color);
    if (options.current === color) {
      button.dataset.current = "true";
    }
    button.addEventListener("click", () => options.onPick(color));
    grid.append(button);
  }
  const clear = el("button", "cyx-toolbar-panel-action", "无颜色");
  clear.type = "button";
  clear.addEventListener("click", () => options.onPick(null));

  host.append(grid, clear);
  return () => {
    host.replaceChildren();
  };
}

const BORDER_SIDE: BorderSide = { color: "#111827", width: 1 };

export function borderPanel(host: HTMLElement, ctx: ToolbarContext): () => void {
  const presets: readonly { label: string; patch: Record<string, unknown> }[] = [
    { label: "四周", patch: { border: { top: BORDER_SIDE, right: BORDER_SIDE, bottom: BORDER_SIDE, left: BORDER_SIDE } } },
    { label: "上", patch: { border: { top: BORDER_SIDE } } },
    { label: "下", patch: { border: { bottom: BORDER_SIDE } } },
    { label: "左", patch: { border: { left: BORDER_SIDE } } },
    { label: "右", patch: { border: { right: BORDER_SIDE } } },
    { label: "无", patch: { border: null } }
  ];
  const list = el("div", "cyx-toolbar-panel-list");
  for (const preset of presets) {
    const button = el("button", "cyx-toolbar-panel-action", preset.label);
    button.type = "button";
    button.addEventListener("click", () => {
      ctx.sheet.setRangeStyle(ctx.range, preset.patch);
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
  input.type = "text";
  input.placeholder = "包含文本…";
  input.setAttribute("aria-label", "过滤条件");

  const apply = el("button", "cyx-toolbar-panel-action", "应用");
  apply.type = "button";
  apply.addEventListener("click", () => {
    const needle = input.value.trim().toLowerCase();
    ctx.sheet.filterByColumn(column, (value) =>
      needle === "" ? true : String(value ?? "").toLowerCase().includes(needle)
    );
  });

  const nonEmpty = el("button", "cyx-toolbar-panel-action", "只看非空");
  nonEmpty.type = "button";
  nonEmpty.addEventListener("click", () => {
    ctx.sheet.filterByColumn(column, (value) => value !== null && value !== undefined && value !== "");
  });

  const clear = el("button", "cyx-toolbar-panel-action", "清除过滤");
  clear.type = "button";
  clear.addEventListener("click", () => ctx.sheet.clearFilter());

  form.append(input, apply, nonEmpty, clear);
  host.append(form);
  return () => {
    host.replaceChildren();
  };
}

export function findPanel(host: HTMLElement, ctx: ToolbarContext): () => void {
  const form = el("div", "cyx-toolbar-panel-list");

  const query = el("input", "cyx-toolbar-panel-input");
  query.type = "text";
  query.placeholder = "查找…";
  query.setAttribute("aria-label", "查找内容");

  const replacement = el("input", "cyx-toolbar-panel-input");
  replacement.type = "text";
  replacement.placeholder = "替换为…";
  replacement.setAttribute("aria-label", "替换为");

  const status = el("div", "cyx-toolbar-panel-status", "");

  const find = el("button", "cyx-toolbar-panel-action", "查找全部");
  find.type = "button";
  find.addEventListener("click", () => {
    void ctx.sheet.find(query.value).then((matches) => {
      status.textContent =
        matches === null ? "已取消" : `${matches.length} 个匹配`;
    });
  });

  const next = el("button", "cyx-toolbar-panel-action", "下一个");
  next.type = "button";
  next.addEventListener("click", () => {
    const match = ctx.sheet.findNext();
    status.textContent = match === null ? "没有匹配" : `第 ${match.row} 行第 ${match.column} 列`;
  });

  const previous = el("button", "cyx-toolbar-panel-action", "上一个");
  previous.type = "button";
  previous.addEventListener("click", () => {
    const match = ctx.sheet.findPrevious();
    status.textContent = match === null ? "没有匹配" : `第 ${match.row} 行第 ${match.column} 列`;
  });

  const replaceAll = el("button", "cyx-toolbar-panel-action", "全部替换");
  replaceAll.type = "button";
  replaceAll.addEventListener("click", () => {
    void ctx.sheet.replaceAll(query.value, replacement.value).then((count) => {
      status.textContent = count === null ? "已取消" : `替换了 ${count} 处`;
    });
  });

  form.append(query, replacement, find, previous, next, replaceAll, status);
  host.append(form);
  return () => {
    host.replaceChildren();
  };
}
