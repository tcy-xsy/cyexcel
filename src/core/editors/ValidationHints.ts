import type { CellRect } from "../layout";
import type { SceneSnapshot } from "../scene";
import type { CellValue } from "../sheet/types";
import type { CellRange } from "../types";
import { describeCellHints, type DataValidationRule } from "../sheet/validations";
import { visibleRectForRange } from "./visibleCellRect";

export interface ValidationHintsOptions {
  /** 编辑层。提示挂在这里而不是 `document.body`，理由同 `ValidationBubble`。 */
  readonly container: HTMLElement;
  readonly getRule: (row: number, column: number) => DataValidationRule | undefined;
  readonly getCellValue: (row: number, column: number) => CellValue | undefined;
  readonly findMergeAt?: (row: number, column: number) => CellRange | undefined;
  /** 用户在下拉里选了一项。 */
  readonly onPick: (row: number, column: number, value: string) => void;
}

const ARROW_WIDTH = 18;

/**
 * 选中单元格时的两样东西（数据验证设计稿 §8）：`list` 规则的下拉箭头和列表，以及
 * 规则的输入提示。都是 DOM，挂在编辑层里；一次只跟着一个活动格。
 *
 * **为什么全用 `<button>`。** 核心的指针路由把 `input/select/textarea/button/option`
 * 认作"原生输入目标"，不当成单元格点击处理；换成 `div` 的话，点一下下拉项会被路由到
 * 下面的画布上，先选中了别的单元格。
 *
 * 这一层在 node 里测不了（没有 DOM），由 playground 手验；"该显示什么"的决策是纯函数
 * `describeCellHints`，那部分有单测。
 */
export class ValidationHints {
  readonly #options: ValidationHintsOptions;
  #arrow: HTMLButtonElement | undefined;
  #prompt: HTMLDivElement | undefined;
  #list: HTMLDivElement | undefined;
  #target: { readonly row: number; readonly column: number } | undefined;
  #removeOutsideListeners: (() => void) | undefined;

  constructor(options: ValidationHintsOptions) {
    this.#options = options;
  }

  /**
   * 用最新的 Scene 和活动格重新摆放。`active` 为 `null`（正在编辑、没有选区……）时
   * 全部收起。每次 Scene 发布、选区变化、规则变化都会调用，所以要便宜：没有规则就是
   * 一次查询加 `hide()`。
   */
  sync(scene: SceneSnapshot, active: { readonly row: number; readonly column: number } | null): void {
    if (active === null) {
      this.hide();
      return;
    }
    const rule = this.#options.getRule(active.row, active.column);
    const hints = rule === undefined ? undefined : describeCellHints(rule);
    if (hints === undefined) {
      this.hide();
      return;
    }
    const merge = this.#options.findMergeAt?.(active.row, active.column);
    const range: CellRange = merge ?? {
      startRow: active.row,
      endRow: active.row,
      startColumn: active.column,
      endColumn: active.column
    };
    const rect = visibleRectForRange(scene, range);
    if (rect === null) {
      this.hide();
      return;
    }

    // 换了一个格子：上一格打开的列表作废
    if (this.#target?.row !== active.row || this.#target?.column !== active.column) {
      this.#closeList();
    }
    this.#target = { row: active.row, column: active.column };

    if (hints.dropdownItems !== undefined) {
      this.#showArrow(rect, hints.dropdownItems);
    } else {
      this.#removeArrow();
      this.#closeList();
    }

    if (this.#list !== undefined) {
      // 列表开着的时候提示会盖在它上面，先收起
      this.#removePrompt();
      this.#placeList(rect);
    } else if (hints.prompt !== undefined || hints.promptTitle !== undefined) {
      this.#showPrompt(rect, hints.promptTitle, hints.prompt);
    } else {
      this.#removePrompt();
    }
  }

  hide(): void {
    this.#target = undefined;
    this.#removeArrow();
    this.#removePrompt();
    this.#closeList();
  }

  dispose(): void {
    this.hide();
  }

  // -- 箭头 -------------------------------------------------------------

  #showArrow(rect: CellRect, items: readonly string[]): void {
    if (this.#arrow === undefined) {
      const arrow = document.createElement("button");
      arrow.type = "button";
      arrow.textContent = "▾";
      arrow.setAttribute("aria-label", "打开下拉列表");
      arrow.dataset.cyValidation = "dropdown-arrow";
      arrow.style.cssText = [
        "position:absolute",
        "z-index:15",
        "box-sizing:border-box",
        `width:${ARROW_WIDTH}px`,
        "padding:0",
        "border:1px solid #c4c7cc",
        "border-radius:3px",
        "background:#f8f9fa",
        "color:#5f6368",
        "font:12px/1 system-ui,sans-serif",
        "cursor:pointer",
        "pointer-events:auto"
      ].join(";");
      // 点箭头不该把焦点从画布上抢走，否则键盘导航立刻失灵
      arrow.addEventListener("mousedown", (event) => event.preventDefault());
      this.#arrow = arrow;
      this.#options.container.append(arrow);
    }
    const arrow = this.#arrow;
    arrow.style.left = `${rect.x + rect.width - ARROW_WIDTH - 1}px`;
    arrow.style.top = `${rect.y + 1}px`;
    arrow.style.height = `${Math.max(0, rect.height - 2)}px`;
    // 每次都重新绑定：条目和目标格随规则、选区变化
    arrow.onclick = () => {
      if (this.#list !== undefined) {
        this.#closeList();
        return;
      }
      this.#openList(rect, items);
    };
  }

  #removeArrow(): void {
    this.#arrow?.remove();
    this.#arrow = undefined;
  }

  // -- 下拉列表 ---------------------------------------------------------

  #openList(rect: CellRect, items: readonly string[]): void {
    const target = this.#target;
    if (target === undefined) {
      return;
    }
    const current = this.#options.getCellValue(target.row, target.column);
    const currentText = current === undefined ? "" : String(current);

    const list = document.createElement("div");
    list.setAttribute("role", "listbox");
    list.dataset.cyValidation = "dropdown-list";
    list.style.cssText = [
      "position:absolute",
      "z-index:25",
      "box-sizing:border-box",
      "max-height:200px",
      "overflow-y:auto",
      "padding:2px 0",
      "background:#fff",
      "border:1px solid #c4c7cc",
      "border-radius:4px",
      "box-shadow:0 2px 8px rgba(0,0,0,.18)",
      "font:13px/1.4 system-ui,sans-serif",
      "pointer-events:auto"
    ].join(";");
    list.addEventListener("mousedown", (event) => event.preventDefault());

    for (const item of items) {
      const option = document.createElement("button");
      option.type = "button";
      option.setAttribute("role", "option");
      option.textContent = item;
      const selected = item === currentText;
      option.setAttribute("aria-selected", String(selected));
      option.style.cssText = [
        "display:block",
        "width:100%",
        "text-align:left",
        "padding:4px 10px",
        "border:0",
        `background:${selected ? "#e8f0fe" : "transparent"}`,
        "color:#202124",
        "font:inherit",
        "cursor:pointer"
      ].join(";");
      option.addEventListener("mouseenter", () => {
        option.style.background = "#f1f3f4";
      });
      option.addEventListener("mouseleave", () => {
        option.style.background = selected ? "#e8f0fe" : "transparent";
      });
      option.addEventListener("click", () => {
        this.#closeList();
        this.#options.onPick(target.row, target.column, item);
      });
      list.append(option);
    }

    this.#options.container.append(list);
    this.#list = list;
    this.#placeList(rect);
    this.#removePrompt();

    // 点到列表外面或按 Esc 关掉。捕获阶段监听，免得被别处的 stopPropagation 吞掉。
    const onPointerDown = (event: Event): void => {
      const node = event.target as Node | null;
      if (node !== null && (list.contains(node) || this.#arrow?.contains(node))) {
        return;
      }
      this.#closeList();
    };
    const onKeyDown = (event: KeyboardEvent): void => {
      if (event.key === "Escape") {
        this.#closeList();
      }
    };
    document.addEventListener("pointerdown", onPointerDown, true);
    document.addEventListener("keydown", onKeyDown, true);
    this.#removeOutsideListeners = () => {
      document.removeEventListener("pointerdown", onPointerDown, true);
      document.removeEventListener("keydown", onKeyDown, true);
    };
  }

  #placeList(rect: CellRect): void {
    if (this.#list === undefined) {
      return;
    }
    this.#list.style.left = `${rect.x}px`;
    this.#list.style.top = `${rect.y + rect.height}px`;
    this.#list.style.minWidth = `${rect.width}px`;
  }

  #closeList(): void {
    this.#removeOutsideListeners?.();
    this.#removeOutsideListeners = undefined;
    this.#list?.remove();
    this.#list = undefined;
  }

  // -- 输入提示 ---------------------------------------------------------

  #showPrompt(rect: CellRect, title: string | undefined, message: string | undefined): void {
    if (this.#prompt === undefined) {
      const bubble = document.createElement("div");
      bubble.setAttribute("role", "note");
      bubble.dataset.cyValidation = "prompt";
      bubble.style.cssText = [
        "position:absolute",
        "z-index:14",
        "box-sizing:border-box",
        "max-width:240px",
        "padding:6px 8px",
        "background:#fffbe6",
        "border:1px solid #e0c46c",
        "border-radius:4px",
        "box-shadow:0 1px 4px rgba(0,0,0,.15)",
        "font:12px/1.5 system-ui,sans-serif",
        "color:#202124",
        // 只是提示：不接指针事件，不挡住下面的单元格
        "pointer-events:none"
      ].join(";");
      this.#prompt = bubble;
      this.#options.container.append(bubble);
    }
    const bubble = this.#prompt;
    bubble.replaceChildren();
    if (title !== undefined) {
      const heading = document.createElement("div");
      heading.style.fontWeight = "600";
      heading.textContent = title;
      bubble.append(heading);
    }
    if (message !== undefined) {
      const body = document.createElement("div");
      body.textContent = message;
      bubble.append(body);
    }
    bubble.style.left = `${rect.x}px`;
    bubble.style.top = `${rect.y + rect.height + 4}px`;
  }

  #removePrompt(): void {
    this.#prompt?.remove();
    this.#prompt = undefined;
  }
}
