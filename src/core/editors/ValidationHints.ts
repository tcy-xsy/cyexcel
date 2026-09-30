import type { CellRect } from "../layout";
import type { SceneSnapshot } from "../scene";
import type { CellValue } from "../sheet/types";
import type { CellRange } from "../types";
import { describeDropdownItems, type DataValidationRule } from "../sheet/validations";
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

/** 列表与单元格之间的间距，同 Univer 的 `offset: [0, 3]`。 */
const LIST_GAP = 3;

interface CellRef {
  readonly row: number;
  readonly column: number;
}

/**
 * 数据验证里两样挂在编辑层的 DOM（设计稿 §8）：`list` 规则的下拉列表，以及鼠标停在
 * 无效单元格上时的错误提示。一次各至多一个。
 *
 * **箭头不在这里。** 下拉箭头画在画布上（`ContentLayerRenderer` 的 `getCellDropdown`），
 * 所有 list 单元格常驻，点击走画布的命中测试（`CyExcelCore#routeBodyDropdownArrow`）
 * 再调 `toggleList`——这是 Univer 的做法：箭头是绘制的一部分，列表才是弹出层。
 *
 * **列表的生命周期同 Univer 的 `DropdownManagerService`：** 同一时刻至多一个；点列表外面
 * 关闭；活动格一换就关闭（Univer 是选区移到没有下拉的格子时关，这里更严一点，换到另一个
 * list 格也关，免得列表悬在上一格旁边）；选一项后关闭并写入。
 *
 * **错误提示同 Univer 的 `DataValidationAlertController`：** 鼠标停在违规的格子上，
 * 在格子旁弹一个宽 200 的小提示（标题「错误」加规则的错误信息）；移开、滚动、选区变化、
 * 开始编辑都收起。什么时候弹由调用方决定（去抖、判断违规），这里只管画。
 *
 * **为什么列表项用 `<button>`。** 核心的指针路由把 `input/select/textarea/button/option`
 * 认作"原生输入目标"，不当成单元格点击处理；换成 `div` 的话，点一下列表项会被路由到
 * 下面的画布上，先选中了别的单元格。
 *
 * 这一层在 node 里测不了（没有 DOM），由 playground 手验；"该显示什么"的决策是纯函数
 * （`describeDropdownItems`、`outcomeFromResult`），那部分有单测。
 */
export class ValidationHints {
  readonly #options: ValidationHintsOptions;
  #alert: HTMLDivElement | undefined;
  #list: HTMLDivElement | undefined;
  /** 打开着的列表属于哪一格。 */
  #listTarget: CellRef | undefined;
  #removeOutsideListeners: (() => void) | undefined;
  /** 最近一次被"点外面"关掉的列表：哪个指针事件、哪一格。`toggleList` 靠它认出"这一下点的就是箭头"。 */
  #closedByPointer: { readonly event: Event; readonly cell: CellRef } | undefined;

  constructor(options: ValidationHintsOptions) {
    this.#options = options;
  }

  /** 列表当前是不是开着（供调用方判断"再点一次是关"）。 */
  get isListOpen(): boolean {
    return this.#list !== undefined;
  }

  /**
   * 用最新的 Scene 和活动格重新摆放列表。`active` 为 `null`（正在编辑、没有选区……）时
   * 全部收起。每次 Scene 发布、选区变化、规则变化都会调用，所以要便宜。鼠标悬停的错误
   * 提示在这里一并收起：滚动或选区一变，提示指着的位置就不对了。
   */
  sync(scene: SceneSnapshot, active: CellRef | null): void {
    this.hideAlert();
    if (active === null) {
      this.hide();
      return;
    }
    // 活动格换了：上一格打开的列表作废
    if (this.#listTarget !== undefined && (this.#listTarget.row !== active.row || this.#listTarget.column !== active.column)) {
      this.#closeList();
    }
    if (this.#list === undefined) {
      return;
    }
    const rule = this.#options.getRule(active.row, active.column);
    const items = rule === undefined ? undefined : describeDropdownItems(rule);
    const rect = this.#rectOf(scene, active);
    if (items === undefined || rect === null) {
      // 规则被清掉或改成没有下拉的样子，或者这一格滚出了视口
      this.#closeList();
      return;
    }
    this.#placeList(rect);
  }

  /**
   * 点画布箭头时调用：这一格的列表开着就关，没开就开。
   *
   * `event` 是这次点击的指针事件。文档级的"点外面关闭"监听在捕获阶段先于画布的路由
   * 跑，点箭头时它已经把列表关掉了——如果 `toggleList` 只看"列表现在开没开"，就会把
   * 刚关掉的列表又打开，"再点一次关闭"永远做不到。所以认事件：同一个指针事件刚把
   * 这一格的列表关掉，这次就不再打开。
   */
  toggleList(scene: SceneSnapshot, cell: CellRef, event?: Event): void {
    const closed = this.#closedByPointer;
    this.#closedByPointer = undefined;
    if (
      event !== undefined &&
      closed !== undefined &&
      closed.event === event &&
      closed.cell.row === cell.row &&
      closed.cell.column === cell.column
    ) {
      return;
    }
    if (this.#list !== undefined && this.#listTarget?.row === cell.row && this.#listTarget.column === cell.column) {
      this.#closeList();
      return;
    }
    this.#closeList();
    const rule = this.#options.getRule(cell.row, cell.column);
    const items = rule === undefined ? undefined : describeDropdownItems(rule);
    const rect = this.#rectOf(scene, cell);
    if (items === undefined || rect === null) {
      return;
    }
    this.hideAlert();
    this.#openList(cell, rect, items);
  }

  hide(): void {
    this.hideAlert();
    this.#closeList();
  }

  /** 在 `cell` 旁显示错误提示（鼠标停在违规格子上）。列表开着时不显示，免得盖住它。 */
  showAlert(scene: SceneSnapshot, cell: CellRef, message: string): void {
    const rect = this.#rectOf(scene, cell);
    if (rect === null || this.#list !== undefined) {
      return;
    }
    this.hideAlert();
    const bubble = document.createElement("div");
    bubble.setAttribute("role", "alert");
    bubble.dataset.cyValidation = "alert";
    bubble.style.cssText = [
      "position:absolute",
      "z-index:14",
      "box-sizing:border-box",
      "width:200px",
      "padding:8px 10px",
      "background:#fff",
      "border:1px solid #fe4b4b",
      "border-radius:4px",
      "box-shadow:0 2px 8px rgba(0,0,0,.18)",
      "font:12px/1.5 system-ui,sans-serif",
      "color:#202124",
      // 只是提示：不接指针事件，不挡住下面的单元格
      "pointer-events:none"
    ].join(";");
    const heading = document.createElement("div");
    heading.style.cssText = "font-weight:600;color:#d93025";
    heading.textContent = "错误";
    const body = document.createElement("div");
    body.textContent = message;
    bubble.append(heading, body);
    bubble.style.left = `${rect.x}px`;
    bubble.style.top = `${rect.y + rect.height + LIST_GAP}px`;
    this.#options.container.append(bubble);
    this.#alert = bubble;
  }

  hideAlert(): void {
    this.#alert?.remove();
    this.#alert = undefined;
  }

  dispose(): void {
    this.hide();
  }

  #rectOf(scene: SceneSnapshot, cell: CellRef): CellRect | null {
    const merge = this.#options.findMergeAt?.(cell.row, cell.column);
    const range: CellRange = merge ?? {
      startRow: cell.row,
      endRow: cell.row,
      startColumn: cell.column,
      endColumn: cell.column
    };
    return visibleRectForRange(scene, range);
  }

  // -- 下拉列表 ---------------------------------------------------------

  #openList(cell: CellRef, rect: CellRect, items: readonly string[]): void {
    const current = this.#options.getCellValue(cell.row, cell.column);
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
    // 点列表不该把焦点从画布上抢走，否则键盘导航立刻失灵
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
        this.#options.onPick(cell.row, cell.column, item);
      });
      list.append(option);
    }

    this.#options.container.append(list);
    this.#list = list;
    this.#listTarget = cell;
    this.#placeList(rect);

    // 点到列表外面或按 Esc 关掉（Univer 的 `onClickOutside`）。捕获阶段监听，免得被别处
    // 的 stopPropagation 吞掉。点在画布箭头上的那一下也会先在这里关掉，把事件记下来，
    // 见 `toggleList`。
    const onPointerDown = (event: Event): void => {
      const node = event.target as Node | null;
      if (node !== null && list.contains(node)) {
        return;
      }
      this.#closeList();
      this.#closedByPointer = { event, cell };
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
    this.#list.style.top = `${rect.y + rect.height + LIST_GAP}px`;
    this.#list.style.minWidth = `${rect.width}px`;
  }

  #closeList(): void {
    this.#removeOutsideListeners?.();
    this.#removeOutsideListeners = undefined;
    this.#list?.remove();
    this.#list = undefined;
    this.#listTarget = undefined;
  }
}
