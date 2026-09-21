import type { CyExcelCore } from "../CyExcelCore";
import type { CellStyle } from "../style/types";
import type { CellRange } from "../types";

/** 压成一行五组，不做多标签页 ribbon——功能量填不满，填不满的标签页更难用。 */
export type ToolbarGroup = "history" | "format" | "align" | "data" | "structure";

export const TOOLBAR_GROUPS: readonly ToolbarGroup[] = Object.freeze([
  "history",
  "format",
  "align",
  "data",
  "structure"
]);

/**
 * 工具栏与 core 之间唯一的接口面：项只能经由它调公开 API，不碰 Store。
 * `styleOf` 单列出来是为了让 `state.ts` 能脱离真实 CyExcelCore 被测。
 */
export interface ToolbarContext {
  readonly sheet: CyExcelCore;
  /** 当前活动选区 */
  readonly range: CellRange;
  /** 某一格解析后的样式（列默认 + 格样式合并） */
  styleOf(row: number, column: number): CellStyle | undefined;
}

export type ToolbarItemType = "button" | "toggle" | "select" | "color" | "panel" | "dialog";

export interface ToolbarSelectOption {
  readonly value: string | number;
  readonly label: string;
}

export interface ToolbarItem {
  /** 唯一 id。照 Univer 的惯例，取它调的那个能力的名字（`bold`、`undo`、`merge`）。 */
  readonly id: string;
  readonly group: ToolbarGroup;
  /** 组内排序。默认 schema 留了空隙（10、20、30…），方便宿主往中间插。 */
  readonly order: number;
  readonly type: ToolbarItemType;
  /** icons.ts 里的 SVG 串。没有图标的项显示 label 文本。 */
  readonly icon?: string;
  /** 按钮文本 / aria-label，永远要有——图标按钮也需要它给读屏用。 */
  readonly label: string;
  readonly tooltip?: string;
  run?(ctx: ToolbarContext): void | Promise<void>;
  /** toggle 是否高亮。 */
  activated?(ctx: ToolbarContext): boolean;
  disabled?(ctx: ToolbarContext): boolean;
  /** select / color 的当前值。 */
  value?(ctx: ToolbarContext): string | number | undefined;
  /** select 的候选项。 */
  readonly options?: readonly ToolbarSelectOption[];
  /** select / color 选中某个值时调用。 */
  apply?(ctx: ToolbarContext, value: string | number): void | Promise<void>;
  /** panel 自绘，返回 disposer。`onChosen` 把选中的值回报给分段按钮记住。 */
  panel?(ctx: ToolbarContext, host: HTMLElement, onChosen?: (value: string) => void): () => void;
  /**
   * 对话框自绘，返回 disposer。有这个字段的项渲染成普通按钮，点击打开一个**非模态、
   * 可拖动**的对话框（见 `createDialog`）——给查找替换这类"装不进小面板、而且开着时
   * 还得能操作表格"的功能用。`close` 交给内容，让它能自己收起。
   *
   * 同一项只会有一个对话框：已经开着时再点按钮是把焦点送回去，不是开第二个。
   */
  dialog?(ctx: ToolbarContext, body: HTMLElement, close: () => void): () => void;
  /**
   * 分段按钮的初始值。同时有 `panel` 和 `apply` 的项渲染成"图标 + 箭头"：点图标套用
   * 上次选过的值，点箭头展开面板（同 Excel 与 Univer 的 BUTTON_SELECTOR）。
   */
  readonly defaultValue?: string;
}

export type ToolbarSchema = readonly ToolbarItem[];

/** 一项当前的状态。`resolveItemState` 算出来，渲染层照着改 DOM。 */
export interface ToolbarItemState {
  readonly activated: boolean;
  readonly disabled: boolean;
  readonly value: string | number | undefined;
}

export interface ToolbarOptions {
  /** 整体替换默认 schema。 */
  readonly schema?: ToolbarSchema;
  /** 在默认 schema 之外追加项。与 `schema` 同时给时，追加到给定的那份上。 */
  readonly items?: readonly ToolbarItem[];
}
