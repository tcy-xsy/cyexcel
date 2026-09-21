import type { CyExcelCore } from "../CyExcelCore";
import type { CellStyle } from "../style/types";
import type { CellRange } from "../types";

/**
 * 工具栏分组。压成一行五组，不做 Univer 那种六个标签页的 ribbon——我们的功能量
 * 填不满四个标签页，而填不满的标签页比不分页更难用（见 step25-toolbar-plan.md §3）。
 */
export type ToolbarGroup = "history" | "format" | "align" | "data" | "structure";

export const TOOLBAR_GROUPS: readonly ToolbarGroup[] = Object.freeze([
  "history",
  "format",
  "align",
  "data",
  "structure"
]);

/**
 * 工具栏与 core 之间**唯一**的接口面。
 *
 * 项的实现只能经由它调公开 API，不碰 Store、不碰 SheetModel——与
 * `createBasicContextMenuPlugin` 的"不直接操作 Store"是同一条约束
 * （doc/customization-and-hooks.md §11.2）。
 *
 * `styleOf` 单独列出来而不是让调用方自己 `sheet.getCellStyle(...)`，是因为
 * `state.ts` 的纯函数要能在没有真实 `CyExcelCore` 的情况下被测——测试传一个
 * 手搓的 context 进去就行。
 */
export interface ToolbarContext {
  readonly sheet: CyExcelCore;
  /** 当前活动选区 */
  readonly range: CellRange;
  /** 某一格解析后的样式（列默认 + 格样式合并） */
  styleOf(row: number, column: number): CellStyle | undefined;
}

export type ToolbarItemType = "button" | "toggle" | "select" | "color" | "panel";

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
  /** panel 自绘，返回 disposer。 */
  panel?(ctx: ToolbarContext, host: HTMLElement): () => void;
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
