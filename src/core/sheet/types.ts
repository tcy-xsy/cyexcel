import type { CellRange } from "../types";

/**
 * 一个单元格。`v` 是唯一必需字段（spec §11 不变量 2）。
 * 未占用的格子在 `cellData` 中不存在，不是存一条空记录（不变量 1）。
 *
 * **不可变契约：** `CellRecord` 在快照与活矩阵（`CellMatrix`）之间按引用共享，
 * 不做深拷贝——逐格拷贝正是这个设计要省掉的分配开销。因此一条记录一旦创建就
 * 必须当作不可变对象：任何"改一格"都要整体替换成一条新记录（`{ ...old, v: next }`
 * 之类），绝不能就地修改已存在记录的字段。就地修改会通过共享引用直接污染快照
 * 或活矩阵中的另一份。
 */
export interface CellRecord {
  /** 值 */
  readonly v: unknown;
  /**
   * 语义类型名，覆盖列默认。
   *
   * 注意：与 Univer 的 `t` **同名不同义**——Univer 的是原始类型枚举
   * （1=string/2=number/3=boolean/4=force text），我们的是注册表里的类型名
   * （"text"/"number"/"date"/…），描述行为而非原始类型。见 spec §4.1。
   */
  readonly t?: string;
  /** 样式 id，指向 workbook 级 `styles`。存储层只存 id（ADR §3.4） */
  readonly s?: string;
  /** 公式串 */
  readonly f?: string;
  /** 共享公式 id */
  readonly si?: string;
}

/** 行高来源。`default` 表示该行不出现在 `rowData` 里，故不是本类型的成员。 */
export type RowHeightSource = "manual" | "auto";

export interface RowMeta {
  /** 行高，逻辑像素 */
  readonly h?: number;
  /** 行高来源。持久化它是为了让加载时能先用缓存值、避免重测抖动（spec §5.2） */
  readonly hp?: RowHeightSource;
  /** 隐藏 */
  readonly hd?: boolean;
}

export interface ColumnMeta {
  readonly w?: number;
  readonly hd?: boolean;
  /** 列默认语义类型 */
  readonly t?: string;
  /** 列默认样式 id */
  readonly s?: string;
  readonly minW?: number;
  readonly maxW?: number;
  readonly sortable?: boolean;
}

export interface ResourceEntry {
  readonly name: string;
  /**
   * 声明索引方式。`row`/`column` 的资源，其重键与逆操作由核心提供并纳入同一
   * transaction（spec §7.2）；`none` 是不透明 blob，核心不碰。
   */
  readonly indexedBy: "row" | "column" | "none";
  readonly data: unknown;
}

export interface SheetSnapshot {
  readonly id: string;
  readonly name: string;
  /** 数据之后保留多少空白行。坐标空间由它派生，`rowCount` 不存（spec §3.3） */
  readonly blankRows: number;
  readonly blankColumns: number;
  readonly defaultRowHeight: number;
  readonly defaultColumnWidth: number;
  readonly freeze?: { readonly rows: number; readonly columns: number };
  /** 稀疏：row -> column -> CellRecord */
  readonly cellData: Readonly<Record<number, Readonly<Record<number, CellRecord>>>>;
  readonly rowData: Readonly<Record<number, RowMeta>>;
  readonly columnData: Readonly<Record<number, ColumnMeta>>;
  readonly mergeData: readonly CellRange[];
}

export interface WorkbookSnapshot {
  /** schema 版本。每次不兼容变更加一（spec §3.1） */
  readonly version: number;
  readonly sheetOrder: readonly string[];
  readonly sheets: Readonly<Record<string, SheetSnapshot>>;
  /** workbook 级样式去重表 */
  readonly styles: Readonly<Record<string, unknown>>;
  readonly resources: readonly ResourceEntry[];
}

/** 当前 schema 版本。 */
export const SNAPSHOT_VERSION = 1;

/** spec §3.3 的默认值，取 Google Sheets 新建表的量级。 */
export const DEFAULT_BLANK_ROWS = 1000;
export const DEFAULT_BLANK_COLUMNS = 26;
