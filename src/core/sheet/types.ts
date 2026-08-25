import type { CellRange } from "../types";

/**
 * # 别名规则（模块级，一条规则推出四条形态）
 *
 * 存储层一律**浅共享**：容器拷一层、值按引用共享。深拷贝任意插件 JSON 会把成本和
 * 失败模式（函数、循环引用、宿主对象）转移到核心，所以不做。由此得到四条具体形态：
 *
 * 1. `CellRecord` 的值不可就地改；
 * 2. `ResourceEntry.data` 的嵌套值不可就地改（`"none"` 除外，它**整体**是别名）；
 * 3. `IndexedTable.read()` 交回的是**活容器**，核心会就地改它；
 * 4. `WorkbookModel.resource()` 交回的是**拷贝**，写进去会丢，要写用 `resourceTable()`。
 *
 * 四条各自的理由写在各自的声明上，这里只保证它们读起来是同一条规则的四个面，
 * 而不是四条互相独立、需要各背一遍的约定。
 */

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

/** 一行的元数据。存在 `rowData` 表里，适用本文件顶部的别名规则（形态 3）。 */
export interface RowMeta {
  /** 行高，逻辑像素 */
  readonly h?: number;
  /** 行高来源。持久化它是为了让加载时能先用缓存值、避免重测抖动（spec §5.2） */
  readonly hp?: RowHeightSource;
  /** 隐藏 */
  readonly hd?: boolean;
}

/** 一列的元数据。存在 `columnData` 表里，适用本文件顶部的别名规则（形态 3）。 */
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
  /**
   * 资源数据。别名的强度**按 `indexedBy` 分成两种**，不是一种——见本文件顶部的
   * 别名规则形态 2。
   *
   * **`"row"` / `"column"`（拷一层）：** 与 `CellRecord` 一样的浅层共享契约。
   * `resourceToTable` 把 `data` 展开进一张 `IndexedTable`，`tableToResource` 再展开
   * 出来，所以顶层容器是各自独立的，**嵌套值是共享引用**——由调用方原始条目、活表
   * 与转出的条目同时持有。嵌套值必须被当作不可变对象，就地修改会通过这份共享引用
   * 污染原始条目和活表，且绕过 mutation，违反 spec §11 不变量 9（`resources` 只承担
   * save/load 往返，可逆性一律走 mutation）。
   *
   * **`"none"`（整体别名，一层都没拷）：** `addResource` 存进去的 `data` 就是调用方
   * 交进来的那个引用本身，`resource()` 交回存着它的那条条目，`toSnapshot()` 又把同
   * 一条条目放进 `resources` 数组。三处的 `data` 自始至终是同一个对象，因此改动调用方
   * 手里那份会**直接**改到快照里去。这里不加拷贝是有意的：
   * `"none"` 的 `data` 可以是原始值、数组、任何东西，`{ ...data }` 对它不是"拷一层"
   * 而是把它变成另一个形状。核心不碰不透明 blob（spec §7.2），也包括不替它拷贝。
   */
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
