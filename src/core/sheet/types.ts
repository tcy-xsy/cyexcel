import type { CellRange } from "../types";
import type { CellStyle } from "../style/types";

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
/**
 * 单元格能存的值——**只有这三种 JSON 原生类型**，与 Univer 的 `CellValue` 一字不差
 * （核对记录见 ADR 0004 §2.1），也正是 xlsx 能表达的全部。
 *
 * 没有 `Date`：日期是"数字（序列号）+ 日期 numFmt"，见 `style/dateSerial.ts`。
 * 没有数组和对象：它们既不能安全地 JSON 往返，也没有 xlsx 对应物。把它们挡在类型
 * 之外，是为了让"快照往返无损"和"导出 xlsx 是直译"这两件事从类型上就成立，而不是
 * 靠约定。
 */
export type CellValue = string | number | boolean;

export interface CellRecord {
  /**
   * 值。`undefined` 是唯一的"空"——**这里比 Univer 严一档**：它的 `v` 是
   * `Nullable<CellValue>`（含 `null`），我们没有 `null`。
   *
   * 这个差异只是表示法，不是语义：Univer 的 `isNullCell()` 用宽松相等，`null` 与
   * `undefined` 落进同一个分支，它允许 `null` 落盘但不靠它区分任何东西。我们干脆
   * 不存，是同一语义的更强表示（ADR 0004 §2.1 第 2、3 条）。
   *
   * **代价在导入**：读 Univer 快照时 `{ v: null }` 要在导入层归一成"没有这一格"，
   * 不能原样透传——见 `sheet/import.ts` 的 `valueToRecord()`。
   */
  readonly v: CellValue | undefined;
  /**
   * 值的**物理类型**，与 xlsx `<c t>` 和 Univer 的 `CellValueType` 一一对应。
   *
   * 不是"这一格按什么行为解析"——那是列类型的事。这里只说值本身是什么，于是
   * xlsx 往返零翻译。数字是默认形态，**省略不写**；日期也不在这里，它是
   * "数字（序列号）+ 日期 numFmt"，与 Excel 一致。
   */
  readonly t?: CellValueType;
  /** 样式 id，指向 workbook 级 `styles`。存储层只存 id（ADR §3.4） */
  readonly s?: string;
  /** 公式串 */
  readonly f?: string;
  /** 共享公式 id */
  readonly si?: string;
  /**
   * 这一格的图片 URL（http(s) 或 `data:`）。与 `v` 独立——单元格可以同时有 `img` 和
   * 自己的 `v`（比如 `v` 是这张图的说明文字），也可以只有 `img` 没有 `v`。
   *
   * 这是"图片是不是这一格的内容"的单元格级开关，不依赖列配置：一列里任何一格都能
   * 单独设 `img`，跟这一列有没有配 `renderer: createImageCellRenderer()` 无关——
   * 两者是并存的两条路径，列级 renderer 面向整列都是图片字段的批量场景。
   *
   * 不参与 xlsx 导出（`exportExcel.ts` 只读 `v`），与 `f` 目前的处境一致，是已知
   * 缺口而不是这次改动引入的新问题。
   */
  readonly img?: string;
}

/**
 * 值的物理类型（xlsx `<c t>` 的直译，与 Univer 的 `CellValueType` 同义同值）。
 *
 * `FORCE_STRING` 对应 Excel 的文本格式 `@`：`"00123"` 原样保留，不被当成数字 123。
 * 数字不需要标记（省略 `t` 即数字），所以这里没有 NUMBER 之外的"默认"成员之说。
 */
export const CellValueType = {
  STRING: 1,
  NUMBER: 2,
  BOOLEAN: 3,
  FORCE_STRING: 4
} as const;

export type CellValueType = (typeof CellValueType)[keyof typeof CellValueType];

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
  /** 列默认类型名（注册表里的行为类型，不是 `CellRecord.t` 那个物理类型） */
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
   * 这条资源的下标属于**哪一张 sheet**。
   *
   * `indexedBy: "row" | "column"` 的资源，其下标是某一张 sheet 的行号/列号，
   * 重键也发生在该 sheet 的结构操作上（spec §7.2）。当前
   * `sheetOrder.length === 1`（不变量 6），缺省即那唯一一张，没有歧义。
   *
   * **现在就留这个字段，是照 §3.2 对 `sheets` / `sheetOrder` 的同一个理由**：
   * 本类型进快照，M3 再加就是格式变更，要按 §3.1 升 `version` 并给存量快照
   * 一条迁移路径。资源名是 workbook 唯一而表名是 sheet 唯一，多表下"两张 sheet
   * 各挂一份同名资源"是正常场景，缺了这个字段表达不出来。
   *
   * `indexedBy: "none"` **不得**给出：它是 workbook 级的不透明 blob，没有轴，
   * 也就没有所属的 sheet。
   */
  readonly sheetId?: string;
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
  /** workbook 级样式去重表（字段集见 `style/types.ts`，ADR 0004 §3.4） */
  readonly styles: Readonly<Record<string, CellStyle>>;
  readonly resources: readonly ResourceEntry[];
}

/** 当前 schema 版本。 */
export const SNAPSHOT_VERSION = 1;

/** spec §3.3 的默认值，取 Google Sheets 新建表的量级。 */
export const DEFAULT_BLANK_ROWS = 1000;
export const DEFAULT_BLANK_COLUMNS = 26;
