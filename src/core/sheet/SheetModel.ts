import type { CellRange } from "../types";
import { CyExcelError, CyExcelErrorCode } from "../errors";
import { CellMatrix } from "./CellMatrix";
import { assertValidBlankCount, resolveCoordinateSpace } from "./coordinateSpace";
import { IndexedTable } from "./IndexedTable";
import { IndexedTableRegistry } from "./IndexedTableRegistry";
import { deserializeSheet, serializeSheet } from "./serialize";
import type { CellRecord, ColumnMeta, RowMeta, SheetSnapshot } from "./types";
import { DEFAULT_BLANK_COLUMNS, DEFAULT_BLANK_ROWS } from "./types";
import {
  buildDeleteRowsPlan,
  buildInsertRowsPlan,
  buildMoveRowsPlan
} from "./rowPlan";
import {
  buildDeleteColumnsShift,
  buildInsertColumnsShift,
  buildMoveColumnsShift
} from "./columnShift";
import { validateAxisChange } from "../layout/structuralChange";

/**
 * extent（最大占用下标加一）换算成 `resolveCoordinateSpace` 要的 `maxOccupied`。
 *
 * 两者差一个"空的怎么表示"：extent 用 0，`maxOccupied` 用 `undefined`——后者必须
 * 是 `undefined` 而不是 -1，否则空表的坐标空间会算成 `blank` 之外再少一格。
 */
function maxOccupiedFrom(extent: number): number | undefined {
  return extent === 0 ? undefined : extent - 1;
}

export interface SheetModelOptions {
  readonly id?: string;
  readonly name?: string;
  readonly blankRows?: number;
  readonly blankColumns?: number;
  readonly defaultRowHeight?: number;
  readonly defaultColumnWidth?: number;
  readonly freeze?: { readonly rows: number; readonly columns: number };
}

/**
 * 一张表的模型：稀疏记录矩阵 + 行列元数据 + 由 `blank*` 派生的坐标空间。
 *
 * 本阶段不接线到 `TableStore` / `LayoutEngine` / `CyExcelCore`——那是阶段 3。
 */
export class SheetModel {
  readonly #matrix: CellMatrix;
  readonly #tables = new IndexedTableRegistry();
  readonly #rowData: IndexedTable<RowMeta>;
  readonly #columnData: IndexedTable<ColumnMeta>;
  #mergeData: readonly CellRange[];
  readonly #id: string;
  readonly #name: string;
  readonly #blankRows: number;
  readonly #blankColumns: number;
  readonly #defaultRowHeight: number;
  readonly #defaultColumnWidth: number;
  readonly #freeze: { readonly rows: number; readonly columns: number } | undefined;

  private constructor(init: {
    matrix: CellMatrix;
    rowData: Readonly<Record<number, RowMeta>>;
    columnData: Readonly<Record<number, ColumnMeta>>;
    mergeData: readonly CellRange[];
    id: string;
    name: string;
    blankRows: number;
    blankColumns: number;
    defaultRowHeight: number;
    defaultColumnWidth: number;
    freeze?: { readonly rows: number; readonly columns: number };
  }) {
    assertValidBlankCount(init.blankRows, "blankRows");
    assertValidBlankCount(init.blankColumns, "blankColumns");
    this.#matrix = init.matrix;
    this.#rowData = this.#tables.register(new IndexedTable<RowMeta>("rowData", "row", init.rowData));
    this.#columnData = this.#tables.register(
      new IndexedTable<ColumnMeta>("columnData", "column", init.columnData)
    );
    this.#mergeData = init.mergeData;
    this.#id = init.id;
    this.#name = init.name;
    this.#blankRows = init.blankRows;
    this.#blankColumns = init.blankColumns;
    this.#defaultRowHeight = init.defaultRowHeight;
    this.#defaultColumnWidth = init.defaultColumnWidth;
    this.#freeze = init.freeze ?? undefined;
  }

  static createEmpty(options: SheetModelOptions = {}): SheetModel {
    return new SheetModel({
      matrix: new CellMatrix(),
      rowData: {},
      columnData: {},
      mergeData: [],
      id: options.id ?? "sheet-1",
      name: options.name ?? "Sheet1",
      blankRows: options.blankRows ?? DEFAULT_BLANK_ROWS,
      blankColumns: options.blankColumns ?? DEFAULT_BLANK_COLUMNS,
      defaultRowHeight: options.defaultRowHeight ?? 28,
      defaultColumnWidth: options.defaultColumnWidth ?? 120,
      ...(options.freeze !== undefined && { freeze: options.freeze })
    });
  }

  static fromSnapshot(snapshot: SheetSnapshot): SheetModel {
    const restored = deserializeSheet(snapshot);
    return new SheetModel({
      matrix: restored.matrix,
      // deserializeSheet 已经拷贝过 rowData/columnData/mergeData，这里不需要再拷一层
      rowData: restored.rowData,
      columnData: restored.columnData,
      mergeData: restored.mergeData,
      id: restored.id,
      name: restored.name,
      blankRows: restored.blankRows,
      blankColumns: restored.blankColumns,
      defaultRowHeight: restored.defaultRowHeight,
      defaultColumnWidth: restored.defaultColumnWidth,
      ...(restored.freeze !== undefined ? { freeze: restored.freeze } : {})
    });
  }

  /** 裸值，不是记录（ADR 0004 §3.6）。 */
  getCell(row: number, column: number): unknown {
    return this.#matrix.get(row, column)?.v;
  }

  getRecord(row: number, column: number): CellRecord | undefined {
    return this.#matrix.get(row, column);
  }

  /** 只改值，保留该格已有的 `t`/`s`/`f`/`si`。 */
  setCell(row: number, column: number, value: unknown): void {
    const existing = this.#matrix.get(row, column);
    this.#matrix.set(row, column, existing ? { ...existing, v: value } : { v: value });
  }

  setRecord(row: number, column: number, record: CellRecord): void {
    this.#matrix.set(row, column, record);
  }

  deleteCell(row: number, column: number): void {
    this.#matrix.delete(row, column);
  }

  /**
   * 坐标空间上界（spec §3.3：`rowCount = maxOccupiedRow + 1 + blankRows`）。
   *
   * 这里的 `maxOccupiedRow` 取的是**格子与每一张登记行表的并集**，与结构操作
   * 用的是同一个数（`#rowExtentIncludingTables()`）。spec §3.3 没有直接回答
   * "只有元数据的行算不算被占用"，但这个类不能对它给出两个答案：`reorderRows`
   * 走的是并集那个答案，`rowCount()` 曾经只问矩阵。后果是实测出来的——
   * `createEmpty({ blankRows: 0 })` 加一张在第 3 行有内容的行表，`rowCount()`
   * 返回 0，于是 `validateAxisChange` 以轴长 0 拒掉每一个结构操作，整张表锁死。
   * 默认的 1000 行空白平时把它盖住了，语义仍然是坏的。
   */
  rowCount(): number {
    return resolveCoordinateSpace(
      maxOccupiedFrom(this.#rowExtentIncludingTables()),
      this.#blankRows
    );
  }

  columnCount(): number {
    return resolveCoordinateSpace(
      maxOccupiedFrom(this.#columnExtentIncludingTables()),
      this.#blankColumns
    );
  }

  /**
   * 最大占用**行号**加一，**只算格子**——与 `rowCount()` 是两个不同的数
   * （spec §3.3 后果 ④）。
   *
   * 它刻意不含侧表：名字说的就是"有多少行装着数据"，而"坐标空间该多大"是另一个
   * 问题，由 `rowCount()` 回答。要含侧表的那个数见 `#rowExtentIncludingTables()`。
   */
  occupiedRowCount(): number {
    const max = this.#matrix.maxOccupiedRow();
    return max === undefined ? 0 : max + 1;
  }

  /**
   * 序列化这一张 sheet。
   *
   * **只输出 `rowData` / `columnData`，经 `registerTable()` 登记的其它侧表不在这里
   * 出现，这是有意的**：那些表的持久化形态是 workbook 级的 `resources` 条目
   * （spec §3 把 `resources` 与 `sheets` 并列放在顶层），由 `WorkbookModel.toSnapshot()`
   * 收拢。在这里顺手输出它们会让同一条数据在快照里出现两次。
   */
  toSnapshot(): SheetSnapshot {
    return serializeSheet({
      id: this.#id,
      name: this.#name,
      blankRows: this.#blankRows,
      blankColumns: this.#blankColumns,
      defaultRowHeight: this.#defaultRowHeight,
      defaultColumnWidth: this.#defaultColumnWidth,
      matrix: this.#matrix,
      rowData: this.#rowData.read(),
      columnData: this.#columnData.read(),
      mergeData: this.#mergeData,
      ...(this.#freeze !== undefined ? { freeze: this.#freeze } : {})
    });
  }

  /**
   * 合并区跟着结构操作正确搬移是一整块 Excel 语义（跨插入点延展、被删范围
   * 完全覆盖则消失、部分覆盖则收缩），属于阶段 5。在那之前**拒绝**而不是
   * 错着搬——静默错搬比拒绝糟得多。仓库先例：headerGroups 同样挡住列结构变更。
   */
  #assertNoMerges(operation: string): void {
    if (this.#mergeData.length > 0) {
      throw new CyExcelError(
        CyExcelErrorCode.MutationNotSupported,
        `${operation} is not supported while merged ranges exist`
      );
    }
  }

  /**
   * 行的"占用范围"是**格子与所有按行索引的侧表**的并集。
   *
   * 名字里的 `IncludingTables` 是与公开的 `occupiedRowCount()` 划清界限用的：
   * 那个只算格子，这个把登记的侧表也算进去，两者不可互换。
   *
   * 这个 extent 同时驱动矩阵的搬运、每一张行表的重键与 `rowCount()` 的坐标空间，
   * 而侧表可以存在于没有任何格子的行上（给空行设行高、给空行挂业务主键，都是正常
   * 状态）。只按格子算的话，超出格子范围的那些条目会被每一次结构操作静默丢弃——
   * 阶段二里这个洞出现过四次。
   *
   * 它对**任意多张**登记在册的行表成立，不再是硬编码的两张：登记了就自动入并集；
   * 但**没有任何东西强制一张新侧表走登记处**——那一步仍然靠人。阶段四会用一个
   * 只能产出已登记表的工厂收掉这个口子。ADR §9 不变量 4、spec §7.2。
   */
  #rowExtentIncludingTables(): number {
    const maxCellRow = this.#matrix.maxOccupiedRow();
    const cellExtent = maxCellRow === undefined ? 0 : maxCellRow + 1;
    return Math.max(cellExtent, this.#tables.rowExtent());
  }

  /**
   * 列轴的同一个并集，只服务 `columnCount()`。
   *
   * 列的结构操作不需要它：`#remapColumns` 是逐个占用列过的纯算术、不以 extent
   * 为界（见 `insertColumns` 上的注释）。但坐标空间需要——一张只有列元数据的表，
   * 轴长必须盖得住那些列，否则 `validateAxisChange` 会挡掉合法操作。
   */
  #columnExtentIncludingTables(): number {
    const maxCellColumn = this.#matrix.maxOccupiedColumn();
    const cellExtent = maxCellColumn === undefined ? 0 : maxCellColumn + 1;
    return Math.max(cellExtent, this.#tables.columnExtent());
  }

  #applyRowPlan(plan: Int32Array, sourceExtent: number): void {
    this.#matrix.applyRowPlan(plan, sourceExtent);
    this.#tables.applyRowPlan(plan);
  }

  /**
   * 六个结构方法共用的参数防线，复用被替换的 `ArrayDataSource` 那一份
   * （`../layout/structuralChange`），不另写。三条口径：
   *
   * ① **放在 `#assertNoMerges` 之前**——参数非法与有没有合并区无关，
   *    先报 `MutationNotSupported` 会误导调用方。
   * ② **轴长取 `rowCount()` / `columnCount()`（坐标空间，spec §3.3），
   *    不是占用范围**——"在空白区里选几行删掉"是 Excel 允许的合法操作，
   *    用占用范围会把它误挡掉。
   * ③ **抛的是 `RangeError` 而不是 `CyExcelError`，不包装**——与被替换的
   *    `ArrayDataSource` 保持同一种失败方式，比与 `#assertNoMerges` 内部
   *    一致更重要；错误类型的统一属于后续公开 API 收口。
   *
   * 没有这道防线时的静默后果：`insertRows(-2, 1)` 让 `newRow - count`
   * 算出 `-1`，而 `-1` 就是 `ROW_PLAN_HOLE`——哨兵撞车，一行真数据变成
   * 一个洞；`moveRows(0, 1, -1)` 走 `rest.splice(负数)` 从末尾倒插；
   * `moveColumns(0, 2, 1)` 算出 `insertAt = -1`，产出负的列键。
   */
  insertRows(start: number, count: number): void {
    validateAxisChange({ axis: "row", operation: "insert", start, count }, this.rowCount());
    this.#assertNoMerges("insertRows");
    const extent = this.#rowExtentIncludingTables();
    this.#applyRowPlan(buildInsertRowsPlan(extent, start, count), extent);
  }

  deleteRows(start: number, count: number): void {
    validateAxisChange({ axis: "row", operation: "delete", start, count }, this.rowCount());
    this.#assertNoMerges("deleteRows");
    const extent = this.#rowExtentIncludingTables();
    this.#applyRowPlan(buildDeleteRowsPlan(extent, start, count), extent);
  }

  moveRows(start: number, count: number, target: number): void {
    validateAxisChange(
      { axis: "row", operation: "move", start, count, target, targetSpace: "pre-removal" },
      this.rowCount()
    );
    this.#assertNoMerges("moveRows");
    const extent = this.#rowExtentIncludingTables();
    this.#applyRowPlan(buildMoveRowsPlan(extent, start, count, target), extent);
  }

  reorderRows(order: Int32Array): void {
    this.#assertNoMerges("reorderRows");
    const extent = this.#rowExtentIncludingTables();
    if (extent > order.length) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidRange,
        `occupied row ${extent - 1} >= order length ${order.length}, would be silently dropped`
      );
    }
    this.#matrix.reorderRows(order);
    this.#tables.applyRowPlan(order);
  }

  /**
   * 列轴对 A1 那一类"占用范围外丢数据"天然免疫：`#remapColumns` 是逐个
   * **占用列**过的纯算术，不以 extent 为界——`deleteColumns(10, 2)` 在只占用
   * 0–4 列时全部 `< 10`，原样不动。所以列轴只补上面那道参数校验。
   */
  insertColumns(start: number, count: number): void {
    validateAxisChange(
      { axis: "column", operation: "insert", start, count },
      this.columnCount()
    );
    this.#assertNoMerges("insertColumns");
    const shift = buildInsertColumnsShift(start, count);
    this.#matrix.insertColumns(start, count);
    this.#tables.applyColumnShift(shift);
  }

  deleteColumns(start: number, count: number): void {
    validateAxisChange(
      { axis: "column", operation: "delete", start, count },
      this.columnCount()
    );
    this.#assertNoMerges("deleteColumns");
    const shift = buildDeleteColumnsShift(start, count);
    this.#matrix.deleteColumns(start, count);
    this.#tables.applyColumnShift(shift);
  }

  moveColumns(start: number, count: number, target: number): void {
    validateAxisChange(
      { axis: "column", operation: "move", start, count, target, targetSpace: "pre-removal" },
      this.columnCount()
    );
    this.#assertNoMerges("moveColumns");
    const shift = buildMoveColumnsShift(start, count, target);
    this.#matrix.moveColumns(start, count, target);
    this.#tables.applyColumnShift(shift);
  }

  /**
   * 行的占用范围——**格子与每一张登记在册的行表的并集**，即
   * `#rowExtentIncludingTables()` 的公开出口。与只算格子的 `occupiedRowCount()`
   * 是两个不同的数，也与坐标空间 `rowCount()` 是两个不同的数。
   *
   * **为什么要公开它。** `reorderRows` 拿这个数判断一个 `order` 合不合法
   * （`extent > order.length` 就会静默丢行），而求逆要在 `MutationCoordinator`
   * 的 prepare 期做**同一个**判断——commit 期才发现不合法就晚了。外面算不出这个
   * 数：`occupiedRowCount()` 不含侧表，`rowCount()` 掺了 `blankRows` 而 `blankRows`
   * 又只在快照里露面；靠这两个反推等于把并集的算术抄第二份。抄第二份就会分叉，
   * 而这里分叉的后果是 prepare 与 commit 对"什么是合法 order"给出两个答案。
   * 所以公开的是同一个私有方法的出口，不是它的一份复制。
   */
  rowExtent(): number {
    return this.#rowExtentIncludingTables();
  }

  getRowMeta(row: number): RowMeta | undefined {
    return this.#rowData.get(row);
  }

  setRowMeta(row: number, meta: RowMeta): void {
    this.#rowData.set(row, meta);
  }

  getColumnMeta(column: number): ColumnMeta | undefined {
    return this.#columnData.get(column);
  }

  setColumnMeta(column: number, meta: ColumnMeta): void {
    this.#columnData.set(column, meta);
  }

  merges(): readonly CellRange[] {
    return this.#mergeData;
  }

  occupiedCount(): number {
    return this.#matrix.occupiedCount();
  }

  /** 最大占用列号加一——与 `columnCount()`（坐标空间）是两个不同的数。 */
  occupiedColumnCount(): number {
    const max = this.#matrix.maxOccupiedColumn();
    return max === undefined ? 0 : max + 1;
  }

  /**
   * 登记一张按下标索引的侧表，此后它跟着 7 个结构操作一起被重键（spec §7.2：
   * 声明 `indexedBy` 的资源，重键由**核心**提供，不外包给插件作者）。
   *
   * 返回传入的那张表，调用方拿它读写；搬运不需要调用方参与。
   *
   * **它只管会话内的重键，不管持久化。** 上面那句 §7.2 读起来像插件入口，但经这里
   * 登记的表**永远不进任何快照**——`toSnapshot()` 只序列化 `rowData` / `columnData`。
   * 要一张表既被重键又被存下来，走 `WorkbookModel.addResource`：那条路先登记进这里，
   * 再由 workbook 在 `resources` 里输出它。
   */
  registerTable<T>(table: IndexedTable<T>): IndexedTable<T> {
    return this.#tables.register(table);
  }

  /** 已登记的侧表名，含核心自己的 `rowData` / `columnData`。 */
  tableNames(): readonly string[] {
    return this.#tables.names();
  }

  /**
   * 取一张已登记的侧表的**活句柄**，用于读写。
   *
   * 与 `tableNames()` 不同，这个是给写路径用的：spec §7.1 的业务主键要在新建行时
   * 写入，光能往返不够。返回的就是登记处里那一张，改动立即对结构操作可见。
   */
  table(name: string): IndexedTable<unknown> | undefined {
    return this.#tables.table(name);
  }

  /**
   * 这张表的 id。
   *
   * `#id` 一直是私有的、只在 `toSnapshot()` 里出现，于是"拿一下 id"要序列化整张
   * 表。任务四的 `WorkbookModel.sheetIds()` 需要它，补一个访问器而不是让调用方
   * 走 `toSnapshot().id`。
   */
  id(): string {
    return this.#id;
  }
}
