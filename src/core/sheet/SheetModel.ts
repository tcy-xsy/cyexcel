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
import {
  relocateMergesForDelete,
  relocateMergesForInsert,
  relocateMergesForMove
} from "./merges";
import { validateAxisChange } from "../layout/structuralChange";
import { assertValidRowOrder } from "../layout/rowOrder";

/**
 * extent（最大占用下标加一）换算成 `resolveCoordinateSpace` 要的 `maxOccupied`。
 *
 * 两者差一个"空的怎么表示"：extent 用 0，`maxOccupied` 用 `undefined`——后者必须
 * 是 `undefined` 而不是 -1，否则空表的坐标空间会算成 `blank` 之外再少一格。
 */
function maxOccupiedFrom(extent: number): number | undefined {
  return extent === 0 ? undefined : extent - 1;
}

/**
 * `pinCoordinateSpace` 要的空白数：目标轴长减去占用范围。
 *
 * 差额为负说明占用范围已经越过了要钉的那个轴长——那只可能是一次越界的值写，而那
 * 种写在 `TableStore` 的 prepare 期就被挡下了。到了这里再遇上，说明有人绕过了那道
 * 闸，此时**抛**而不是夹到 0：夹的话轴会被悄悄放大到 `extent`，与要钉的值不符，
 * 而这个函数存在的全部理由就是不让轴悄悄变。
 */
function blankFor(target: number, extent: number, axis: "row" | "column"): number {
  const blank = target - extent;
  if (blank < 0) {
    throw new CyExcelError(
      CyExcelErrorCode.InvalidMutation,
      `Cannot pin ${axis} count to ${target}: ${extent} ${axis}s are occupied`
    );
  }
  return blank;
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
  #blankRows: number;
  #blankColumns: number;
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

  /**
   * 克隆一份独立的模型：此后往任一边写，另一边都不受影响。
   *
   * **它是暂存层的地基。** 一次事务开始时克隆出一个影子，逆在影子上求、预览在影子
   * 上前进、commit 才写活模型。这样暂存层里一行新的结构代数都没有——影子就是一个
   * `SheetModel`，跑的是同一份 `insertRows` / `applyRowPlan` / 登记处重键代码。
   *
   * **必须连登记处一起克隆，这条会静默失效。** `invertSheetMutation` 的
   * `capturePluginTables` 是遍历 `tableNames()` 抓被删区间条目的。影子的登记处若是
   * 空的，它返回 `tables: {}`——不报错、测试全绿、插件侧表的撤销数据静默丢失，把
   * spec §11 不变量 8 整条架空。所以这里遍历的是**全部**已登记的表，不是
   * `rowData` / `columnData` 两张。
   *
   * 拷贝的深度：格子矩阵按行桶写时复制（`CellMatrix.clone`），侧表条目拷一层，
   * 记录与插件载荷按引用共享（`types.ts` 顶部的浅共享规则）。合并区是只读数组，
   * 直接共享。
   */
  clone(): SheetModel {
    const copy = new SheetModel({
      matrix: this.#matrix.clone(),
      // 构造函数会把这两个交给 `IndexedTable` 的构造函数，那里拷一层
      rowData: this.#rowData.read(),
      columnData: this.#columnData.read(),
      mergeData: this.#mergeData,
      id: this.#id,
      name: this.#name,
      blankRows: this.#blankRows,
      blankColumns: this.#blankColumns,
      defaultRowHeight: this.#defaultRowHeight,
      defaultColumnWidth: this.#defaultColumnWidth,
      ...(this.#freeze !== undefined ? { freeze: this.#freeze } : {})
    });
    for (const name of this.#tables.names()) {
      // 这两张由构造函数登记过了，再登记一次会撞名字
      if (name === "rowData" || name === "columnData") {
        continue;
      }
      const table = this.#tables.table(name);
      if (table) {
        copy.registerTable(table.clone());
      }
    }
    return copy;
  }

  /** 裸值，不是记录（ADR 0004 §3.6）。 */
  getCell(row: number, column: number): unknown {
    return this.#matrix.get(row, column)?.v;
  }

  getRecord(row: number, column: number): CellRecord | undefined {
    return this.#matrix.get(row, column);
  }

  /**
   * Iterates every occupied cell with its bare value — the "cells that
   * actually have something in them" walk Step 23's search scanner needs,
   * as opposed to scanning the full (possibly million-row) coordinate space.
   */
  *occupiedCells(): Generator<{ row: number; column: number; value: unknown }> {
    for (const [row, bucket] of this.#matrix.rows()) {
      for (const key of Object.keys(bucket)) {
        const column = Number(key);
        yield { row, column, value: (bucket[column] as CellRecord).v };
      }
    }
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
   * 把坐标空间钉回给定的大小，差额由 `blankRows` / `blankColumns` 吸收。
   *
   * **这是"值写不改坐标空间"这条不变量的执行体**（另一半是 `TableStore` 里那道
   * 拒绝越界写的闸）。轴是派生量（`maxOccupied + 1 + blank`，spec §3.3），于是
   * 单纯写值就能把它挪动，两个方向都会：
   *
   * - **缩**：清空最后一行的每一个格子，`maxOccupied` 退回去，`rowCount()` 跟着变
   *   小。旧的行数组模型里这做不到——`ClearRangeCommand` 清空整张表之后表还是那么
   *   多行。稀疏矩阵里它是默认行为，而这条变更只带一条 `range-value` 补丁：渲染层
   *   不知道表变短了。
   * - **涨**：往一行**空白行**里写第一个值（导入把数据末尾的空行折进了 `blankRows`），
   *   `maxOccupied` 前进而 `blank` 不动，轴凭空多出一行。
   *
   * 两种都不是"清空/填写一个格子"该有的效果。这里的口径是：格子的占用范围变了，
   * 空白的数量就反向变同样多——**被清空的尾行本来就变回了空白行**，这正是
   * `blankRows` 这个字段的含义（与 `rowsToWorkbookSnapshot` 折算尾部空行是同一条
   * 口径）。因此它也自然是自逆的：逆变更把值写回去时，同一条规则把 `blank` 还原。
   *
   * 结构变更**不**走这里——插入/删除行列本来就该改坐标空间。
   */
  pinCoordinateSpace(rowCount: number, columnCount: number): void {
    // 两个都先算出来再赋值：列那一边抛的时候，行不该已经被改掉一半
    const blankRows = blankFor(rowCount, this.#rowExtentIncludingTables(), "row");
    const blankColumns = blankFor(columnCount, this.#columnExtentIncludingTables(), "column");
    this.#blankRows = blankRows;
    this.#blankColumns = blankColumns;
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
   * `insertRows` / `deleteRows` / `moveRows` / `insertColumns` / `deleteColumns` /
   * `moveColumns` 六个结构方法已经改用 `merges.ts` 的 `relocateMergesFor*` 正确
   * 搬移合并区（跨插入点延展、被删范围完全覆盖则消失、部分覆盖则收缩、跨移动
   * 边界则拒绝），不再调用这个方法——现在只有 `reorderRows` 还用它。
   *
   * `reorderRows` 接受的是一个**任意排列**，不是"挪一段连续区间"，没有良定义的
   * "合并区怎么跟着走"规则（挪到哪算合并区的哪一部分被打散，排列本身说不清楚），
   * 所以它继续无条件拒绝存在合并区的情形，而不是尝试 relocate。
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
   * 结构操作之后把坐标空间钉到它**应该**是的那个数：插/删是 `之前 ± count`，
   * 移动与重排是 `之前`。差额由 `blank*` 吸收。
   *
   * **为什么每一个结构方法都要显式钉一次。** 轴是派生量
   * （`maxOccupied + 1 + blank`，spec §3.3），而这些方法搬的是**占用范围**——
   * 于是轴的变化量由"占用范围恰好被搬动了多少"决定，而不是由操作本身决定，
   * 两者在占用范围的边界上就分家：
   *
   * - **插在占用范围之外**（尾部空白区里，`start >= extent`）：没有格子要搬，
   *   `extent` 与 `blank` 都不动，轴原地不动——插了一行而轴没变长。删同理。
   * - **把仅有的占用行往回移**：`extent` 自己缩了，`blank` 不动，轴凭空变短——
   *   移动改变了行数。
   *
   * 平时被默认的 1000 行/列空白盖住，看起来只是"末尾插入不生效"；把空白调到 0
   * 就直接坏给你看，而且**撤销会抛**：`InsertColumnsCommand(3, …)` 在三列的表上让
   * 布局涨到 4 列而模型仍是 3，逆变更于是拿 `delete-columns 3,1` 去撞一张三列的表，
   * `validateAxisChange` 报 `delete range is outside axis length`。
   *
   * 口径与 `pinCoordinateSpace` 是同一条——"占用范围变了多少，空白就反向变多少"，
   * 只是这里的目标轴长是显式给出的，而不是"保持不变"。因此它对落在占用范围
   * **里面**的操作是恒等的（`extent` 自己动了那么多，`blank` 不动），只在空白区
   * 那一侧才真的改 `blank`。
   */
  #pinRowSpace(target: number): void {
    this.#blankRows = blankFor(target, this.#rowExtentIncludingTables(), "row");
  }

  #pinColumnSpace(target: number): void {
    this.#blankColumns = blankFor(target, this.#columnExtentIncludingTables(), "column");
  }

  /**
   * 六个结构方法共用的参数防线，复用被替换的 `ArrayDataSource` 那一份
   * （`../layout/structuralChange`），不另写。三条口径：
   *
   * ① **放在合并区处理之前**（`relocateMergesFor*` 或 `reorderRows` 用的
   *    `#assertNoMerges`）——参数非法与有没有合并区无关，先报合并区那一类
   *    错误会误导调用方。
   * ② **轴长取 `rowCount()` / `columnCount()`（坐标空间，spec §3.3），
   *    不是占用范围**——"在空白区里选几行删掉"是 Excel 允许的合法操作，
   *    用占用范围会把它误挡掉。
   * ③ **抛的是 `RangeError` 而不是 `CyExcelError`，不包装**——与被替换的
   *    `ArrayDataSource` 保持同一种失败方式，比与合并区错误内部一致更
   *    重要；错误类型的统一属于后续公开 API 收口。
   *
   * 没有这道防线时的静默后果：`insertRows(-2, 1)` 让 `newRow - count`
   * 算出 `-1`，而 `-1` 就是 `ROW_PLAN_HOLE`——哨兵撞车，一行真数据变成
   * 一个洞；`moveRows(0, 1, -1)` 走 `rest.splice(负数)` 从末尾倒插；
   * `moveColumns(0, 2, 1)` 算出 `insertAt = -1`，产出负的列键。
   */
  insertRows(start: number, count: number): void {
    const space = this.rowCount();
    validateAxisChange({ axis: "row", operation: "insert", start, count }, space);
    this.#mergeData = relocateMergesForInsert(this.#mergeData, "row", start, count);
    const extent = this.#rowExtentIncludingTables();
    this.#applyRowPlan(buildInsertRowsPlan(extent, start, count), extent);
    this.#pinRowSpace(space + count);
  }

  deleteRows(start: number, count: number): void {
    const space = this.rowCount();
    validateAxisChange({ axis: "row", operation: "delete", start, count }, space);
    this.#mergeData = relocateMergesForDelete(this.#mergeData, "row", start, count);
    const extent = this.#rowExtentIncludingTables();
    this.#applyRowPlan(buildDeleteRowsPlan(extent, start, count), extent);
    this.#pinRowSpace(space - count);
  }

  moveRows(start: number, count: number, target: number): void {
    const space = this.rowCount();
    validateAxisChange(
      { axis: "row", operation: "move", start, count, target, targetSpace: "pre-removal" },
      space
    );
    const relocated = relocateMergesForMove(this.#mergeData, "row", start, count, target);
    if (!relocated.ok) {
      throw new CyExcelError(CyExcelErrorCode.MergeSpansMoveBoundary, "moveRows would split a merged range");
    }
    this.#mergeData = relocated.merges;
    const extent = this.#rowExtentIncludingTables();
    this.#applyRowPlan(buildMoveRowsPlan(extent, start, count, target), extent);
    this.#pinRowSpace(space);
  }

  /**
   * `reorderRows` 的参数防线不是 `validateAxisChange`（它判的是 start/count/target，
   * 而这里要判的是"是不是一个排列"），但**顺序口径与上面六个方法一致**：先校验
   * 参数、再查合并区。理由就是那六个共用的注释 ①——"参数非法与有没有合并区无关，
   * 先报 `MutationNotSupported` 会误导调用方"——这条理由对一个不合法的排列同样
   * 成立：`reorderRows([0, 0, 5])` 与有没有合并区毫无关系。
   *
   * 这里此前是反过来的（先查合并区），是阶段三补合并区防线时留下的不一致，
   * `invert.ts` 只能逐成员镜像它；阶段 4A 任务四把它拉齐，两边一起改。
   *
   * `assertValidRowOrder` 因此要在这里**显式**调一次，而不是等 `CellMatrix.reorderRows`
   * 里那次——那一次在合并区检查之后，顺序就又反了。`CellMatrix` 那一次照旧保留：
   * 它是那个类自己的入口防线，对不经本方法的调用方仍然必要，而这个函数是纯校验、
   * 幂等，多调一次没有副作用。
   */
  reorderRows(order: Int32Array): void {
    const space = this.rowCount();
    const extent = this.#rowExtentIncludingTables();
    if (extent > order.length) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidRange,
        `occupied row ${extent - 1} >= order length ${order.length}, would be silently dropped`
      );
    }
    assertValidRowOrder(order, order.length);
    this.#assertNoMerges("reorderRows");
    this.#matrix.reorderRows(order);
    this.#tables.applyRowPlan(order);
    // 重排也要钉：把仅有的那一行占用行排到前面去，`extent` 就自己缩了，轴会跟着
    // 凭空变短（`#pinRowSpace` 的注释里的第二种情形）。
    this.#pinRowSpace(space);
  }

  /**
   * 列轴对 A1 那一类"占用范围外丢数据"天然免疫：`#remapColumns` 是逐个
   * **占用列**过的纯算术，不以 extent 为界——`deleteColumns(10, 2)` 在只占用
   * 0–4 列时全部 `< 10`，原样不动。所以列轴只补上面那道参数校验。
   */
  insertColumns(start: number, count: number): void {
    const space = this.columnCount();
    validateAxisChange({ axis: "column", operation: "insert", start, count }, space);
    this.#mergeData = relocateMergesForInsert(this.#mergeData, "column", start, count);
    const shift = buildInsertColumnsShift(start, count);
    this.#matrix.insertColumns(start, count);
    this.#tables.applyColumnShift(shift);
    this.#pinColumnSpace(space + count);
  }

  deleteColumns(start: number, count: number): void {
    const space = this.columnCount();
    validateAxisChange({ axis: "column", operation: "delete", start, count }, space);
    this.#mergeData = relocateMergesForDelete(this.#mergeData, "column", start, count);
    const shift = buildDeleteColumnsShift(start, count);
    this.#matrix.deleteColumns(start, count);
    this.#tables.applyColumnShift(shift);
    this.#pinColumnSpace(space - count);
  }

  moveColumns(start: number, count: number, target: number): void {
    const space = this.columnCount();
    validateAxisChange(
      { axis: "column", operation: "move", start, count, target, targetSpace: "pre-removal" },
      space
    );
    const relocated = relocateMergesForMove(this.#mergeData, "column", start, count, target);
    if (!relocated.ok) {
      throw new CyExcelError(
        CyExcelErrorCode.MergeSpansMoveBoundary,
        "moveColumns would split a merged range"
      );
    }
    this.#mergeData = relocated.merges;
    const shift = buildMoveColumnsShift(start, count, target);
    this.#matrix.moveColumns(start, count, target);
    this.#tables.applyColumnShift(shift);
    this.#pinColumnSpace(space);
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

  /**
   * 遍历所有**被自定义过**的行——稀疏表里真实存在的那些条目，不是 0..rowCount。
   *
   * 有它是因为「哪些行的行高来源是 auto」这个问题只能靠遍历回答，而遍历坐标空间
   * 是不可接受的：空间可以到百万级，条目通常只有几十个（自动换行测量过的那些）。
   * 唯一的替代是在外面另存一张 auto 行的集合，那就又多一份能与权威分叉的存储
   * ——ADR 0004 阶段 4B T8 要消灭的正是这个（设计稿 §4.1）。
   *
   * 产出的是活对象的引用（同 `IndexedTable.read()` 的口径），遍历期间不要写这张表。
   */
  *rowMetaEntries(): IterableIterator<readonly [number, RowMeta]> {
    for (const [key, meta] of Object.entries(this.#rowData.read())) {
      yield [Number(key), meta];
    }
  }

  setRowMeta(row: number, meta: RowMeta): void {
    this.#rowData.set(row, meta);
  }

  /**
   * 删掉这一行的元数据条目。
   *
   * 与 `deleteCell` 同一条理由：`rowData` 是稀疏表，"没有这个键"就是"这一行没有
   * 被自定义过"本身（spec §11 不变量 1 对格子的那条口径，行元数据同理）。写一个
   * 空的 `{}` 进去不等价——它会让这一行进入 `rowExtent()` 的并集，把坐标空间撑大
   * 一格，快照里也会多出一条无内容的条目。
   *
   * 有这个方法，是因为 `set-row-meta` 的逆在"原先没有条目"时携带的是 `undefined`
   * （见 `mutation.ts`），施加它必须能把条目真的删掉。
   */
  deleteRowMeta(row: number): void {
    this.#rowData.delete(row);
  }

  getColumnMeta(column: number): ColumnMeta | undefined {
    return this.#columnData.get(column);
  }

  setColumnMeta(column: number, meta: ColumnMeta): void {
    this.#columnData.set(column, meta);
  }

  /** 删掉这一列的元数据条目，理由同 `deleteRowMeta`。 */
  deleteColumnMeta(column: number): void {
    this.#columnData.delete(column);
  }

  merges(): readonly CellRange[] {
    return this.#mergeData;
  }

  /**
   * 整体替换合并区列表——`mergeCells` / `unmergeCells` 等命令层入口写这个，几何
   * 计算（重叠判定、吸收）由 `merges.ts` 的纯函数在调用方那一侧做完，这里只管
   * 冻结落地。**这一处主动冻结，是 `#mergeData` 目前唯一会冻结的写入路径**——
   * 构造函数、`fromSnapshot`/`deserializeSheet`、以及六个结构方法（insert/
   * delete/move 行与列）直接赋值 `#mergeData` 的地方都不冻结，读者不要把这里的
   * 冻结误当成某种全局约定。
   */
  setMerges(merges: readonly CellRange[]): void {
    this.#mergeData = Object.freeze(merges.map((range) => Object.freeze({ ...range })));
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
