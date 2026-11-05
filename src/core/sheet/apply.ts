import { CyExcelError, CyExcelErrorCode } from "../errors";
import type { ColumnCells, RowCells, SheetMutation, SideTableEntries } from "./mutation";
import type { SheetModel } from "./SheetModel";
import type { CellRecord, ColumnMeta, RowMeta } from "./types";

/**
 * 把一条 `SheetMutation` 施加到模型上。
 *
 * **它在 4A 是测试里的一个本地 helper，4B 提升成生产代码。** 提升不是整理代码：
 * 暂存层（`SheetTransaction`）的正确性论证依赖"影子模型与活模型跑的是同一份施加
 * 逻辑"——commit 走 `applySheetMutation(live, m)`、预览走
 * `applySheetMutation(shadow, m)`、回滚走 `applySheetMutation(live, inv)`，三条路
 * 一份实现。留在测试里就意味着生产另写一份，而这条分支已经连续三次栽在"同一套
 * 语义两份实现"上（阶段二的侧表 extent、阶段三的 `reorderRows` 校验顺序、阶段 4A
 * 的合法 `order`）。附带的收获是往返测试从此测的是生产代码。
 *
 * **本函数只调 `SheetModel` 的公开方法**，不碰它的内部。`restore-*` 是"插空行/列
 * ＋逐格写回 ＋写回元数据 ＋写回插件侧表条目"的组合，`SheetModel.insertRows` 确实
 * 只插空行、不带记录，但补写回去用的也都是既有的公开方法，组合出来的效果与"带
 * 记录的插入"在快照上逐字段一致。
 *
 * **校验不在这里。** 参数与前置条件（合并区、`restore-*` 的 `tables` 形状）由
 * `invertSheetMutation` 的 `assertMutationIsAppliable` 在 prepare 期查完，理由是
 * `MutationCoordinator` 要求"能查出来的都在 prepare 查出来，commit 因此不会失败"。
 * 这里只留一道守卫：`writeTables` 点名一张没登记的表时抛，而不是静默跳过——静默
 * 跳过等于把插件侧表的撤销数据丢掉，正是 4A 任务四要修的那个洞。
 */
export function applySheetMutation(model: SheetModel, mutation: SheetMutation): void {
  switch (mutation.type) {
    case "set-cell":
      writeValues(model, () => {
        writeCell(model, mutation.row, mutation.column, mutation.record);
      });
      return;

    case "set-range": {
      const { startRow, startColumn, endRow, endColumn } = mutation.range;
      writeValues(model, () => {
        // 尺寸取自 `range` 而不是 `records`：`records` 可能是不齐的，缺的那几格按
        // 空格处理（与 `invert.ts` 的 `captureRange` 是同一条口径，两边必须一致，
        // 否则 `set-range` 的往返会在缺格上分叉）。
        for (let row = startRow; row <= endRow; row += 1) {
          for (let column = startColumn; column <= endColumn; column += 1) {
            writeCell(model, row, column, mutation.records[row - startRow]?.[column - startColumn]);
          }
        }
      });
      return;
    }

    case "insert-rows":
      withCoordinateSpace(model, { rows: mutation.rows.length }, () => {
        model.insertRows(mutation.start, mutation.rows.length);
        writeRows(model, mutation.start, mutation.rows, undefined);
      });
      return;

    case "restore-rows":
      withCoordinateSpace(model, { rows: mutation.rows.length }, () => {
        model.insertRows(mutation.start, mutation.rows.length);
        writeRows(model, mutation.start, mutation.rows, mutation.meta);
        writeTables(model, mutation.start, mutation.tables);
        model.setMerges(mutation.mergesBefore);
        model.setValidations(mutation.validationsBefore);
      });
      return;

    case "delete-rows":
      withCoordinateSpace(model, { rows: -mutation.count }, () => {
        model.deleteRows(mutation.start, mutation.count);
      });
      return;

    case "move-rows":
      withCoordinateSpace(model, {}, () => {
        model.moveRows(mutation.start, mutation.count, mutation.target);
      });
      return;

    case "reorder-rows":
      withCoordinateSpace(model, {}, () => {
        model.reorderRows(mutation.order);
      });
      return;

    case "insert-columns":
      withCoordinateSpace(model, { columns: mutation.count }, () => {
        model.insertColumns(mutation.start, mutation.count);
      });
      return;

    case "delete-columns":
      withCoordinateSpace(model, { columns: -mutation.count }, () => {
        model.deleteColumns(mutation.start, mutation.count);
      });
      return;

    case "restore-columns":
      withCoordinateSpace(model, { columns: mutation.count }, () => {
        model.insertColumns(mutation.start, mutation.count);
        for (let offset = 0; offset < mutation.count; offset += 1) {
          const column = mutation.cells[offset];
          if (column) {
            // `ColumnCells` 的键是**行号**——写成列号就是一次静默转置
            for (const rowKey of Object.keys(column)) {
              writeCell(model, Number(rowKey), mutation.start + offset, column[Number(rowKey)]);
            }
          }
          const meta = mutation.meta[offset];
          if (meta !== undefined) {
            model.setColumnMeta(mutation.start + offset, meta);
          }
        }
        writeTables(model, mutation.start, mutation.tables);
        model.setMerges(mutation.mergesBefore);
        model.setValidations(mutation.validationsBefore);
      });
      return;

    case "move-columns":
      withCoordinateSpace(model, {}, () => {
        model.moveColumns(mutation.start, mutation.count, mutation.target);
      });
      return;

    // 元数据写也是"值写"那一类：行高、列宽、隐藏位改不了表有多少行。`rowData` /
    // `columnData` 是登记在册的侧表，占用范围算它们一份（`#rowExtentIncludingTables`），
    // 所以给尾部空白区里的一行设行高会把轴推长——正是这里要挡住的。
    case "set-row-meta":
      writeValues(model, () => {
        if (mutation.meta === undefined) {
          model.deleteRowMeta(mutation.row);
        } else {
          model.setRowMeta(mutation.row, mutation.meta);
        }
      });
      return;

    case "set-column-meta":
      writeValues(model, () => {
        if (mutation.meta === undefined) {
          model.deleteColumnMeta(mutation.column);
        } else {
          model.setColumnMeta(mutation.column, mutation.meta);
        }
      });
      return;

    // 合并区变化不改坐标空间——仍然套 `writeValues`，与 `set-row-meta` 同款，
    // 钉住这条纪律，即使 `setMerges` 实际上不可能触发轴变化。
    case "set-merges":
      writeValues(model, () => {
        model.setMerges(mutation.merges);
      });
      return;

    case "set-validations":
      writeValues(model, () => {
        model.setValidations(mutation.validations);
      });
      return;
  }
}

/**
 * 跑一条变更，并把坐标空间钉到它**应该**变成的那个大小：`之前 + delta`。
 *
 * **每一条变更都从这里走，`delta` 就是那张表**——一条变更把轴改动多少，只由它是
 * 哪一条决定，不由它恰好搬动了多少占用范围决定。轴是派生量
 * （`maxOccupied + 1 + blank`，spec §3.3），两者在占用范围的边界上分家，两个方向
 * 都会：
 *
 * - **值写让轴变了**：清空最后一行的每一个格子，`maxOccupied` 退回去，轴跟着缩；
 *   往尾部空白区里写第一个值，轴凭空多一行。理由与实现在
 *   `SheetModel.pinCoordinateSpace` 的注释里。
 * - **结构变更让轴变得不对**：`insert-rows` 是"插空行 ＋ 写回内容"的组合，
 *   `SheetModel.insertRows` 把轴涨了 `count`，随后的写值又把占用范围推上去，轴于是
 *   涨了两次——影子模型里插一行，`getRowCount()` 从 1 变成 3。`restore-*` 同理。
 *
 * 所以钉在**整条变更的两端**，而不是每个格子上，也不是只在 `SheetModel` 那一层：
 * 中途的状态（插好了空行还没写回内容、清了一半的尾行）不是任何人观察得到的状态，
 * 逐格去钉会把 `blank` 在循环里推来推去，还会让"先清后写"的顺序影响最终结果。
 *
 * `SheetModel` 的那几个结构方法各自也钉一次，两者不冲突：那一层守的是**直接调用
 * 它的人**拿到的契约（"插 n 行，轴就长 n 行"），这一层守的是**组合**出来的整条
 * 变更。同一个目标值钉两次是幂等的。
 */
function withCoordinateSpace(
  model: SheetModel,
  delta: { readonly rows?: number; readonly columns?: number },
  run: () => void
): void {
  const rowCount = model.rowCount();
  const columnCount = model.columnCount();
  run();
  model.pinCoordinateSpace(rowCount + (delta.rows ?? 0), columnCount + (delta.columns ?? 0));
}

/** 只写值的那几条：轴与跑之前一样大。 */
function writeValues(model: SheetModel, write: () => void): void {
  withCoordinateSpace(model, {}, write);
}

function writeCell(model: SheetModel, row: number, column: number, record: CellRecord | undefined): void {
  if (record === undefined) {
    // spec §11 不变量 1：未占用的格子在 `cellData` 中不存在，不是存一条空记录
    model.deleteCell(row, column);
  } else {
    model.setRecord(row, column, record);
  }
}

function writeRows(
  model: SheetModel,
  start: number,
  rows: readonly RowCells[],
  meta: readonly (RowMeta | undefined)[] | undefined
): void {
  for (let offset = 0; offset < rows.length; offset += 1) {
    const cells = rows[offset];
    if (cells) {
      // `RowCells` 的键是**列号**，与 `ColumnCells` 互为转置
      for (const columnKey of Object.keys(cells)) {
        writeCell(model, start + offset, Number(columnKey), cells[Number(columnKey)]);
      }
    }
    const rowMeta = meta?.[offset];
    if (rowMeta !== undefined) {
      model.setRowMeta(start + offset, rowMeta);
    }
  }
}

/**
 * 把 `restore-*` 带回来的插件侧表条目写回原位。
 *
 * 用的仍然是公开方法：`SheetModel.table(name)` 交回登记处里那张表的**活句柄**
 * （那个方法的注释写明它就是给写路径用的），与 `setRowMeta` / `setColumnMeta`
 * 属同一类。
 *
 * `undefined` 的位置表示那一行/列在这张表里本来就没有条目，跳过不写——稀疏表里
 * "有没有这个键"就是"有没有这条条目"本身（spec §11 不变量 1 对格子的同一条口径），
 * 写一个 `undefined` 进去等于凭空造一条。
 */
function writeTables(model: SheetModel, start: number, tables: SideTableEntries): void {
  for (const name of Object.keys(tables)) {
    const entries = tables[name] as readonly unknown[];
    const table = model.table(name);
    if (!table) {
      // 这条在正常路径上到不了——`assertRestoreTablesAreWritable` 在 prepare 期
      // 已经查过。留着是因为"点名一张没登记的表"的另一种处理是静默跳过，而静默
      // 跳过丢的正是这个字段存在的理由（spec §11 不变量 8）。
      throw new CyExcelError(
        CyExcelErrorCode.InvalidMutation,
        `restore names table "${name}", which is not registered on this sheet`
      );
    }
    for (let offset = 0; offset < entries.length; offset += 1) {
      const value = entries[offset];
      if (value !== undefined) {
        table.set(start + offset, value);
      }
    }
  }
}

/**
 * `restore-*` 的 `tables` 的三条不变量，在 prepare 期一次查完。
 *
 * `invert.ts` 的 `assertMutationIsAppliable` 把它们记成"4B 的前置条件，不是可选项"，
 * 理由是这几条一旦不成立，后果都是**静默**的：
 *
 * 1. **长度必须等于被恢复的行/列数。** 偏长的数组会顺着 `writeTables` 的写路径越过
 *    被恢复的那一段，写进 insert 往下推的那几行——那是对**没被碰过的数据**的破坏，
 *    比丢条目更糟。
 * 2. **表名必须是这张 sheet 上已登记的表。** 4A 里这条只在测试的 `writeTables` 里
 *    存在。它在 4B 变成真实的运行期风险：撤销记录会在历史栈里存活很久，其间插件
 *    完全可以注销一张表，等到重放时表就不在了。
 * 3. **那张表的 `indexedBy` 必须与被恢复的轴一致。** 行轴的 `restore` 拿着一张列表
 *    的条目去写行下标，写进去不报错，只会静默错位。
 *
 * 放在 `apply.ts` 而不是 `invert.ts`，是因为守卫的是**写路径**（`writeTables`），
 * 让检查与它守的那段代码待在同一个文件里；`invert.ts` 引用它，两处因此不会分叉。
 */
export function assertRestoreTablesAreWritable(
  model: SheetModel,
  axis: "row" | "column",
  tables: SideTableEntries,
  expectedLength: number
): void {
  for (const name of Object.keys(tables)) {
    const entries = tables[name] as readonly unknown[];
    if (entries.length !== expectedLength) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidMutation,
        `restore table "${name}" carries ${entries.length} entries, expected ${expectedLength}`
      );
    }
    const table = model.table(name);
    if (!table) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidMutation,
        `restore names table "${name}", which is not registered on this sheet`
      );
    }
    if (table.indexedBy !== axis) {
      throw new CyExcelError(
        CyExcelErrorCode.InvalidMutation,
        `restore table "${name}" is indexed by ${table.indexedBy}, cannot restore onto the ${axis} axis`
      );
    }
  }
}

/**
 * 一条变更进历史栈占多少字节的**粗估**。
 *
 * 口径与仓库既有的几处一致（`RowHeightMutation` 的 `64 + n * 32`、
 * `GridStructureOwner` 的 `192 + count * 96`）：按条目数乘常数，不做深度遍历。
 * 精度要求不高——它唯一的消费者是历史栈的内存上限，那是一个"别把一亿个格子留在
 * 内存里"的粗闸门，不是会计。**刻意不用 `JSON.stringify`**：那会把"估个大小"变成
 * 与被估内容同阶的一次遍历加分配，而这个函数在每条变更上都跑。
 *
 * 因此这几个常数不必对，只需要**同阶**：带回 100 行记录的 `restore-rows` 要比
 * `delete-rows` 贵两个数量级，这一点它做到了。
 */
export function estimateSheetMutationBytes(mutation: SheetMutation): number {
  switch (mutation.type) {
    case "set-cell":
      return BASE_BYTES + (mutation.record === undefined ? 0 : RECORD_BYTES);
    case "set-range": {
      let records = 0;
      let slots = 0;
      for (const line of mutation.records) {
        slots += line.length;
        for (const record of line) {
          if (record !== undefined) {
            records += 1;
          }
        }
      }
      // Empty entries still retain reference slots and a container for every row.
      return BASE_BYTES + records * RECORD_BYTES + mutation.records.length * 32 + slots * 8;
    }
    case "insert-rows":
      return BASE_BYTES + countRowCells(mutation.rows) * RECORD_BYTES;
    case "restore-rows":
      return (
        BASE_BYTES +
        countRowCells(mutation.rows) * RECORD_BYTES +
        countDefined(mutation.meta) * META_BYTES +
        countTableEntries(mutation.tables) * ENTRY_BYTES +
        mutation.mergesBefore.length * MERGE_BYTES +
        mutation.validationsBefore.length * VALIDATION_BYTES
      );
    case "restore-columns":
      return (
        BASE_BYTES +
        countRowCells(mutation.cells) * RECORD_BYTES +
        countDefined(mutation.meta) * META_BYTES +
        countTableEntries(mutation.tables) * ENTRY_BYTES +
        mutation.mergesBefore.length * MERGE_BYTES +
        mutation.validationsBefore.length * VALIDATION_BYTES
      );
    case "reorder-rows":
      // `Int32Array` 是这条变更里唯一随表长增长的东西，按它的实际字节数算
      return BASE_BYTES + mutation.order.length * 4;
    case "set-row-meta":
    case "set-column-meta":
      return BASE_BYTES + (mutation.meta === undefined ? 0 : META_BYTES);
    case "set-merges":
      return BASE_BYTES + mutation.merges.length * MERGE_BYTES;
    case "set-validations":
      return BASE_BYTES + mutation.validations.length * VALIDATION_BYTES;
    case "delete-rows":
    case "move-rows":
    case "insert-columns":
    case "delete-columns":
    case "move-columns":
      // 只有几个下标，不随表长增长
      return BASE_BYTES;
  }
}

/** 一条变更本身的固定开销（判别式、下标、对象头）。 */
const BASE_BYTES = 64;
/** 一条 `CellRecord`：一个小对象加上 `v`/`t`/`s`/`f` 几个字段。 */
const RECORD_BYTES = 48;
/** 一条数据验证规则：对象头加判据、区域数组，粗估。 */
const VALIDATION_BYTES = 160;
/** 一条 `RowMeta` / `ColumnMeta`：比记录少几个字段。 */
const META_BYTES = 32;
/** 一条插件侧表条目：载荷由插件定义，核心估不了，按小对象算。 */
const ENTRY_BYTES = 32;
/** 一条 `CellRange`：四个数字字段。 */
const MERGE_BYTES = 32;

function countRowCells(rows: readonly (RowCells | ColumnCells)[]): number {
  let cells = 0;
  for (const row of rows) {
    cells += Object.keys(row).length;
  }
  return cells;
}

function countDefined(meta: readonly (RowMeta | ColumnMeta | undefined)[]): number {
  let defined = 0;
  for (const entry of meta) {
    if (entry !== undefined) {
      defined += 1;
    }
  }
  return defined;
}

function countTableEntries(tables: SideTableEntries): number {
  let entries = 0;
  for (const name of Object.keys(tables)) {
    for (const value of tables[name] as readonly unknown[]) {
      if (value !== undefined) {
        entries += 1;
      }
    }
  }
  return entries;
}
