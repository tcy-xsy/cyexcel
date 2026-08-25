import { CyExcelError, CyExcelErrorCode } from "../errors";
import { invertRowOrder } from "../layout/rowOrder";
import { validateAxisChange } from "../layout/structuralChange";
import type { CellRange } from "../types";
import type { ColumnCells, RowCells, SheetMutation } from "./mutation";
import type { SheetModel } from "./SheetModel";
import type { CellRecord, ColumnMeta, RowMeta } from "./types";

/**
 * 求一条变更的逆变更。
 *
 * 契约：在**施加 `m` 之前**的模型上取 `inv = invertSheetMutation(model, m)`，
 * 然后施加 `m`、再施加 `inv`，模型逐字段回到原样（以 `toSnapshot()` 比较）。
 * `inv` 必须在施加 `m` **之前**求，因为它要读的正是 `m` 将要覆盖掉的那部分状态。
 *
 * 两条硬约束贯穿本文件：
 *
 * 1. **精确，不许近似。** 凡是当前 `SheetMutation` 表达不出来的逆，这里**拒绝**，
 *    不给一条会静默丢数据的近似——见 `assertNoPluginTablesOnAxis`。仓库先例是
 *    `SheetModel.#assertNoMerges`：合并区的正确搬移没实现之前，它拒绝而不是错着
 *    搬，理由写作"静默错搬比拒绝糟得多"。
 * 2. **非侵入。** 本文件只读模型，一个写方法都不调。理由是 `MutationCoordinator`
 *    先 prepare 全部参与者、再 commit：prepare 期间任一参与者抛错，模型必须原封
 *    不动。所以连本文件里抛错的那几条路径也必须在抛之前没有写过任何东西——这里
 *    的做法是根本不写，而不是"先写后回滚"。
 *
 *    **"非侵入"说的是不写，不是不别名。** 逆变更里的 `CellRecord` / `RowMeta` /
 *    `ColumnMeta` 是模型里那几个对象**本身**（`inv.rows[0][0] === model.getRecord(…)`
 *    按引用相等），不是拷贝。这是 `types.ts` 顶部那条模块级别名规则的直接后果——
 *    "存储层一律**浅共享**：容器拷一层、值按引用共享"——而不是这里的疏漏：逐格
 *    深拷贝正是这个数据模型要省掉的分配开销。代价是那条规则的第 1 形态必须成立：
 *    "`CellRecord` 的值不可就地改"。谁就地改了一条记录，改动会同时出现在模型、
 *    快照和已入栈的 undo 记录里。容器层（`RowCells` / `ColumnCells` / `records`
 *    的数组）是这里新建的，不与模型共享。
 */
export function invertSheetMutation(model: SheetModel, mutation: SheetMutation): SheetMutation {
  assertAxisChangeIsValid(model, mutation);
  switch (mutation.type) {
    case "set-cell":
      // 逆是"把这一格恢复成现在的样子"。现在是空格时 `record` 就是 `undefined`，
      // 那条逆施加下去是删格——不是"写一条空记录"（spec §11 不变量 1：
      // "未占用的格子不出现在 `cellData` 里，不存空记录。稀疏性依赖这一点。"）。
      return {
        type: "set-cell",
        row: mutation.row,
        column: mutation.column,
        record: model.getRecord(mutation.row, mutation.column)
      };

    case "set-range":
      return { type: "set-range", range: mutation.range, records: captureRange(model, mutation.range) };

    case "insert-rows":
      // 插入的是 `rows.length` 行，逆就是把这几行原样删掉。它们是新行，删掉不丢
      // 任何原有数据，所以逆不需要携带记录。
      return { type: "delete-rows", start: mutation.start, count: mutation.rows.length };

    case "delete-rows": {
      assertNoPluginTablesOnAxis(model, "row", "delete-rows");
      const rows: RowCells[] = [];
      const meta: (RowMeta | undefined)[] = [];
      // 抓满 `count` 行，而不是只抓落在占用范围内的那几行：删除范围可以整段或部分
      // 落在空白区（spec §3.3 的坐标空间比占用范围大），那几行抓出来是空的 `{}`，
      // 但个数要对得上——逆的逆是 `delete-rows(start, rows.length)`，少抓一行，
      // 逆的逆就变成了另一条变更。
      for (let offset = 0; offset < mutation.count; offset += 1) {
        const row = mutation.start + offset;
        rows.push(captureRowCells(model, row));
        meta.push(model.getRowMeta(row));
      }
      return { type: "restore-rows", start: mutation.start, rows, meta };
    }

    case "restore-rows":
      // 与 `insert-rows` 同一条理由：`restore-rows` 也是"凭空多出 rows.length 行"。
      return { type: "delete-rows", start: mutation.start, count: mutation.rows.length };

    case "move-rows": {
      const { start, count, target } = mutation;
      return { type: "move-rows", start: movedBlockLandsAt(start, count, target), count, target: inverseMoveTarget(start, count, target) };
    }

    case "reorder-rows":
      // 约定是 `order[newPosition] = oldPosition`（ADR 0003，`rowOrder.ts` 顶部把
      // 它写成断言而非文档）。逆排列的算术只该存在一份，复用 `invertRowOrder`。
      return { type: "reorder-rows", order: invertRowOrder(mutation.order) };

    case "insert-columns":
      return { type: "delete-columns", start: mutation.start, count: mutation.count };

    case "delete-columns": {
      assertNoPluginTablesOnAxis(model, "column", "delete-columns");
      const cells: ColumnCells[] = [];
      const meta: (ColumnMeta | undefined)[] = [];
      for (let offset = 0; offset < mutation.count; offset += 1) {
        const column = mutation.start + offset;
        // `ColumnCells` 的键是**行号**，是 `RowCells` 的转置。两个类型在任务二被
        // 特意分开命名，就是为了这一行不会写反——写反了不报错，只是把每一格恢复
        // 到错误的位置上。
        cells.push(captureColumnCells(model, column));
        meta.push(model.getColumnMeta(column));
      }
      return { type: "restore-columns", start: mutation.start, count: mutation.count, cells, meta };
    }

    case "restore-columns":
      return { type: "delete-columns", start: mutation.start, count: mutation.count };

    case "move-columns": {
      const { start, count, target } = mutation;
      return { type: "move-columns", start: movedBlockLandsAt(start, count, target), count, target: inverseMoveTarget(start, count, target) };
    }
  }
}

/**
 * 参数校验前置到求逆这一步。
 *
 * **为什么不等 `SheetModel` 自己拦。** `MutationCoordinator` 的前提是"凡是能查出来
 * 的都在 prepare 期间查出来，commit 因此不会失败"。求逆属于 prepare：它是在还没动
 * 模型的时候算出撤销所需的那条变更。一条 `move-rows(1, 2, 2)`（target 落在块内）
 * 若不在这里拦，`movedBlockLandsAt` 会照样算出一个数、返回一条形状合法但语义是垃圾
 * 的逆，等到 commit 期间 `SheetModel.moveRows` 才抛——而那时这条垃圾逆已经进了撤销
 * 记录，且 commit 本该是不会失败的那一步。
 *
 * 复用 `../layout/structuralChange` 的 `validateAxisChange`——与 `SheetModel` 的六个
 * 结构方法是同一份校验、同一种失败方式（`RangeError`），不另写一份会分叉的。轴长按
 * 它那里的口径取 `rowCount()` / `columnCount()`（坐标空间，spec §3.3），不是占用范围。
 *
 * **三个成员有意不在这里校验：**
 * - `restore-rows` / `restore-columns`：只作为逆变更出现，由本文件从一条已经校验过
 *   的 delete 产出，不接受外部构造。
 * - `reorder-rows`：它要的是"是不是一个排列"（`assertValidRowOrder`）与 extent 检查，
 *   `validateAxisChange` 表达不了；补它是另一个决定，留给 4B 接线时一并做。
 */
function assertAxisChangeIsValid(model: SheetModel, mutation: SheetMutation): void {
  switch (mutation.type) {
    case "insert-rows":
      validateAxisChange(
        { axis: "row", operation: "insert", start: mutation.start, count: mutation.rows.length },
        model.rowCount()
      );
      return;
    case "delete-rows":
      validateAxisChange(
        { axis: "row", operation: "delete", start: mutation.start, count: mutation.count },
        model.rowCount()
      );
      return;
    case "move-rows":
      validateAxisChange(
        {
          axis: "row",
          operation: "move",
          start: mutation.start,
          count: mutation.count,
          target: mutation.target,
          targetSpace: "pre-removal"
        },
        model.rowCount()
      );
      return;
    case "insert-columns":
      validateAxisChange(
        { axis: "column", operation: "insert", start: mutation.start, count: mutation.count },
        model.columnCount()
      );
      return;
    case "delete-columns":
      validateAxisChange(
        { axis: "column", operation: "delete", start: mutation.start, count: mutation.count },
        model.columnCount()
      );
      return;
    case "move-columns":
      validateAxisChange(
        {
          axis: "column",
          operation: "move",
          start: mutation.start,
          count: mutation.count,
          target: mutation.target,
          targetSpace: "pre-removal"
        },
        model.columnCount()
      );
      return;
    default:
      return;
  }
}

/**
 * ## `move` 的逆——推导，不是背下来的
 *
 * 仓库口径（`structuralChange.ts` 的 `AxisChangeDescriptor.targetSpace: "pre-removal"`）：
 * `target` 是**原**下标空间里的落点，即"这块落在原下标第几个之前"。决定性证据是
 * `tests/unit/core/data/ArrayDataSource.test.ts:207`——`moveRows(1, 2, 4)`，注释
 * "move B,C to just before E"，`A,B,C,D,E` 得到 `A,D,B,C,E`。
 *
 * 两处实现（`buildMoveRowsPlan`、`buildMoveColumnsShift`）共用同一行算术：
 *
 * ```
 * insertAt = target > start ? target - count : target
 * ```
 *
 * 它们都是"先把块摘出去得到余下序列 `rest`，再把块插回 `rest` 的第 `insertAt` 位"。
 * `insertAt` 因此有两个身份：它既是**摘块之后的插入位**，也正是**块在移动后所处的
 * 起始下标**——因为插入位之前的那些元素一个没动。
 *
 * 逆要做的事：把块从 `insertAt` 搬回去。搬回去意味着，在同一个 `rest`（摘块只跟块
 * 的位置有关，与块搬到哪儿无关，所以正逆两次摘出来的 `rest` 是同一个序列）里，把块
 * 重新插在第 `start` 位——因为原序列本来就是 `rest` 在第 `start` 位插入该块。
 *
 * 于是逆的入参是 `move(insertAt, count, target')`，且要求它算出的
 * `insertAt' = target' > insertAt ? target' - count : target'` 等于 `start`。分两支：
 *
 * - `insertAt < start`（块被**前**移）：取 `target' = start + count`。
 *   此时 `target' > insertAt` 成立（`start + count > start > insertAt`），
 *   走减法支，`insertAt' = start + count - count = start` ✓
 * - `insertAt >= start`（块被**后**移，或压根没动）：取 `target' = start`。
 *   此时 `target' > insertAt` 不成立，走直取支，`insertAt' = start` ✓
 *
 * 两条手算校验（都用上面那份 `A,B,C,D,E`）：
 *
 * - `move(1, 2, 4)` → `insertAt = 2`，`2 >= 1` 取后一支 → 逆是 `move(2, 2, 1)`。
 *   在 `A,D,B,C,E` 上摘出 `B,C` 得 `A,D,E`，`insertAt' = 1`，插回得 `A,B,C,D,E` ✓
 * - `move(3, 1, 0)` → `insertAt = 0`，`0 < 3` 取前一支 → 逆是 `move(0, 1, 4)`。
 *   在 `D,A,B,C,E` 上摘出 `D` 得 `A,B,C,E`，`insertAt' = 4 - 1 = 3`，
 *   插回得 `A,B,C,D,E` ✓
 *
 * 这两个逆也都过得了 `validateAxisChange`：`target'` 不会落进块内。
 * 前一支要落进块内需要 `start + count < insertAt + count`，即 `start < insertAt`，
 * 与该支的前提 `insertAt < start` 矛盾；后一支要落进块内需要 `start > insertAt`，
 * 与该支的前提矛盾。
 */
function movedBlockLandsAt(start: number, count: number, target: number): number {
  return target > start ? target - count : target;
}

function inverseMoveTarget(start: number, count: number, target: number): number {
  return movedBlockLandsAt(start, count, target) < start ? start + count : start;
}

/**
 * 删除的逆当前只带得回**格子 + `rowData` / `columnData`**，带不回经
 * `SheetModel.registerTable()` 登记的其它按下标索引的侧表——`restore-rows` /
 * `restore-columns` 没有装它们的位置（任务二的类型如此）。
 *
 * 而 spec §11 不变量 8 写的是"声明 `indexedBy: "row" | "column"` 的资源，其重键与
 * **逆操作**由核心提供并纳入同一 transaction，不外包给插件作者（§7.2）"。删除会把
 * 落在被删范围内的条目从这些表里搬走（`IndexedTableRegistry.applyRowPlan`），逆变更
 * 却没有任何字段能把它们带回来——重键做到了，逆操作没有。
 *
 * 这个缺口**测不出来**：`SheetModel.toSnapshot()` 只序列化 `rowData` / `columnData`
 * （见该方法的注释），插件资源的持久化形态是 workbook 级的 `resources`。所以逐字段
 * 比较快照的精确性测试对它是瞎的。
 *
 * 因此这里**拒绝**。给一条丢数据的逆等于让 undo 静默吃掉插件数据，而这正是
 * ADR §3.1 那条被 spec §7.2 称作"不变量第 4 条"的要求要防的：
 * "每增加一张按 `(row, column)` 索引的侧表，就要为这 7 个操作各写一遍重键逻辑、
 * 各写一遍逆操作，并全部纳入同一个 transaction。"补齐的办法是给 `restore-*` 加一个
 * 装侧表条目的字段——那是类型层面的改动，不在本任务范围内。
 */
function assertNoPluginTablesOnAxis(
  model: SheetModel,
  axis: "row" | "column",
  operation: string
): void {
  const coreTable = axis === "row" ? "rowData" : "columnData";
  for (const name of model.tableNames()) {
    if (name === coreTable) {
      continue;
    }
    if (model.table(name)?.indexedBy !== axis) {
      continue;
    }
    throw new CyExcelError(
      CyExcelErrorCode.MutationNotSupported,
      `${operation} cannot be inverted exactly: the registered ${axis}-indexed table "${name}" has no slot in restore-${axis}s`
    );
  }
}

/**
 * 抓一行的格子（列号 -> 记录，稀疏）。
 *
 * **为什么按列号数着扫，而不是走行桶。** `SheetModel` 今天没有对外暴露"按行桶迭代"
 * 的读接口，`CellMatrix.rows()` 是内部的。这里的扫描以
 * `occupiedColumnCount()`（最大占用列号加一）为界，**不是**以 `columnCount()`
 * （坐标空间，含调用方配的、无上限的 `blankColumns`）为界，所以它不是 ADR §3.1
 * 点名的那个"列操作退化成 `rowCount × columnCount` 的稠密扫描"——代价随实际内容
 * 走，不随空白区走。要把它降到与占用格数同阶，得给 `SheetModel` 补一个稀疏读
 * 接口；那属于 4B 接线时一并决定，本任务不擅自扩 `SheetModel` 的公开面。
 */
function captureRowCells(model: SheetModel, row: number): RowCells {
  const cells: Record<number, CellRecord> = {};
  const columns = model.occupiedColumnCount();
  for (let column = 0; column < columns; column += 1) {
    const record = model.getRecord(row, column);
    if (record !== undefined) {
      cells[column] = record;
    }
  }
  return cells;
}

/** 抓一列的格子——键是**行号**，与 `captureRowCells` 互为转置，别写反。 */
function captureColumnCells(model: SheetModel, column: number): ColumnCells {
  const cells: Record<number, CellRecord> = {};
  const rows = model.occupiedRowCount();
  for (let row = 0; row < rows; row += 1) {
    const record = model.getRecord(row, column);
    if (record !== undefined) {
      cells[row] = record;
    }
  }
  return cells;
}

/**
 * 抓一个区域的记录矩阵。
 *
 * 尺寸取自 `range` 而不是正变更的 `records`：逆必须覆盖 `set-range` 写过的**整个**
 * 区域，而 `records` 可能是不齐的（缺的那几格按空格处理）。按 `records` 的尺寸抓，
 * 缺的那几格就恢复不回来。
 */
function captureRange(
  model: SheetModel,
  range: CellRange
): readonly (readonly (CellRecord | undefined)[])[] {
  const records: (CellRecord | undefined)[][] = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    const line: (CellRecord | undefined)[] = [];
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      line.push(model.getRecord(row, column));
    }
    records.push(line);
  }
  return records;
}
