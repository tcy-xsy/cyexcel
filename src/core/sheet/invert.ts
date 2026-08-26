import { CyExcelError, CyExcelErrorCode } from "../errors";
import { assertValidRowOrder, invertRowOrder } from "../layout/rowOrder";
import { validateAxisChange } from "../layout/structuralChange";
import type { CellRange } from "../types";
import type { ColumnCells, RowCells, SheetMutation, SideTableEntries } from "./mutation";
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
 *    不给一条会静默丢数据的近似。仓库先例是 `SheetModel.#assertNoMerges`：合并区
 *    的正确搬移没实现之前，它拒绝而不是错着搬，理由写作"静默错搬比拒绝糟得多"。
 *    删除对插件侧表曾经是这样一条拒绝（任务三），任务四给 `restore-*` 加上
 *    `tables` 字段之后改成了**捕获并带回**——见 `capturePluginTables`。
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
  assertMutationIsAppliable(model, mutation);
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
      return {
        type: "restore-rows",
        start: mutation.start,
        rows,
        meta,
        tables: capturePluginTables(model, "row", mutation.start, mutation.count)
      };
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
      return {
        type: "restore-columns",
        start: mutation.start,
        count: mutation.count,
        cells,
        meta,
        tables: capturePluginTables(model, "column", mutation.start, mutation.count)
      };
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
 * 六个轴向成员复用 `../layout/structuralChange` 的 `validateAxisChange`——与
 * `SheetModel` 的六个结构方法是同一份校验、同一种失败方式（`RangeError`），不另写
 * 一份会分叉的。轴长按它那里的口径取 `rowCount()` / `columnCount()`（坐标空间，
 * spec §3.3），不是占用范围。
 *
 * `reorder-rows` 套不进 `validateAxisChange`（它要判的是"是不是一个排列"，不是
 * start/count/target），所以逐条对着 `SheetModel.reorderRows` 的 commit 期检查复制
 * **顺序**、复用**同一批函数**——见下面那一支的注释。
 *
 * 合并区这条**模型状态**前置条件，九个结构成员一个不落地查，见 `assertNoMerges`。
 *
 * **有意不在这里校验的，只剩一条：** `restore-rows` / `restore-columns` 的**参数**。
 * 这两个成员只作为逆变更出现（`mutation.ts` 的类型注释如此），由本文件从一条已经过
 * 上面校验的 delete 产出，不接受外部构造，所以它们的 `start` / 行列数按构造就是合法
 * 的。**4B 若开始从外部接收它们，这条前提就没了，这里要补上 insert 形状的校验**——
 * 在那之前，一条参数就非法、同时又撞上合并区的 `restore-*` 会在这里先报合并区，
 * 而 commit 期先报参数；对按构造合法的那些（也就是全部）两边一致。
 */
function assertMutationIsAppliable(model: SheetModel, mutation: SheetMutation): void {
  switch (mutation.type) {
    case "reorder-rows": {
      // 与 `SheetModel.reorderRows` **逐条对齐**，包括先后顺序：那里先查 extent、
      // 再校验排列、最后查合并区，所以这里也是这个顺序——顺序反了，一个两种都
      // 不合法的 order 会在 prepare 与 commit 得到不同的错误。
      //
      // extent 取 `model.rowExtent()`，那是 `#rowExtentIncludingTables()` 的公开
      // 出口，与 `reorderRows` 读的是同一个方法而不是它的一份复制；排列检查直接
      // 复用 `assertValidRowOrder`，与 `SheetModel.reorderRows` 调的是同一个函数。
      // 两处因此不可能对"什么是合法 order"给出两个答案。
      const extent = model.rowExtent();
      if (extent > mutation.order.length) {
        throw new CyExcelError(
          CyExcelErrorCode.InvalidRange,
          `occupied row ${extent - 1} >= order length ${mutation.order.length}, would be silently dropped`
        );
      }
      assertValidRowOrder(mutation.order, mutation.order.length);
      assertNoMerges(model, "reorderRows");
      return;
    }
    case "insert-rows":
      validateAxisChange(
        { axis: "row", operation: "insert", start: mutation.start, count: mutation.rows.length },
        model.rowCount()
      );
      assertNoMerges(model, "insertRows");
      return;
    case "delete-rows":
      validateAxisChange(
        { axis: "row", operation: "delete", start: mutation.start, count: mutation.count },
        model.rowCount()
      );
      assertNoMerges(model, "deleteRows");
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
      assertNoMerges(model, "moveRows");
      return;
    case "insert-columns":
      validateAxisChange(
        { axis: "column", operation: "insert", start: mutation.start, count: mutation.count },
        model.columnCount()
      );
      assertNoMerges(model, "insertColumns");
      return;
    case "delete-columns":
      validateAxisChange(
        { axis: "column", operation: "delete", start: mutation.start, count: mutation.count },
        model.columnCount()
      );
      assertNoMerges(model, "deleteColumns");
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
      assertNoMerges(model, "moveColumns");
      return;
    // `restore-*` 的参数不在这里校验（见上），但合并区这一条要查：施加它们走的是
    // `insertRows` / `insertColumns`，合并区存在时正是那两个方法在 commit 期拒绝。
    case "restore-rows":
      assertNoMerges(model, "insertRows");
      return;
    case "restore-columns":
      assertNoMerges(model, "insertColumns");
      return;
    // 格子变更：`setRecord` / `deleteCell` 在 commit 期不查合并区，这里也不查。
    case "set-cell":
    case "set-range":
      return;
  }
}

/**
 * 合并区存在时，九个结构操作在 commit 期全都会被 `SheetModel.#assertNoMerges` 拒掉
 * （"合并区跟着结构操作正确搬移是一整块 Excel 语义……属于阶段 5。在那之前**拒绝**
 * 而不是错着搬"）。这里把同一条前置条件提到 prepare 期。
 *
 * **为什么这是安全的等价，而不是一次抢跑。** `SheetMutation` 没有任何改合并区的
 * 成员，所以合并区在一个 transaction 之内不会变——prepare 期看到的合并区状态与
 * commit 期必然相同，提前拒绝与到时候拒绝是同一个判断。
 *
 * **错误码与消息逐字照抄 `#assertNoMerges`。** 那个方法是私有的，这里只能复制这行
 * 消息；复制就会分叉，所以测试直接比对两处抛出的 constructor 与 message 是否相等，
 * 把这份复制钉住（`occupied row … would be silently dropped` 那条同理）。
 *
 * **调用位置逐个跟着 `SheetModel` 走。** 七个方法现在口径一致：先校验参数、后查
 * 合并区，理由是六个轴向方法共用的那条注释"① 放在 `#assertNoMerges` 之前——参数
 * 非法与有没有合并区无关，先报 `MutationNotSupported` 会误导调用方"。任务三交付时
 * `reorderRows` 是唯一的例外（它先查合并区），这里当时只能逐成员镜像那个例外；
 * 任务四把 `SheetModel.reorderRows` 拉齐到那条理由上，两边一起改——**改的是
 * `SheetModel`，不是让求逆这一侧偷偷与它分叉**，prepare 与 commit 仍然对同一个输入
 * 给同一个错误。`restore-*` 走的是 `insertRows` / `insertColumns`，跟着那两个走。
 */
function assertNoMerges(model: SheetModel, operation: string): void {
  if (model.merges().length > 0) {
    throw new CyExcelError(
      CyExcelErrorCode.MutationNotSupported,
      `${operation} is not supported while merged ranges exist`
    );
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
 * 抓住被删的那几行/列在**每一张插件侧表**里的条目，装进 `restore-*` 的 `tables`。
 *
 * **为什么这是核心的义务而不是插件作者的。** spec §11 不变量 8：声明
 * `indexedBy: "row" | "column"` 的资源，其重键与**逆操作**由核心提供并纳入同一
 * transaction，不外包给插件作者（§7.2）。重键那一半早就在
 * （`IndexedTableRegistry.applyRowPlan` / `applyColumnShift`），逆操作这一半在任务四
 * 之前没有落脚点：`restore-*` 没有装它们的字段，于是任务三只能**拒绝**求逆——那让
 * "注册了插件行表就不能删行"成立，到接线那一刻就是对 §7.2 的功能倒退，不是保守。
 *
 * **删除对这些条目的语义是丢弃，不是搬去别处**——这一条是照着实现读出来的，不是
 * 顺着描述假定的：行轴走 `remapByRowPlan`，它按行计划**逐位重建**一个新映射，被删
 * 的下标根本不出现在计划里，于是不会被写进新映射；列轴走 `remapByColumnShift`，
 * 被删列的 `remap` 返回 `undefined`，那一条直接跳过。两条路都是**原地消失**。所以
 * 逆变更要做的是把原值按原下标写回去，不是把它从别处搬回来。
 *
 * **核心自己的 `rowData` / `columnData` 不进 `tables`**：它们承载的行/列元数据已经
 * 由 `restore-*` 的 `meta` 字段带走，重复携带会在恢复时对同一条数据写两遍。
 *
 * **按表名装，不按顺序装。** 登记处是一个 `Map`，`tableNames()` 给的是登记顺序；
 * 而恢复发生在另一次调用里，那时表的集合可能已经变了（插件完全可以后登记一张新
 * 表）。按位置对齐的话，多一张表就整体串位——所以键是**表名**，位置只在同一张表
 * 内部表示行/列偏移。
 *
 * 轴向匹配的表**一律**出现在结果里，哪怕它在被删区间里一条条目都没有（那时数组全是
 * `undefined`）。"出现与否"不带信息，省得消费者去分辨"这张表本来就没条目"和"求逆
 * 漏了这张表"。
 */
function capturePluginTables(
  model: SheetModel,
  axis: "row" | "column",
  start: number,
  count: number
): SideTableEntries {
  const coreTable = axis === "row" ? "rowData" : "columnData";
  const tables: Record<string, readonly unknown[]> = {};
  for (const name of model.tableNames()) {
    if (name === coreTable) {
      continue;
    }
    const table = model.table(name);
    if (table?.indexedBy !== axis) {
      continue;
    }
    const entries: unknown[] = [];
    // 与 `captureRowCells` 同一条口径：抓满 `count` 项，含落在空白区的那几项（抓出来
    // 是 `undefined`），这样第 i 项永远对应 `start + i`。少抓一项，整个数组就与
    // `rows` / `count` 错位了。
    for (let offset = 0; offset < count; offset += 1) {
      entries.push(table.get(start + offset));
    }
    tables[name] = entries;
  }
  return tables;
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
