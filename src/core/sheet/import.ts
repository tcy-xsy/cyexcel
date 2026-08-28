import { WorkbookModel } from "./WorkbookModel";
import type { RowCells } from "./mutation";
import type { SheetModelOptions } from "./SheetModel";
import type { CellRecord, WorkbookSnapshot } from "./types";

/**
 * 导入用的一行：字段名 -> 值。
 *
 * 刻意不复用 `src/core/types.ts` 的 `RowData`——那个类型是被替代的那一套的入口，
 * T7 会把它一起删掉，而导入器是**永久设施**（设计稿 §3.2：从 JSON 装载是真需求，
 * xlsx 导入将来落在同一个位置）。
 */
export type ImportRow = Readonly<Record<string, unknown>>;

export interface RowsToWorkbookSnapshotOptions extends SheetModelOptions {
  /**
   * 数据之后**额外**保留的空白行/列，默认 0。
   *
   * 注意它是"额外"而不是"全部"：导入器保证产出的坐标空间恰好是
   * `rows.length + blankRows` × `fields.length + blankColumns`。见下面的函数注释。
   */
  readonly blankRows?: number;
  readonly blankColumns?: number;
}

/**
 * 一个值变成一条记录，`undefined` 变成"没有这个格子"。
 *
 * **`undefined` 与 `null` 不同口径，这是有意的。** 稀疏矩阵里"未占用"就是"这个键
 * 不存在"（spec §11 不变量 1：不存空记录），所以只有 `undefined` 能表达"没有"。
 * `null` 是一个普普通通的值，存成 `{ v: null }`——与 `set-cell` 那条变更里
 * `record: undefined` 表示删格、`{ v: null }` 表示一条真记录是同一条口径。设计稿
 * §9 第 3 条当时倾向让 `null` 也不产生格子，这里改成不那样：让同一个值在"导入"与
 * "写入"两条路上有两种含义，是下一个静默缺陷的生成条件。
 *
 * **不带 `t`/`s`：** 导入器不猜类型也不派样式（列默认类型走 `ColumnMeta.t`，样式是
 * 阶段 5）。
 */
export function valueToRecord(value: unknown): CellRecord | undefined {
  return value === undefined ? undefined : { v: value };
}

/**
 * 一行的值数组（按列号定位）变成稀疏的 `RowCells`。
 *
 * `TableStore.createInsertRowsMutation`（插入的行按位置给值）与
 * `rowsToWorkbookSnapshot`（导入）共用这一个，是为了让"空值算不算一个格子"在两条
 * 路上只有一个答案——插入一行空行与导入一行空行必须得到同一张矩阵。
 */
export function rowValuesToCells(values: readonly unknown[]): RowCells {
  const cells: Record<number, CellRecord> = {};
  for (let column = 0; column < values.length; column += 1) {
    const record = valueToRecord(values[column]);
    if (record !== undefined) {
      cells[column] = record;
    }
  }
  return cells;
}

/**
 * 对象数组 + 字段列表 → 一份 `WorkbookSnapshot`（设计稿 §3.2）。
 *
 * ```ts
 * const snapshot = rowsToWorkbookSnapshot(rows, ["sku", "name", "price"]);
 * new CyExcelCore({ container, columns, snapshot });
 * ```
 *
 * **`fields` 是导入的参数，不是列配置的字段。** 数据长什么形状是数据自己的事实
 * （ADR §4.5 否决了拿 `field` 当数据格式开关）；第 `i` 个字段就是第 `i` 列，列配置
 * 只管宽度、类型这些纯展示属性。
 *
 * **坐标空间恰好是 `rows.length + blankRows` × `fields.length + blankColumns`。**
 * 这一条要专门保证，因为坐标空间是**派生**的（`maxOccupied + 1 + blank`）而稀疏
 * 矩阵里全空的尾行根本不占位置：`[{ a: 1 }, { a: undefined }]` 只占一行，直接交给
 * `createEmpty({ blankRows: 0 })` 得到的是一张 1 行的表，导入 2 行却少一行。所以
 * 这里把"数据末尾的空行"折算进 `blankRows`——它们**本来就是**空白行，这正是
 * `blankRows` 这个字段的含义。
 *
 * **写回没有了。** 产出的是一份快照，与调用方手里的 `rows` 数组再无联系；改表格
 * 不会改到那些对象里去。这是 ADR §3.2「调用方不能替换数据层」的直接后果，是意图
 * 之内的破坏性变更（设计稿 §3.3）。
 */
export function rowsToWorkbookSnapshot(
  rows: readonly ImportRow[],
  fields: readonly string[],
  options: RowsToWorkbookSnapshotOptions = {}
): WorkbookSnapshot {
  const { blankRows: extraBlankRows = 0, blankColumns: extraBlankColumns = 0, ...sheetOptions } = options;

  // 第一遍只找占用范围，**不留下任何中间结构**。第二遍才写格子。
  //
  // 一趟写完更直白，但那需要先知道 blank 数才能建模型，而 blank 数要等扫完才知道
  // ——于是一趟版只能先把记录攒进一个数组。那个数组与最终的矩阵同阶：100 万行 × 50
  // 列的导入会先造 5000 万个临时对象再造一遍真的，峰值内存翻倍，而这正是导入器最
  // 容易被喂大数据的地方。多扫一遍是 O(rows × fields) 的读，比多分配一份便宜得多。
  let occupiedRows = 0;
  let occupiedColumns = 0;
  for (let row = 0; row < rows.length; row += 1) {
    const source = rows[row] as ImportRow;
    for (let column = 0; column < fields.length; column += 1) {
      if (source[fields[column] as string] !== undefined) {
        occupiedRows = row + 1;
        occupiedColumns = Math.max(occupiedColumns, column + 1);
      }
    }
  }

  const workbook = WorkbookModel.createEmpty({
    sheet: {
      ...sheetOptions,
      // 见上面「坐标空间恰好是……」：尾部的空行/空列折算进 blank，它们本来就是空白。
      blankRows: Math.max(0, rows.length - occupiedRows) + extraBlankRows,
      blankColumns: Math.max(0, fields.length - occupiedColumns) + extraBlankColumns
    }
  });
  const sheet = workbook.sheet();
  for (let row = 0; row < occupiedRows; row += 1) {
    const source = rows[row] as ImportRow;
    for (let column = 0; column < occupiedColumns; column += 1) {
      const record = valueToRecord(source[fields[column] as string]);
      if (record !== undefined) {
        sheet.setRecord(row, column, record);
      }
    }
  }
  return workbook.toSnapshot();
}
