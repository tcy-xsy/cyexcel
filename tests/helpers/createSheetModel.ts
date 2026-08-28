import type { ImportRow, RowsToWorkbookSnapshotOptions } from "../../src/core/sheet/import";
import { rowsToWorkbookSnapshot } from "../../src/core/sheet/import";
import type { SheetModel } from "../../src/core/sheet/SheetModel";
import { WorkbookModel } from "../../src/core/sheet/WorkbookModel";

/**
 * 测试用的表模型工厂：从行对象数组导入一张 `SheetModel`。
 *
 * 走的是生产同一个导入器（`rowsToWorkbookSnapshot`），不是测试专用的另一条
 * 路径——`ArrayDataSource` 时代那种"测试构造得跟生产不一样"的坑，正是
 * ADR 0004 §3.2 要消掉的东西。行列数因此有保证：
 * `rowCount === rows.length`、`columnCount === fields.length`。
 */
export function createSheetFromRows(
  rows: readonly ImportRow[],
  fields: readonly string[],
  options: RowsToWorkbookSnapshotOptions = {}
): SheetModel {
  return WorkbookModel.fromSnapshot(rowsToWorkbookSnapshot(rows, fields, options)).sheet();
}
