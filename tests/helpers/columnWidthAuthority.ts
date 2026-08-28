import {
  resolveInitialAxisGeometry,
  seedAxisMeta
} from "../../src/core/integration";
import { DEFAULT_COLUMN_WIDTH, LayoutEngine, type LayoutOptions } from "../../src/core/layout";
import type { SheetModel } from "../../src/core/sheet/SheetModel";

/**
 * 列宽与「这列是不是隐藏的」的**权威**读法，以及一个不会一出生就违反不变量的夹具。
 *
 * ADR 0004 阶段 4B T8b 之后 `LayoutEngine` 不再存「隐藏」和「恢复宽度」，
 * `getColumnRestoreWidth`/`isColumnHidden` 也随之删掉；权威是 `columnData[c]` 的
 * `w`/`hd`，`ColumnAxisSnapshot` 只留由它派生的几何宽度（设计稿 §4.2/§4.4）。
 */

export function columnWidthOf(sheet: SheetModel, column: number, defaultColumnWidth: number): number {
  return sheet.getColumnMeta(column)?.w ?? defaultColumnWidth;
}

export function isColumnHidden(sheet: SheetModel, column: number): boolean {
  return sheet.getColumnMeta(column)?.hd === true;
}

/**
 * 索引里那个数**应该**是多少：隐藏即 0，否则是显式宽度、再否则是缺省宽。
 *
 * 故意不复用 `src/core/mutations/columnWidthMeta.ts` 里的同名函数——那样属性测试就成了
 * 拿生产代码验生产代码。这里把 §4.2 的列轴不变量用测试自己的话重写一遍，两边对不上时
 * 才是真的对不上。
 */
export function geometricColumnWidthOf(
  sheet: SheetModel,
  column: number,
  defaultColumnWidth: number
): number {
  const meta = sheet.getColumnMeta(column);
  if (meta?.hd === true) {
    return 0;
  }
  return meta?.w ?? defaultColumnWidth;
}

/**
 * 建一个与模型对得上的 `LayoutEngine`——测试里 `new LayoutEngine({ columnWidths })` 的替身。
 *
 * 直接 `new LayoutEngine({ columnWidths: [50, 60, 70] })` 而模型里没有 `columnData`，
 * 造出来的夹具**当场就违反 §4.2 的列轴不变量**：索引说第 1 列 60 宽，权威说它没有显式
 * 宽度、该按缺省算。这类夹具下「改宽再撤销」会恢复到缺省宽而不是 60，测出来的失败与
 * 生产无关。所以走生产同一条接缝（`resolveInitialAxisGeometry` + `seedAxisMeta`，
 * 即 `CyExcelCore.buildLayout` 用的那两个函数），把配置里的宽度补进权威。
 */
export function createSeededLayout(
  sheet: SheetModel,
  options: Omit<LayoutOptions, "columnWidths"> & {
    readonly columnWidths: readonly (number | undefined)[];
  }
): LayoutEngine {
  const { columnWidths, ...rest } = options;
  const defaultColumnWidth = options.defaultColumnWidth ?? DEFAULT_COLUMN_WIDTH;
  // 钉回**模型自己**的轴长而不是布局那两个数：夹具常把布局配得比数据大（空表配 5 行），
  // 拿布局的数去钉会把模型的坐标空间一起撑开，那是这个辅助函数不该有的副作用。
  const space = { rowCount: sheet.rowCount(), columnCount: sheet.columnCount() };
  const geometry = resolveInitialAxisGeometry({
    sheet,
    rowCount: options.rowCount,
    columnCount: options.columnCount,
    ...(options.rowHeights !== undefined ? { configuredRowHeights: options.rowHeights } : {}),
    configuredColumnWidths: columnWidths,
    defaultColumnWidth
  });
  const layout = new LayoutEngine({
    ...rest,
    ...(geometry.rowHeights.size > 0 ? { rowHeights: geometry.rowHeights } : {}),
    columnWidths: geometry.columnWidths,
    defaultColumnWidth
  });
  seedAxisMeta(sheet, geometry, space);
  return layout;
}
