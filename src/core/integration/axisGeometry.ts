import { geometricColumnWidthOf } from "../mutations/columnWidthMeta";
import type { SheetModel } from "../sheet/SheetModel";
import type { ColumnMeta } from "../sheet/types";

/**
 * 构造那一刻的行高/列宽从哪里来——模型与配置的合流点。
 *
 * ADR 0004 阶段 4B T8 之后，行高的权威是 `rowData[r].h`、列宽与隐藏的权威是
 * `columnData[c]` 的 `w`/`hd`，`LayoutEngine` 那边只是稠密索引。但入口选项里仍然有
 * 两处几何配置——`CyExcelCoreOptions.rowHeights` 和 `ColumnDefinition.width`——它们
 * 说的是同一件事。**构造是唯一一处两个来源同时存在的接缝**，所以合流规则要写在一处：
 *
 * 1. **模型赢**。快照里存过的 `h`/`w` 是用户上次存下来的状态，配置只是缺省。
 * 2. **配置写回模型**。模型没说话的地方，把配置里的显式值补成 `rowData`/`columnData`
 *    条目——否则这份几何就只活在布局索引里，成了权威之外的第二份存储，正是 §5.1 要
 *    删掉的那个东西。
 *
 * 这是在修一个 T8 之前就有的缺陷：`buildLayout` 把 `options.rowHeights` 喂给布局却
 * 从不写进 `rowData`，也从不读快照里的 `rowData[r].h` / `columnData[c].w`。于是
 * **快照里的行高和列宽根本不往返**——存下来的表重新装载后回到缺省值。T8 把两份存储
 * 收成一份，这条接缝不补上，收敛就只做了一半（设计稿 §6.4）。
 */

export interface InitialAxisGeometryOptions {
  readonly sheet: SheetModel;
  readonly rowCount: number;
  readonly columnCount: number;
  /** 入口选项里的 `rowHeights`，缺省时视作空。 */
  readonly configuredRowHeights?: ReadonlyMap<number, number> | undefined;
  /** 每一列声明的宽度，`undefined` 表示这一列没有显式宽度。 */
  readonly configuredColumnWidths: readonly (number | undefined)[];
  readonly defaultColumnWidth: number;
}

export interface InitialAxisGeometry {
  /** 喂给 `LayoutOptions.rowHeights` 的那张表（只含有显式高度的行）。 */
  readonly rowHeights: ReadonlyMap<number, number>;
  /** 喂给 `LayoutOptions.columnWidths` 的稠密几何宽度（隐藏列为 0）。 */
  readonly columnWidths: readonly number[];
  /** 配置提供、模型尚未记录的行高——构造完布局之后要补进 `rowData`。 */
  readonly rowSeeds: ReadonlyMap<number, number>;
  /** 同上，列宽那一半。 */
  readonly columnSeeds: ReadonlyMap<number, number>;
}

/** 只算不写：这一步可能因为配置非法而被 `LayoutEngine` 拒掉，那时模型必须一个字节都没动。 */
export function resolveInitialAxisGeometry(
  options: InitialAxisGeometryOptions
): InitialAxisGeometry {
  const { sheet, rowCount, columnCount, defaultColumnWidth } = options;
  const rowHeights = new Map<number, number>();
  const rowSeeds = new Map<number, number>();
  for (let row = 0; row < rowCount; row += 1) {
    const stored = sheet.getRowMeta(row)?.h;
    const configured = options.configuredRowHeights?.get(row);
    if (stored !== undefined) {
      rowHeights.set(row, stored);
    } else if (configured !== undefined) {
      rowHeights.set(row, configured);
      rowSeeds.set(row, configured);
    }
  }

  const columnWidths: number[] = [];
  const columnSeeds = new Map<number, number>();
  for (let column = 0; column < columnCount; column += 1) {
    const meta = sheet.getColumnMeta(column);
    const configured = options.configuredColumnWidths[column];
    if (meta?.w === undefined && configured !== undefined) {
      columnSeeds.set(column, configured);
    }
    // 隐藏位只有模型说了算（配置里没有它），几何宽度照 §4.4 的同一个函数算：隐藏即 0。
    const effective: ColumnMeta | undefined = meta?.w === undefined && configured !== undefined
      ? { ...meta, w: configured }
      : meta;
    columnWidths.push(geometricColumnWidthOf(effective, defaultColumnWidth));
  }

  return {
    rowHeights,
    columnWidths: Object.freeze(columnWidths),
    rowSeeds,
    columnSeeds
  };
}

/**
 * 把配置提供的那部分几何补进权威。
 *
 * **只在 `LayoutEngine` 构造成功之后调用**：那一步会校验键范围和数值，失败即抛，
 * 此时模型还没被碰过。
 *
 * 写完之后 `pinCoordinateSpace` 把轴钉回调用者算好的那个长度：`rowData`/`columnData`
 * 都算进 `#*ExtentIncludingTables()`，给第 N 行补一条元数据会让派生轴凭空长一格
 * （spec §3.3）。行只写 `h` 不写 `hp`——来源留在 `"default"`，自动行高才会照旧接管
 * 这些行。
 */
export function seedAxisMeta(
  sheet: SheetModel,
  geometry: Pick<InitialAxisGeometry, "rowSeeds" | "columnSeeds">,
  space: { readonly rowCount: number; readonly columnCount: number }
): void {
  if (geometry.rowSeeds.size === 0 && geometry.columnSeeds.size === 0) {
    return;
  }
  for (const [row, height] of geometry.rowSeeds) {
    sheet.setRowMeta(row, Object.freeze({ ...sheet.getRowMeta(row), h: height }));
  }
  for (const [column, width] of geometry.columnSeeds) {
    sheet.setColumnMeta(column, Object.freeze({ ...sheet.getColumnMeta(column), w: width }));
  }
  sheet.pinCoordinateSpace(space.rowCount, space.columnCount);
}
