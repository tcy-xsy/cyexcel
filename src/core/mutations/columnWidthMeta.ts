import type { ColumnMeta } from "../sheet/types";

/**
 * 列宽权威（`columnData`）与布局几何之间的那层翻译——`rowHeightMeta.ts` 的列轴对应物。
 *
 * ADR 0004 阶段 4B T8b 之后，「第 c 列有多宽、它是不是隐藏的」的答案只有一处：
 * `SheetModel.getColumnMeta(c)` 的 `w` / `hd`。`ColumnAxisSnapshot` 只留**几何宽度**
 * （隐藏即 0），它是那份稀疏权威的稠密索引，两边必须在同一个事务里一起改
 * （设计稿 §4.2/§4.3/§4.4）。
 *
 * **「隐藏」这份存储是删掉的，不是镜像过去的**（与 T8a 对 provenance 的处理同一招）：
 * 几何宽度已经把隐藏编码成 0，再留一张 `hiddenColumnWidths` 就是第二份存储，也就是
 * 第二处能分叉的地方。于是 `hideColumn` / `showColumn` 退化成两次宽度写，唯一还
 * 记得「这列本来多宽」的是权威里的 `w`。
 */

/** 这一列**自己的**宽度：隐藏时是「取消隐藏后会恢复成的那个宽度」，不是几何宽度。 */
export function columnWidthOf(meta: ColumnMeta | undefined, defaultColumnWidth: number): number {
  return meta?.w ?? defaultColumnWidth;
}

export function isColumnHiddenIn(meta: ColumnMeta | undefined): boolean {
  return meta?.hd === true;
}

/** 布局索引里存的那个数：隐藏列一律 0（`LayoutOptions.columnWidths` 沿用的约定）。 */
export function geometricColumnWidthOf(
  meta: ColumnMeta | undefined,
  defaultColumnWidth: number
): number {
  return isColumnHiddenIn(meta) ? 0 : columnWidthOf(meta, defaultColumnWidth);
}

/**
 * 把一次列宽写回合并进这一列已有的元数据。
 *
 * **合并而不是整条替换**，理由与 `withRowHeight` 相同：`set-column-meta` 载的是整条
 * `ColumnMeta`，直接写 `{ w }` 会把 `hd`（隐藏）连同 `t`/`s`/`minW`/`maxW`/`sortable`
 * 一起抹掉。改宽度不该顺手取消隐藏——今天 `prepareSetColumnWidth` 对隐藏列写的正是
 * 「恢复宽度」而几何宽度留在 0，这条语义靠这里的合并原样保住。
 *
 * **不因为「宽度恰好等于默认值」就删掉 `w`**：与行高那边不同，列宽没有 provenance
 * 这个第三态可以用来表达「这一列没有显式宽度」。要表达它只能是**整条条目没有 `w`**，
 * 而那件事只有求逆（`set-column-meta`，`meta` 可以是 `undefined`）做得到。所以这个
 * 函数永远写一个显式宽度，返回值也永远不是 `undefined`。
 */
export function withColumnWidth(meta: ColumnMeta | undefined, width: number): ColumnMeta {
  return Object.freeze({ ...meta, w: width });
}

/**
 * 置/清这一列的隐藏位。
 *
 * 取消隐藏时**删掉 `hd` 而不是写 `false`**：条目清空后返回 `undefined`，让
 * `columnData` 保持稀疏（`deleteColumnMeta` 的理由，也是 `restore-columns` 能逐字段
 * 往返的前提——写一个 `{ hd: false }` 会把这一列永久留在 `columnExtent()` 的并集里）。
 */
export function withColumnHidden(
  meta: ColumnMeta | undefined,
  hidden: boolean
): ColumnMeta | undefined {
  const next: { -readonly [K in keyof ColumnMeta]: ColumnMeta[K] } = { ...meta };
  if (hidden) {
    next.hd = true;
  } else {
    delete next.hd;
  }
  return Object.keys(next).length === 0 ? undefined : Object.freeze(next);
}
