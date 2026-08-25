/**
 * 列轴的移位原语。行轴坍缩成一个容洞排列（ADR §3.1），列轴坍缩不到那一步——
 * 它是逐个占用列过的纯算术——但**这段算术只该存在一份**。`CellMatrix`（重键
 * 格子）与 `SheetModel`（重键 `columnData`）共用这里，spec §7.2 里按列索引的
 * 插件资源也用这里，不再手抄。
 *
 * 返回 `undefined` 表示该列被删掉了。
 */
export type ColumnShift = (column: number) => number | undefined;

export function buildInsertColumnsShift(start: number, count: number): ColumnShift {
  return (column) => (column < start ? column : column + count);
}

export function buildDeleteColumnsShift(start: number, count: number): ColumnShift {
  return (column) => {
    if (column < start) {
      return column;
    }
    if (column < start + count) {
      return undefined;
    }
    return column - count;
  };
}

export function buildMoveColumnsShift(
  start: number,
  count: number,
  target: number
): ColumnShift {
  // target 与行轴同一口径：落在**原**下标第几个之前
  const insertAt = target > start ? target - count : target;
  return (column) => {
    if (column >= start && column < start + count) {
      return insertAt + (column - start);
    }
    const withoutBlock = column > start ? column - count : column;
    return withoutBlock >= insertAt ? withoutBlock + count : withoutBlock;
  };
}
