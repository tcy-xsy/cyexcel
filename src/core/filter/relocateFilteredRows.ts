/**
 * 结构操作（插入/删除/移动行）对 `filteredRows` 集合的搬迁——纯函数，不碰
 * `LayoutEngine`/`RowMetrics` 本身，只算下标怎么变（ADR 0003 §3.2/§3.3；风格
 * 参照 `../sheet/merges.ts` 的合并区搬迁函数）。
 *
 * 三个函数都在"这次操作根本没碰到任何被过滤行"时原样返回入参集合，让调用方能用
 * `===` 判断有没有必要重建 `visibleRowMetrics`。
 */

export function relocateFilteredRowsForInsert(
  filteredRows: ReadonlySet<number>,
  start: number,
  count: number
): ReadonlySet<number> {
  if (filteredRows.size === 0 || count === 0) {
    return filteredRows;
  }
  let changed = false;
  const next = new Set<number>();
  for (const row of filteredRows) {
    if (row >= start) {
      next.add(row + count);
      changed = true;
    } else {
      next.add(row);
    }
  }
  return changed ? next : filteredRows;
}

export function relocateFilteredRowsForDelete(
  filteredRows: ReadonlySet<number>,
  start: number,
  count: number
): ReadonlySet<number> {
  if (filteredRows.size === 0 || count === 0) {
    return filteredRows;
  }
  let changed = false;
  const next = new Set<number>();
  for (const row of filteredRows) {
    if (row >= start && row < start + count) {
      // Deleted along with its row — it no longer exists.
      changed = true;
      continue;
    }
    if (row >= start + count) {
      next.add(row - count);
      changed = true;
    } else {
      next.add(row);
    }
  }
  return changed ? next : filteredRows;
}

/**
 * Mirrors `RowMetrics.move()`'s own index algebra exactly: the moved block
 * keeps its filtered rows filtered at their new position; everything else
 * shifts only if the moved block was spliced in ahead of it.
 */
export function relocateFilteredRowsForMove(
  filteredRows: ReadonlySet<number>,
  start: number,
  count: number,
  target: number
): ReadonlySet<number> {
  if (filteredRows.size === 0 || count === 0) {
    return filteredRows;
  }
  const adjustedTarget = target > start ? target - count : target;
  let changed = false;
  const next = new Set<number>();
  for (const row of filteredRows) {
    let mapped: number;
    if (row >= start && row < start + count) {
      mapped = adjustedTarget + (row - start);
    } else {
      const positionAfterRemoval = row < start ? row : row - count;
      mapped = positionAfterRemoval < adjustedTarget ? positionAfterRemoval : positionAfterRemoval + count;
    }
    if (mapped !== row) {
      changed = true;
    }
    next.add(mapped);
  }
  return changed ? next : filteredRows;
}
