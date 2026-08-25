import { ROW_PLAN_HOLE } from "./rowPlan";

/**
 * 按行计划重键一张按行下标索引的映射（`rowData` 就是这样一张表）。
 *
 * 结构操作必须带着它一起搬：插一行之后，原第 5 行的行高若留在第 5 行，
 * 那已经是另一行了。ADR §3.1 把这条列为每张下标索引侧表的义务。
 */
export function remapByRowPlan<T>(
  source: Readonly<Record<number, T>>,
  plan: Int32Array
): Record<number, T> {
  const next: Record<number, T> = {};
  for (let newIndex = 0; newIndex < plan.length; newIndex += 1) {
    const oldIndex = plan[newIndex] as number;
    if (oldIndex === ROW_PLAN_HOLE) {
      continue;
    }
    const value = source[oldIndex];
    if (value !== undefined) {
      next[newIndex] = value;
    }
  }
  return next;
}

/** 按列变更重键。`remap` 返回 `undefined` 表示该项被删除。 */
export function remapByColumnShift<T>(
  source: Readonly<Record<number, T>>,
  remap: (index: number) => number | undefined
): Record<number, T> {
  const next: Record<number, T> = {};
  for (const key of Object.keys(source)) {
    const oldIndex = Number(key);
    const value = source[oldIndex];
    if (value === undefined) {
      continue;
    }
    const newIndex = remap(oldIndex);
    if (newIndex !== undefined) {
      next[newIndex] = value;
    }
  }
  return next;
}
