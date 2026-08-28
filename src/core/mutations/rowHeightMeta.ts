import type { RowHeightProvenance } from "../layout";
import type { RowMeta } from "../sheet/types";

/**
 * 行高权威（`rowData`）与布局词汇（`RowHeightProvenance`）之间的那层翻译。
 *
 * ADR 0004 阶段 4B T8 之后，「第 r 行有多高、这个高度从哪来」的答案只有一处：
 * `SheetModel.getRowMeta(r)` 的 `h` / `hp`。布局层的 `RowMetrics` 是它的稠密索引，
 * 两边必须在同一个事务里一起改（设计稿 §4.2 / §4.3）。改行高的两条变更
 * （`RowHeightMutation` 与 `LayoutRootMutation` 的 `set-row-height`）都要做同样的
 * 读与合并，所以放这里一份，而不是各写各的——各写各的正是分叉的开始。
 */

export function rowHeightOf(meta: RowMeta | undefined, defaultRowHeight: number): number {
  return meta?.h ?? defaultRowHeight;
}

export function rowHeightProvenanceOf(meta: RowMeta | undefined): RowHeightProvenance {
  return meta?.hp ?? "default";
}

/**
 * 把一次行高写回合并进这一行已有的元数据，得到 `set-row-meta` 的载荷。
 *
 * 两个刻意的选择：
 *
 * 1. **合并而不是整条替换。** `set-row-meta` 载的是整条 `RowMeta`，直接写
 *    `{ h, hp }` 会把这一行的 `hd`（隐藏）抹掉——改高度不该顺手取消隐藏。
 * 2. **`provenance === "default"` 时删掉 `h`/`hp`，不写一个等于默认行高的数。**
 *    「没有显式行高」与「显式设成了恰好等于默认值的高度」在模型里是两件事：前者
 *    跟着默认值走，后者不跟。撤销一次「把默认高的行改高」必须回到前者，否则改一
 *    次默认行高，这一行就永远留在旧默认值上。条目被清空时返回 `undefined`，让
 *    `rowData` 保持稀疏（`deleteRowMeta` 的理由）。
 */
export function withRowHeight(
  meta: RowMeta | undefined,
  height: number,
  provenance: RowHeightProvenance
): RowMeta | undefined {
  const next: { h?: number; hp?: Exclude<RowHeightProvenance, "default">; hd?: boolean } = { ...meta };
  if (provenance === "default") {
    delete next.h;
    delete next.hp;
  } else {
    next.h = height;
    next.hp = provenance;
  }
  return Object.keys(next).length === 0 ? undefined : Object.freeze(next);
}
