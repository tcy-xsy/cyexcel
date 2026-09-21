/**
 * 溢出计算：单行工具栏放不下时，哪些项收进"更多"下拉。
 *
 * 纯算术，不碰 DOM——宽度由调用方量好传进来。这样它在 node 里能被完整测试，而
 * `createToolbar` 那边只负责量宽度和改 class（见 step25-toolbar-plan.md §1 的分层）。
 */

export interface OverflowInput {
  /** 每一项占的宽度，顺序与 schema 渲染顺序一致 */
  readonly itemWidths: readonly number[];
  /** 工具栏可用的总宽度 */
  readonly available: number;
  /** "更多"按钮自己占的宽度——它只在真的有东西溢出时才出现 */
  readonly moreButtonWidth: number;
  /** 分组分隔线的宽度总和（它们不参与溢出，但占地方） */
  readonly separatorsWidth?: number;
}

export interface OverflowResult {
  /** 留在工具栏上的项下标 */
  readonly visible: readonly number[];
  /** 收进"更多"下拉的项下标，保持原顺序 */
  readonly overflowed: readonly number[];
}

/**
 * 从左往右能放多少放多少，放不下的全部收起。
 *
 * 一个容易搞错的地方：**"更多"按钮自己也要占宽度**，而它只在有东西溢出时才存在。
 * 所以不能先按总宽塞满、再发现装不下——那样加上"更多"按钮又会挤掉一项，来回反复。
 * 这里的做法是：先看全部项能不能放下（不需要"更多"），放不下才按
 * `available - moreButtonWidth` 重算一遍。
 */
export function computeOverflow(input: OverflowInput): OverflowResult {
  const { itemWidths, available, moreButtonWidth } = input;
  const separators = input.separatorsWidth ?? 0;
  const total = itemWidths.reduce((sum, width) => sum + width, 0);

  if (total + separators <= available) {
    return {
      visible: itemWidths.map((_unused, index) => index),
      overflowed: []
    };
  }

  const budget = available - moreButtonWidth - separators;
  const visible: number[] = [];
  const overflowed: number[] = [];
  let used = 0;
  for (let index = 0; index < itemWidths.length; index += 1) {
    const width = itemWidths[index] as number;
    // 一旦有一项放不下，后面的**全部**收起——不跳过它去塞后面更窄的项，那会让
    // 工具栏的顺序在不同宽度下变来变去，用户记不住任何按钮的位置。
    if (overflowed.length > 0 || used + width > budget) {
      overflowed.push(index);
      continue;
    }
    used += width;
    visible.push(index);
  }
  return { visible, overflowed };
}
