import type { IndexRange } from "./LayoutEngine";

export interface AxisSplit {
  /** The part of `range` that falls inside the frozen band ([0, frozenCount)); `count` is 0 if none does. */
  readonly frozen: IndexRange;
  /** The part of `range` at or past the frozen boundary; `count` is 0 if `range` is entirely frozen. */
  readonly scrollable: IndexRange;
}

/**
 * Splits a single-axis index range into its frozen and scrollable parts.
 * Shared by ContentLayerRenderer, SelectionLayerRenderer, DOMCellManager and
 * the header renderers so all of them treat "frozen row/column band" the
 * same way — see doc/architecture.md §6.1. Frozen rows/columns are always
 * shown regardless of scroll position; this function only tells you which
 * part of an arbitrary range (e.g. a dirty patch range) falls in which
 * band, not what's currently visible — combine with
 * LayoutEngine.getVisibleRowRange()/getVisibleColumnRange() for that.
 */
export function splitByFrozenBoundary(range: IndexRange, frozenCount: number): AxisSplit {
  const rangeEnd = range.start + range.count;

  const frozenEnd = Math.min(rangeEnd, frozenCount);
  const frozenStart = Math.min(range.start, frozenEnd);
  const frozen: IndexRange = { start: frozenStart, count: Math.max(0, frozenEnd - frozenStart) };

  const scrollableStart = Math.max(range.start, frozenCount);
  const scrollable: IndexRange = {
    start: scrollableStart,
    count: Math.max(0, rangeEnd - scrollableStart)
  };

  return { frozen, scrollable };
}
