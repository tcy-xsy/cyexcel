/**
 * Extends a known sequence of values (one column's cells, or one row's
 * cells, in the source selection) by `targetCount` new values, matching the
 * three fill-handle patterns cyexcel supports (doc/progress.md's fill-handle
 * patch — deliberately not Excel's date/weekday/custom-list detection):
 *  - a single known value is copied as-is;
 *  - two or more known values that are ALL numbers are extrapolated with the
 *    linear step between the first and last value;
 *  - anything else (mixed types, non-numeric, or a mix) repeats the known
 *    sequence cyclically.
 * An empty `sourceValues` (e.g. a target column with no rows selected) has
 * nothing to extend from, so it fills with `undefined`.
 */
export function computeFillValues(sourceValues: readonly unknown[], targetCount: number): unknown[] {
  if (targetCount <= 0) {
    return [];
  }
  if (sourceValues.length === 0) {
    return new Array(targetCount).fill(undefined) as unknown[];
  }
  if (sourceValues.length === 1) {
    return new Array(targetCount).fill(sourceValues[0]) as unknown[];
  }
  if (sourceValues.every((value) => typeof value === "number")) {
    const numbers = sourceValues as readonly number[];
    const first = numbers[0]!;
    const last = numbers[numbers.length - 1]!;
    const step = (last - first) / (numbers.length - 1);
    return Array.from({ length: targetCount }, (_unused, index) => last + step * (index + 1));
  }
  return Array.from({ length: targetCount }, (_unused, index) => sourceValues[index % sourceValues.length]);
}
