/**
 * The comparator `sortByColumn()` falls back to when the column's cell type
 * defines no `compare`.
 *
 * It is a *factory* rather than a bare comparator because "empty cells sort
 * last" cannot be expressed by a direction-agnostic comparator:
 * `computeSortOrder` applies the direction by multiplying the comparator's
 * result by `-1` for `"desc"`, which would flip a fixed "nullish is greater"
 * verdict into "nullish is smaller" and float every empty cell to the top.
 * Pre-negating the nullish verdict here is what keeps empties at the bottom
 * in both directions.
 */
export function createDefaultSortCompare(
  direction: "asc" | "desc"
): (left: unknown, right: unknown) => number {
  // Cancels the `* sign` that computeSortOrder applies for "desc".
  const undoDirection = direction === "desc" ? -1 : 1;
  return (left, right) => {
    const leftEmpty = left === null || left === undefined;
    const rightEmpty = right === null || right === undefined;
    if (leftEmpty || rightEmpty) {
      if (leftEmpty && rightEmpty) {
        return 0;
      }
      return (leftEmpty ? 1 : -1) * undoDirection;
    }
    if (typeof left === "number" && typeof right === "number") {
      // Not `left - right`: that overflows to Infinity/NaN on extreme values
      // and on NaN keys, where a 0 (treat as equal, keep stable order) is the
      // only answer that keeps the sort well-ordered.
      return left < right ? -1 : left > right ? 1 : 0;
    }
    // Codepoint order, not `localeCompare`: a locale-sensitive default would
    // make the same data sort differently per machine. A column that wants
    // locale-aware text ordering says so through its cell type's `compare`.
    const leftText = String(left);
    const rightText = String(right);
    return leftText < rightText ? -1 : leftText > rightText ? 1 : 0;
  };
}
