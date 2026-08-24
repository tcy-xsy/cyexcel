/**
 * The comparator `sortByColumn()` sorts with.
 *
 * `compareDefined` is the column's cell type's `compare`, when it defines
 * one; every pair of non-empty values is handed to it untouched. Without it,
 * two non-empty values are ordered numerically when both are numbers and **by
 * their string form otherwise** — so a cell type whose values have a natural
 * order that is not their string order (dates, say) has to define `compare`,
 * or its column sorts by `String(value)`.
 *
 * Empty cells are handled here rather than by the cell type for two reasons.
 * A `CellType.compare` is not told the direction, and "empty cells sort last"
 * cannot be expressed without it: `computeSortOrder` applies the direction by
 * multiplying the comparator's result by `-1` for `"desc"`, which flips a
 * fixed "nullish is greater" verdict into "nullish is smaller" and floats
 * every empty cell to the top. Pre-negating that verdict here is what keeps
 * empties at the bottom in both directions — and doing it once here is also
 * what spares every cell type from having to guard `null`/`undefined` before
 * it can touch a value (`number`'s `left - right` would yield `NaN`, `date`'s
 * `getTime()` would throw).
 */
export function createDefaultSortCompare(
  direction: "asc" | "desc",
  compareDefined?: (left: unknown, right: unknown) => number
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
    if (compareDefined) {
      return compareDefined(left, right);
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
