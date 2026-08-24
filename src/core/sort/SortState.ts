/**
 * Which column the header's sort indicator currently points at, and in
 * which direction. This mirrors the *last* `sortByColumn()` call's outcome,
 * not a live claim about the data's current order — editing a cell after
 * sorting does not clear it (see ADR 0003 §3.4, and Excel's own behavior).
 */
export interface SortState {
  readonly column: number;
  readonly direction: "asc" | "desc";
}

/**
 * The header sort indicator's only two states — ascending and descending —
 * cycling forever between them. There is deliberately no third "restore
 * original order" state a click can land on: a physical sort destroys the
 * original order, so a click that promises to bring it back and then
 * can't would be worse than not offering it at all (ADR 0003 §3.4). Only
 * undo can recover the original order.
 *
 * Clicking a *different* column always starts that column fresh at "asc":
 * `sortByColumn` sorts one column at a time, so only one column's
 * indicator is ever shown.
 */
export function nextSortDirection(current: SortState | null, column: number): "asc" | "desc" {
  if (current !== null && current.column === column && current.direction === "asc") {
    return "desc";
  }
  return "asc";
}
