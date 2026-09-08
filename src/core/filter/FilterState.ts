/**
 * Which column the header's filter indicator currently points at, and the
 * predicate last used to build the active `filteredRows` set. Descriptive
 * only — mirrors `SortState`'s own role: it drives `HeaderState.filtered`
 * and nothing else, and is not transactionally coupled to `LayoutRoot`
 * (unlike `filteredRows` itself, which lives there so undo/redo of a
 * structural row command carries it along automatically). MVP is a single
 * active filter, one column at a time: calling `filterByColumn` again
 * replaces this rather than composing with the previous one.
 */
export interface FilterState {
  readonly column: number;
  readonly predicate: (value: unknown, row: number) => boolean;
}
