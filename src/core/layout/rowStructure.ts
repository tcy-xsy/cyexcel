export interface RowStructureChange {
  readonly operation: "insert" | "delete" | "move";
  readonly start: number;
  readonly count: number;
  /** Only meaningful for "move" — the insertion index in the *original* (pre-removal) index space, same convention `ArrayDataSource.moveRows()` uses. */
  readonly target?: number;
}

/**
 * Maps a row index from before a structural change to where that same
 * logical row lives afterward — used to keep SelectionModel's active/
 * anchor rows pointing at the row the user actually had selected, not
 * just clamped into the new (possibly smaller) bounds.
 *
 * - insert: rows at or after `start` shift down by `count`.
 * - delete: rows before `start` are unaffected; rows inside the deleted
 *   band collapse onto `start` (there's no "correct" surviving row for
 *   them, so they land on whatever now occupies that position); rows
 *   after the band shift up by `count`.
 * - move: rows inside the moved band carry their own offset within that
 *   band to the new location; rows outside it shift as if the band had
 *   been deleted then re-inserted (see the two `afterDelete`/
 *   `adjustedTarget` steps below).
 */
export function remapRowIndex(row: number, change: RowStructureChange): number {
  const { operation, start, count } = change;

  if (operation === "insert") {
    return row >= start ? row + count : row;
  }

  if (operation === "delete") {
    if (row < start) {
      return row;
    }
    if (row < start + count) {
      return start;
    }
    return row - count;
  }

  // move
  const target = change.target ?? start;
  const blockEnd = start + count;
  const adjustedTarget = target > start ? target - count : target;

  if (row >= start && row < blockEnd) {
    // Inside the moved block — carries its own offset to the new location.
    return adjustedTarget + (row - start);
  }

  // Outside the block: equivalent to deleting the band, then inserting it
  // (of the same size) at adjustedTarget.
  const afterDelete = row < start ? row : row - count;
  return afterDelete >= adjustedTarget ? afterDelete + count : afterDelete;
}
