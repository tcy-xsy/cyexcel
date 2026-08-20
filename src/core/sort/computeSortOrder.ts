/**
 * Sorts rows by a user-supplied key extractor and comparator, returning a
 * permutation without touching model state.
 */
export interface SortRequest {
  readonly rowCount: number;
  readonly getSortKey: (row: number) => unknown;
  readonly compare: (left: unknown, right: unknown) => number;
  readonly direction: "asc" | "desc";
  readonly signal?: AbortSignal;
  readonly onProgress?: (ratio: number) => void;
}

/**
 * Computes a row permutation without touching any model state.
 *
 * Nothing is applied here: the caller commits the returned order as one
 * atomic mutation. That is what makes cancellation safe — a cancelled sort
 * has written nothing, so there is no partial state to roll back.
 */
export function computeSortOrder(request: SortRequest): Int32Array | null {
  const { rowCount, getSortKey, compare, direction, signal, onProgress } = request;
  if (signal && signal.aborted) {
    return null;
  }

  // Keys up front: a user-supplied getSortKey may format or parse, and a
  // comparison sort would otherwise call it O(n log n) times.
  const keys: unknown[] = new Array(rowCount) as unknown[];
  for (let row = 0; row < rowCount; row += 1) {
    keys[row] = getSortKey(row);
  }
  onProgress?.(rowCount === 0 ? 1 : 0.5);
  if (signal && signal.aborted) {
    return null;
  }

  const indices = new Array<number>(rowCount);
  for (let row = 0; row < rowCount; row += 1) {
    indices[row] = row;
  }
  const sign = direction === "desc" ? -1 : 1;
  indices.sort((left, right) => {
    const result = compare(keys[left], keys[right]);
    // Stability: Array.prototype.sort is stable per spec, but the index
    // tiebreak also keeps `desc` stable instead of reversing equal runs.
    return result === 0 ? left - right : result * sign;
  });

  if (signal && signal.aborted) {
    return null;
  }
  onProgress?.(1);
  return Int32Array.from(indices);
}
