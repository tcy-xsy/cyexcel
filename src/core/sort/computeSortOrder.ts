import type { CellValue } from "../sheet/types";
/**
 * Sorts rows by a user-supplied key extractor and comparator, returning a
 * permutation without touching model state.
 */
export interface SortRequest {
  readonly rowCount: number;
  readonly getSortKey: (row: number) => CellValue | undefined;
  readonly compare: (left: CellValue | undefined, right: CellValue | undefined) => number;
  readonly direction: "asc" | "desc";
  readonly signal?: AbortSignal;
  readonly onProgress?: (ratio: number) => void;
  /**
   * Ascending row indices to sort; every other row keeps its position
   * (identity mapping). Defaults to every row — the ADR 0003 §3.1 invariant
   * "排序只在可见行之间进行，被过滤掉的行原地不动": a filtered row's slot in
   * `filteredRows` must never move, or the set would need remapping after
   * every sort.
   */
  readonly visibleRows?: readonly number[];
}

/**
 * Computes a row permutation without touching any model state.
 *
 * Nothing is applied here: the caller commits the returned order as one
 * atomic mutation. That is what makes cancellation safe — a cancelled sort
 * has written nothing, so there is no partial state to roll back.
 */
export function computeSortOrder(request: SortRequest): Int32Array | null {
  const { rowCount, getSortKey, compare, direction, signal, onProgress, visibleRows } = request;
  if (signal && signal.aborted) {
    return null;
  }

  const targets = visibleRows ?? fullRange(rowCount);

  // Keys up front: a user-supplied getSortKey may format or parse, and a
  // comparison sort would otherwise call it O(n log n) times. Only visible
  // rows need a key at all — a filtered row never enters the comparison.
  const keys = new Array<CellValue | undefined>(targets.length);
  for (let index = 0; index < targets.length; index += 1) {
    keys[index] = getSortKey(targets[index] as number);
  }
  onProgress?.(targets.length === 0 ? 1 : 0.5);
  if (signal && signal.aborted) {
    return null;
  }

  const localOrder = new Array<number>(targets.length);
  for (let index = 0; index < targets.length; index += 1) {
    localOrder[index] = index;
  }
  const sign = direction === "desc" ? -1 : 1;
  localOrder.sort((left, right) => {
    const result = compare(keys[left], keys[right]);
    // Stability is guaranteed by Array.prototype.sort (ES2019+): equal keys
    // keep their original relative order in both ascending and descending.
    return result * sign;
  });

  if (signal && signal.aborted) {
    return null;
  }
  onProgress?.(1);

  // Identity everywhere, then scatter the sorted visible rows back into
  // exactly the slots they came from — hidden rows never move.
  const order = new Int32Array(rowCount);
  for (let row = 0; row < rowCount; row += 1) {
    order[row] = row;
  }
  for (let index = 0; index < targets.length; index += 1) {
    order[targets[index] as number] = targets[localOrder[index] as number] as number;
  }
  return order;
}

function fullRange(rowCount: number): readonly number[] {
  const indices = new Array<number>(rowCount);
  for (let row = 0; row < rowCount; row += 1) {
    indices[row] = row;
  }
  return indices;
}
