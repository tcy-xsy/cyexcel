/**
 * Row permutations use one direction everywhere: `order[newPosition] =
 * oldPosition`, so building the reordered sequence is `next[i] =
 * prev[order[i]]`. Picking the other direction anywhere makes sorting look
 * correct while undo restores the wrong order, so the convention is asserted
 * rather than documented.
 */
export function assertValidRowOrder(order: Int32Array, expectedLength: number): void {
  if (order.length !== expectedLength) {
    throw new RangeError(`row order length ${order.length} does not match row count ${expectedLength}`);
  }
  const seen = new Uint8Array(expectedLength);
  for (let index = 0; index < order.length; index += 1) {
    const source = order[index] as number;
    if (!Number.isInteger(source) || source < 0 || source >= expectedLength || seen[source] === 1) {
      throw new RangeError(`row order is not a permutation at index ${index}: ${source}`);
    }
    seen[source] = 1;
  }
}

export function invertRowOrder(order: Int32Array): Int32Array {
  const inverse = new Int32Array(order.length);
  for (let index = 0; index < order.length; index += 1) {
    inverse[order[index] as number] = index;
  }
  return inverse;
}

export function applyRowOrder<T>(items: readonly T[], order: Int32Array): T[] {
  const next: T[] = new Array(order.length) as T[];
  for (let index = 0; index < order.length; index += 1) {
    next[index] = items[order[index] as number] as T;
  }
  return next;
}
