import type { Patch } from "../types";

/**
 * Every patch variant whose identity fits in a short string. `row-reorder`
 * is deliberately not one of them: its identity *is* the whole permutation,
 * so a content key means joining an `Int32Array` with one element per row.
 * For a million-row sort that is ~6.9M chars per patch, and a sort commit
 * produces two (`TableStore` and the layout participant): measured at
 * 63-114ms for the pair in Chromium (183-194ms in Node), on every commit and
 * again on every undo/redo — to build a merged array `#publishCommitReceipt`
 * only ever tests for truthiness. `mergePatches` compares those few
 * permutations element-wise instead, which measured 1.3-5.0ms for the pair.
 */
type KeyedPatch = Exclude<Patch, { readonly type: "row-reorder" }>;

function patchKey(patch: KeyedPatch): string {
  switch (patch.type) {
    case "cell-value":
    case "cell-style":
      return `${patch.type}:${patch.row}:${patch.column}`;
    case "range-value":
      return [
        "range-value",
        patch.range.startRow,
        patch.range.startColumn,
        patch.range.endRow,
        patch.range.endColumn
      ].join(":");
    case "row-structure":
      return [
        "row-structure",
        patch.operation,
        patch.start,
        patch.count,
        patch.target ?? ""
      ].join(":");
    case "column-structure":
      return [
        "column-structure",
        patch.operation,
        patch.start,
        patch.count,
        patch.target ?? ""
      ].join(":");
    case "column-layout":
      return `column-layout:${patch.column}`;
    case "row-layout":
      return [
        "row-layout",
        patch.rows.join(","),
        patch.range.startRow,
        patch.range.startColumn,
        patch.range.endRow,
        patch.range.endColumn
      ].join(":");
    case "merge-layout":
      return "merge-layout";
    case "row-visibility":
      return "row-visibility";
    case "full-reset":
      return "full-reset";
  }
}

/**
 * Element-wise, never by identity: `createChangeSet` `.slice()`s each order,
 * so one sort's data and layout permutations are always different objects
 * with equal content and an identity check would stop them merging.
 */
function sameOrder(left: Int32Array, right: Int32Array): boolean {
  if (left === right) {
    return true;
  }
  if (left.length !== right.length) {
    return false;
  }
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) {
      return false;
    }
  }
  return true;
}

/**
 * Deduplicates structurally identical patches, keeping the first
 * occurrence. This only collapses exact repeats within one commit — merging
 * spatially adjacent dirty regions into one repaint is RenderScheduler's
 * job (Step 5), not this function's.
 */
export function mergePatches(patches: readonly Patch[]): Patch[] {
  const seen = new Set<string>();
  // A list rather than a key set, and scanned rather than hashed: see
  // `KeyedPatch`. One commit holds a handful of permutations at most, so a
  // few element-wise scans measured 1.3-5.0ms for a million-row sort's pair
  // and allocate nothing, against 63-114ms and ~14MB to key it by content.
  const seenOrders: Int32Array[] = [];
  const merged: Patch[] = [];

  for (const patch of patches) {
    if (patch.type === "row-reorder") {
      if (seenOrders.some((order) => sameOrder(order, patch.order))) {
        continue;
      }
      seenOrders.push(patch.order);
    } else {
      const key = patchKey(patch);
      if (seen.has(key)) {
        continue;
      }
      seen.add(key);
    }
    merged.push(patch);
  }

  return merged;
}
