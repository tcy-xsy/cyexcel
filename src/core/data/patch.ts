import type { Patch } from "../types";

function patchKey(patch: Patch): string {
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
    case "full-reset":
      return "full-reset";
    case "row-reorder":
      return ["row-reorder", patch.order.join(",")].join(":");
  }
}

/**
 * Deduplicates structurally identical patches, keeping the first
 * occurrence. This only collapses exact repeats within one commit — merging
 * spatially adjacent dirty regions into one repaint is RenderScheduler's
 * job (Step 5), not this function's.
 */
export function mergePatches(patches: readonly Patch[]): Patch[] {
  const seen = new Set<string>();
  const merged: Patch[] = [];

  for (const patch of patches) {
    const key = patchKey(patch);
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    merged.push(patch);
  }

  return merged;
}
