export type StructureAxis = "row" | "column";
export type StructureOperation = "insert" | "delete" | "move";

/**
 * The single validated descriptor consumed by model, layout and selection.
 * Move targets always use the pre-removal index space.
 */
export interface AxisChangeDescriptor {
  readonly axis: StructureAxis;
  readonly operation: StructureOperation;
  readonly start: number;
  readonly count: number;
  readonly target?: number;
  readonly targetSpace?: "pre-removal";
}

export interface StructuralChange {
  readonly operation: "insert" | "delete" | "move";
  readonly start: number;
  readonly count: number;
  /** Only meaningful for "move" — the insertion index in the *original* (pre-removal) index space, same convention `ArrayDataSource.moveRows()`/`moveColumns()` use. */
  readonly target?: number;
}

function assertNonNegativeInteger(name: string, value: number): void {
  if (!Number.isSafeInteger(value) || value < 0) {
    throw new RangeError(`${name} must be a non-negative safe integer: ${value}`);
  }
}

/** Validates structure edits before any participant mutates its own state. */
export function validateAxisChange(change: AxisChangeDescriptor, axisLength: number): void {
  assertNonNegativeInteger("axisLength", axisLength);
  assertNonNegativeInteger("start", change.start);
  assertNonNegativeInteger("count", change.count);
  if (change.count === 0) {
    throw new RangeError("count must be greater than zero");
  }

  switch (change.operation) {
    case "insert":
      if (change.start > axisLength) {
        throw new RangeError(`insert start ${change.start} is outside axis length ${axisLength}`);
      }
      return;
    case "delete":
      if (change.start + change.count > axisLength) {
        throw new RangeError("delete range is outside axis length");
      }
      return;
    case "move": {
      const target = change.target;
      if (target === undefined) {
        throw new RangeError("move requires a target in pre-removal index space");
      }
      assertNonNegativeInteger("target", target);
      if (change.start + change.count > axisLength || target > axisLength) {
        throw new RangeError("move range or target is outside axis length");
      }
      if (target > change.start && target < change.start + change.count) {
        throw new RangeError("move target cannot be inside the moved block");
      }
    }
  }
}

/**
 * Maps an index from before a structural change (row or column — the
 * arithmetic doesn't know or care which axis it's for) to where that same
 * logical row/column lives afterward. Shared by `rowStructure.ts`/
 * `columnStructure.ts` (`doc/progress.md`'s Step 12 entry flagged this as
 * needing a shared abstraction once a second axis needed the same
 * remapping — Step 17 is that second axis).
 *
 * - insert: indices at or after `start` shift by `count`.
 * - delete: indices before `start` are unaffected; indices inside the
 *   deleted band collapse onto `start` (there's no "correct" surviving
 *   index for them, so they land on whatever now occupies that position);
 *   indices after the band shift by `count`.
 * - move: indices inside the moved band carry their own offset within that
 *   band to the new location; indices outside it shift as if the band had
 *   been deleted then re-inserted (see the two `afterDelete`/
 *   `adjustedTarget` steps below).
 */
export function remapIndex(index: number, change: StructuralChange): number {
  const { operation, start, count } = change;

  if (operation === "insert") {
    return index >= start ? index + count : index;
  }

  if (operation === "delete") {
    if (index < start) {
      return index;
    }
    if (index < start + count) {
      return start;
    }
    return index - count;
  }

  // move
  const target = change.target ?? start;
  const blockEnd = start + count;
  const adjustedTarget = target > start ? target - count : target;

  if (index >= start && index < blockEnd) {
    // Inside the moved block — carries its own offset to the new location.
    return adjustedTarget + (index - start);
  }

  // Outside the block: equivalent to deleting the band, then inserting it
  // (of the same size) at adjustedTarget.
  const afterDelete = index < start ? index : index - count;
  return afterDelete >= adjustedTarget ? afterDelete + count : afterDelete;
}
