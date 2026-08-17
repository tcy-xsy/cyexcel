import type { CellRange, Patch } from "../types";

/**
 * Ordered facts emitted by one successful model commit. Unlike the legacy
 * Patch list, this stream is never deduplicated because index remapping and
 * history consumers must observe every structural operation in order.
 */
export type SemanticChange = Patch;

export interface ChangeSet {
  readonly modelVersion: number;
  readonly changes: readonly SemanticChange[];
}

export interface DirtyRegion {
  readonly range: CellRange;
}

/** Rendering-only invalidation derived from a ChangeSet. */
export interface DirtySet {
  readonly content: "none" | "full" | readonly DirtyRegion[];
  readonly skeleton: "none" | "row-axis" | "column-axis" | "full";
  readonly selection: boolean;
}

function snapshotChange(change: SemanticChange): SemanticChange {
  switch (change.type) {
    case "range-value":
      return Object.freeze({ ...change, range: Object.freeze({ ...change.range }) });
    case "row-layout":
      return Object.freeze({
        ...change,
        rows: Object.freeze([
          change.rows[0],
          ...change.rows.slice(1)
        ] as [number, ...number[]]),
        range: Object.freeze({ ...change.range })
      });
    default:
      return Object.freeze({ ...change });
  }
}

export function createChangeSet(modelVersion: number, changes: readonly SemanticChange[]): ChangeSet {
  if (!Number.isSafeInteger(modelVersion) || modelVersion < 0) {
    throw new RangeError(`modelVersion must be a non-negative safe integer: ${modelVersion}`);
  }

  return Object.freeze({
    modelVersion,
    changes: Object.freeze(changes.map(snapshotChange))
  });
}

/**
 * Compiles ordered semantic changes into coalescable drawing work. This is
 * deliberately lossy: consumers that need model semantics must use ChangeSet.
 */
export function compileDirtySet(changes: readonly SemanticChange[]): DirtySet {
  const regions: DirtyRegion[] = [];
  let content: DirtySet["content"] = "none";
  let skeleton: DirtySet["skeleton"] = "none";
  let selection = false;

  for (const change of changes) {
    switch (change.type) {
      case "cell-value":
      case "cell-style":
        if (content !== "full") {
          regions.push({
            range: {
              startRow: change.row,
              startColumn: change.column,
              endRow: change.row,
              endColumn: change.column
            }
          });
          content = regions;
        }
        break;
      case "range-value":
        if (content !== "full") {
          regions.push(Object.freeze({ range: change.range }));
          content = regions;
        }
        break;
      case "row-layout":
        if (content !== "full") {
          regions.push(Object.freeze({ range: change.range }));
          content = regions;
        }
        skeleton = skeleton === "column-axis" ? "full" : "row-axis";
        break;
      case "column-layout":
        content = "full";
        skeleton = skeleton === "row-axis" ? "full" : "column-axis";
        break;
      case "row-structure":
        content = "full";
        skeleton = skeleton === "column-axis" ? "full" : "row-axis";
        selection = true;
        break;
      case "column-structure":
        content = "full";
        skeleton = skeleton === "row-axis" ? "full" : "column-axis";
        selection = true;
        break;
      case "full-reset":
        content = "full";
        skeleton = "full";
        selection = true;
        break;
    }
  }

  return Object.freeze({
    content: content === "full" ? content : Object.freeze([...regions]),
    skeleton,
    selection
  });
}
