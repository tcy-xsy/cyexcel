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

/**
 * Which overlay and pinned-band consumers a ChangeSet invalidates.
 *
 * Declared here rather than imported from render/'s `PresentationDirty`, which
 * it is structurally identical to: data/ must not depend on the renderer. The
 * renderer accepts this shape as-is.
 */
export interface DirtyPresentation {
  readonly domCells: boolean;
  readonly columnHeader: boolean;
  readonly rowHeader: boolean;
  readonly domHeaders: boolean;
}

/** Rendering-only invalidation derived from a ChangeSet. */
export interface DirtySet {
  readonly content: "none" | "full" | readonly DirtyRegion[];
  readonly skeleton: "none" | "row-axis" | "column-axis" | "full";
  readonly selection: boolean;
  readonly presentation: DirtyPresentation;
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
    case "row-reorder":
      // Object.freeze() throws on a non-empty TypedArray, so the caller's
      // live Int32Array must be copied, not frozen in place.
      return Object.freeze({ ...change, order: change.order.slice() });
    case "cell-value":
    case "cell-style":
    case "row-structure":
    case "column-structure":
    case "column-layout":
    case "full-reset":
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
 * Which axis-level facts a commit carried. Every consumer that used to scan the
 * change list itself — and CyExcelCore scanned it three times over — asks for
 * this once instead.
 */
export interface StructuralChangeSummary {
  readonly hasRowStructure: boolean;
  readonly hasColumnStructure: boolean;
  readonly hasColumnLayout: boolean;
  readonly hasRowLayout: boolean;
}

export function summarizeStructuralChanges(
  changes: readonly SemanticChange[]
): StructuralChangeSummary {
  let hasRowStructure = false;
  let hasColumnStructure = false;
  let hasColumnLayout = false;
  let hasRowLayout = false;
  for (const change of changes) {
    switch (change.type) {
      case "row-structure":
      case "row-reorder":
        hasRowStructure = true;
        break;
      case "column-structure":
        hasColumnStructure = true;
        break;
      case "column-layout":
        hasColumnLayout = true;
        break;
      case "row-layout":
        hasRowLayout = true;
        break;
    }
  }
  return Object.freeze({ hasRowStructure, hasColumnStructure, hasColumnLayout, hasRowLayout });
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
  let domCells = false;
  let columnHeader = false;
  let rowHeader = false;
  let domHeaders = false;

  for (const change of changes) {
    switch (change.type) {
      case "cell-value":
      case "cell-style":
        domCells = true;
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
        domCells = true;
        if (content !== "full") {
          regions.push(Object.freeze({ range: change.range }));
          content = regions;
        }
        break;
      case "row-layout":
        domCells = true;
        rowHeader = true;
        if (content !== "full") {
          regions.push(Object.freeze({ range: change.range }));
          content = regions;
        }
        skeleton = skeleton === "column-axis" ? "full" : "row-axis";
        break;
      case "column-layout":
        // Column geometry moves every band at once, so this is the same
        // wholesale invalidation the legacy per-patch path reached for.
        domCells = columnHeader = rowHeader = domHeaders = true;
        content = "full";
        skeleton = skeleton === "row-axis" ? "full" : "column-axis";
        break;
      case "row-structure":
      case "row-reorder":
        // A reorder does not itself change row numbers, but each row's
        // height travels with it, so band boundaries move and the gutter
        // must repaint just as it would for insert/delete/move (see
        // doc/scene-mutation-migration-plan.md §14.7).
        domCells = true;
        rowHeader = true;
        content = "full";
        skeleton = skeleton === "column-axis" ? "full" : "row-axis";
        selection = true;
        break;
      case "column-structure":
        domCells = columnHeader = rowHeader = domHeaders = true;
        content = "full";
        skeleton = skeleton === "row-axis" ? "full" : "column-axis";
        selection = true;
        break;
      case "full-reset":
        domCells = columnHeader = rowHeader = domHeaders = true;
        content = "full";
        skeleton = "full";
        selection = true;
        break;
    }
  }

  return Object.freeze({
    // "none" was previously unreachable: an empty region list was returned as
    // `[]`, so the type advertised a case the function never produced.
    content:
      content === "full" ? content : regions.length === 0 ? "none" : Object.freeze([...regions]),
    skeleton,
    selection,
    presentation: Object.freeze({ domCells, columnHeader, rowHeader, domHeaders })
  });
}
