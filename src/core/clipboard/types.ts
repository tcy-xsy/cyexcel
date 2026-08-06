import type { ColumnClipboardPolicy } from "../cells";
import type { CellRange } from "../types";

export type { ColumnClipboardPolicy };

export type CopyMode = "display" | "raw" | "structured";

/**
 * Built cheaply, before any full serialization happens — see CopyPlanner.
 * `doc/architecture.md` §9's published shape.
 */
export interface CopyPlan {
  readonly range: CellRange;
  readonly rows: number;
  readonly columns: number;
  readonly cells: number;
  readonly estimatedBytes: number;
  readonly containsComplexValues: boolean;
  readonly recommendedMode: "display" | "raw" | "download";
}

/** `doc/architecture.md` §9's suggested defaults. */
export interface ClipboardOptions {
  readonly maxCells?: number;
  readonly maxBytes?: number;
  readonly oversizedAction?: "confirm-or-download" | "download" | "block";
}

export type ClipboardCopyStatus = "copied" | "oversized" | "downloaded" | "cancelled";

export interface ClipboardCopyResult {
  readonly status: ClipboardCopyStatus;
  readonly plan: CopyPlan;
}

export interface ClipboardCopyOptions {
  readonly mode?: CopyMode;
  readonly range?: CellRange;
  /** Proceed with a copy that planCopy()/copy() has flagged as oversized. */
  readonly confirmed?: boolean;
  /** Skip the system clipboard and produce a downloadable file instead. */
  readonly download?: boolean;
  readonly signal?: AbortSignal;
}

export interface ClipboardPasteOptions {
  readonly range?: CellRange;
  /** Pasted text to use instead of reading the system clipboard — mainly for tests and programmatic paste. */
  readonly text?: string;
  readonly signal?: AbortSignal;
}
