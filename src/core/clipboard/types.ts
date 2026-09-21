import type { ColumnClipboardPolicy } from "../cells";
import type { CellRange } from "../types";
import type { CopiedStyles } from "./styleClipboard";

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
  /**
   * 这次粘贴认领到的样式（`CopyStyleCache.claim()` 的结果）。只有"剪贴板里这份
   * 就是本应用刚拷出去的那份"时才会有——跨应用粘贴永远没有，那时只粘文本。
   */
  readonly styles?: CopiedStyles;
  readonly range?: CellRange;
  /** Pasted text to use instead of reading the system clipboard — mainly for tests and programmatic paste. */
  readonly text?: string;
  readonly signal?: AbortSignal;
}
