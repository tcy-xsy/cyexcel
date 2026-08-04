import type { CellRange, Patch } from "../types";

export interface FrameScheduler {
  readonly requestFrame: (callback: () => void) => number;
  readonly cancelFrame: (handle: number) => void;
}

export type DirtyState =
  | { readonly kind: "none" }
  | { readonly kind: "regions"; readonly ranges: readonly CellRange[] }
  | { readonly kind: "full" };

export interface RenderTarget {
  paintContent(dirty: DirtyState): void;
  paintSelection(): void;
}

export interface RenderSchedulerOptions {
  /** No default "real RAF" is assumed — callers inject one for their environment. */
  readonly frameScheduler: FrameScheduler;
  readonly maxDirtyRegions?: number;
}

const DEFAULT_MAX_DIRTY_REGIONS = 16;

function rangesTouch(a: CellRange, b: CellRange): boolean {
  const rowsClose = a.startRow <= b.endRow + 1 && b.startRow <= a.endRow + 1;
  const columnsClose = a.startColumn <= b.endColumn + 1 && b.startColumn <= a.endColumn + 1;
  return rowsClose && columnsClose;
}

function mergeRanges(a: CellRange, b: CellRange): CellRange {
  return {
    startRow: Math.min(a.startRow, b.startRow),
    startColumn: Math.min(a.startColumn, b.startColumn),
    endRow: Math.max(a.endRow, b.endRow),
    endColumn: Math.max(a.endColumn, b.endColumn)
  };
}

/**
 * Coalesces invalidation requests into one requestFrame per animation
 * frame, and tracks content-layer and selection-layer dirtiness
 * independently so a selection-only change never forces a content repaint.
 */
export class RenderScheduler {
  readonly #target: RenderTarget;
  readonly #frameScheduler: FrameScheduler;
  readonly #maxDirtyRegions: number;

  #contentDirty: DirtyState = { kind: "none" };
  #selectionDirty = false;
  #frameHandle: number | null = null;

  constructor(target: RenderTarget, options: RenderSchedulerOptions) {
    this.#target = target;
    this.#frameScheduler = options.frameScheduler;
    this.#maxDirtyRegions = options.maxDirtyRegions ?? DEFAULT_MAX_DIRTY_REGIONS;
  }

  invalidatePatch(patch: Patch): void {
    switch (patch.type) {
      case "cell-value":
      case "cell-style":
        this.#addRegion({
          startRow: patch.row,
          startColumn: patch.column,
          endRow: patch.row,
          endColumn: patch.column
        });
        break;
      case "range-value":
        this.#addRegion(patch.range);
        break;
      case "row-structure":
      case "column-layout":
      case "full-reset":
        this.invalidateAll();
        break;
    }
  }

  invalidateAll(): void {
    this.#contentDirty = { kind: "full" };
    this.#scheduleFrame();
  }

  invalidateSelection(): void {
    this.#selectionDirty = true;
    this.#scheduleFrame();
  }

  dispose(): void {
    if (this.#frameHandle !== null) {
      this.#frameScheduler.cancelFrame(this.#frameHandle);
      this.#frameHandle = null;
    }
  }

  #addRegion(range: CellRange): void {
    if (this.#contentDirty.kind === "full") {
      this.#scheduleFrame();
      return;
    }

    const regions =
      this.#contentDirty.kind === "regions" ? [...this.#contentDirty.ranges] : [];

    let merged = range;
    let changed = true;
    while (changed) {
      changed = false;
      for (let index = 0; index < regions.length; index += 1) {
        const candidate = regions[index];
        if (candidate && rangesTouch(candidate, merged)) {
          merged = mergeRanges(candidate, merged);
          regions.splice(index, 1);
          changed = true;
          break;
        }
      }
    }
    regions.push(merged);

    this.#contentDirty =
      regions.length > this.#maxDirtyRegions
        ? { kind: "full" }
        : { kind: "regions", ranges: regions };

    this.#scheduleFrame();
  }

  #scheduleFrame(): void {
    if (this.#frameHandle !== null) {
      return;
    }
    this.#frameHandle = this.#frameScheduler.requestFrame(() => {
      this.#frameHandle = null;
      this.#flush();
    });
  }

  #flush(): void {
    if (this.#contentDirty.kind !== "none") {
      const dirty = this.#contentDirty;
      this.#contentDirty = { kind: "none" };
      this.#target.paintContent(dirty);
    }
    if (this.#selectionDirty) {
      this.#selectionDirty = false;
      this.#target.paintSelection();
    }
  }
}
