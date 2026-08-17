import type { CellRange, Patch } from "../types";
import type { ModelRevision, RevisionGate } from "../model";
import type { SceneSnapshot } from "../scene";

export interface FrameScheduler {
  readonly requestFrame: (callback: () => void) => number;
  readonly cancelFrame: (handle: number) => void;
}

export type DirtyState =
  | { readonly kind: "none" }
  | { readonly kind: "regions"; readonly ranges: readonly CellRange[] }
  | { readonly kind: "full" };

export interface PresentationDirty {
  readonly domCells: boolean;
  readonly columnHeader: boolean;
  readonly rowHeader: boolean;
  readonly domHeaders: boolean;
}

export interface RenderFrame<TScene extends { readonly modelVersion: number } = SceneSnapshot> {
  readonly scene: TScene;
  readonly revision: ModelRevision;
  readonly dirty: {
    readonly content: DirtyState;
    readonly selection: boolean;
    readonly presentation: PresentationDirty;
  };
}

export interface RenderTarget<TScene extends { readonly modelVersion: number } = SceneSnapshot> {
  paintPresentation(frame: RenderFrame<TScene>): void;
  paintContent(frame: RenderFrame<TScene>): void;
  paintSelection(frame: RenderFrame<TScene>): void;
}

export interface RenderSchedulerOptions<TScene extends { readonly modelVersion: number } = SceneSnapshot> {
  /** No default "real RAF" is assumed — callers inject one for their environment. */
  readonly frameScheduler: FrameScheduler;
  readonly maxDirtyRegions?: number;
  readonly revisionGate: RevisionGate;
  readonly getScene: () => TScene;
  readonly onFrameError?: (error: unknown) => void;
}

const DEFAULT_MAX_DIRTY_REGIONS = 16;
const CLEAN_PRESENTATION: PresentationDirty = Object.freeze({
  domCells: false,
  columnHeader: false,
  rowHeader: false,
  domHeaders: false
});
const FULL_PRESENTATION: PresentationDirty = Object.freeze({
  domCells: true,
  columnHeader: true,
  rowHeader: true,
  domHeaders: true
});

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
export class RenderScheduler<TScene extends { readonly modelVersion: number } = SceneSnapshot> {
  readonly #target: RenderTarget<TScene>;
  readonly #frameScheduler: FrameScheduler;
  readonly #maxDirtyRegions: number;
  readonly #revisionGate: RevisionGate;
  readonly #getScene: () => TScene;
  readonly #onFrameError: ((error: unknown) => void) | undefined;

  #contentDirty: DirtyState = { kind: "none" };
  #selectionDirty = false;
  #presentationDirty: PresentationDirty = CLEAN_PRESENTATION;
  #frameHandle: number | null = null;

  constructor(target: RenderTarget<TScene>, options: RenderSchedulerOptions<TScene>) {
    this.#target = target;
    this.#frameScheduler = options.frameScheduler;
    this.#maxDirtyRegions = options.maxDirtyRegions ?? DEFAULT_MAX_DIRTY_REGIONS;
    this.#revisionGate = options.revisionGate;
    this.#getScene = options.getScene;
    this.#onFrameError = options.onFrameError;
  }

  invalidatePatch(patch: Patch): void {
    switch (patch.type) {
      case "cell-value":
      case "cell-style":
        this.invalidatePresentation({ domCells: true });
        this.#addRegion({
          startRow: patch.row,
          startColumn: patch.column,
          endRow: patch.row,
          endColumn: patch.column
        });
        break;
      case "range-value":
        this.invalidatePresentation({ domCells: true });
        this.#addRegion(patch.range);
        break;
      case "row-layout":
        this.invalidatePresentation({ domCells: true, rowHeader: true });
        this.#addRegion(patch.range);
        break;
      case "row-structure":
        this.invalidatePresentation({ domCells: true, rowHeader: true });
        this.#contentDirty = { kind: "full" };
        this.#scheduleFrame();
        break;
      case "column-structure":
      case "column-layout":
      case "full-reset":
        this.invalidateAll();
        break;
    }
  }

  invalidateAll(): void {
    this.#contentDirty = { kind: "full" };
    this.#presentationDirty = FULL_PRESENTATION;
    this.#scheduleFrame();
  }

  invalidateSelection(): void {
    this.#selectionDirty = true;
    this.#scheduleFrame();
  }

  invalidatePresentation(dirty: Partial<PresentationDirty>): void {
    this.#presentationDirty = Object.freeze({
      domCells: this.#presentationDirty.domCells || dirty.domCells === true,
      columnHeader: this.#presentationDirty.columnHeader || dirty.columnHeader === true,
      rowHeader: this.#presentationDirty.rowHeader || dirty.rowHeader === true,
      domHeaders: this.#presentationDirty.domHeaders || dirty.domHeaders === true
    });
    this.#scheduleFrame();
  }

  /**
   * Paints the complete pending publication synchronously. The already
   * scheduled callback is cancelled first so the same dirty snapshot cannot
   * be replayed by a later animation frame.
   */
  flushNow(): void {
    if (this.#frameHandle !== null) {
      this.#frameScheduler.cancelFrame(this.#frameHandle);
      this.#frameHandle = null;
    }
    if (!this.#hasDirtyWork()) {
      return;
    }
    this.#flush();
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
    const lease = this.#revisionGate.acquireReadLease();
    try {
      const scene = this.#getScene();
      const revision = lease.revision;
      if (scene.modelVersion !== revision) {
        throw new Error(
          `Render frame revision ${revision} does not match scene modelVersion ${scene.modelVersion}`
        );
      }
      const contentDirty = freezeDirtyState(this.#contentDirty);
      const presentationDirty = freezePresentationDirty(this.#presentationDirty);
      const frame: RenderFrame<TScene> = Object.freeze({
        scene,
        revision,
        dirty: Object.freeze({
          content: contentDirty,
          selection: this.#selectionDirty,
          presentation: presentationDirty
        })
      });
      if (hasPresentationDirty(frame.dirty.presentation)) {
        this.#target.paintPresentation(frame);
      }
      if (frame.dirty.content.kind !== "none") {
        this.#target.paintContent(frame);
      }
      if (frame.dirty.selection) {
        this.#target.paintSelection(frame);
      }
      this.#contentDirty = { kind: "none" };
      this.#selectionDirty = false;
      this.#presentationDirty = CLEAN_PRESENTATION;
    } catch (error) {
      if (this.#onFrameError) {
        this.#onFrameError(error);
        return;
      }
      throw error;
    } finally {
      lease.release();
    }
  }

  #hasDirtyWork(): boolean {
    return (
      this.#contentDirty.kind !== "none" ||
      this.#selectionDirty ||
      hasPresentationDirty(this.#presentationDirty)
    );
  }
}

function hasPresentationDirty(dirty: PresentationDirty): boolean {
  return dirty.domCells || dirty.columnHeader || dirty.rowHeader || dirty.domHeaders;
}

function freezePresentationDirty(dirty: PresentationDirty): PresentationDirty {
  return Object.freeze({ ...dirty });
}

function freezeDirtyState(dirty: DirtyState): DirtyState {
  if (dirty.kind !== "regions") {
    return Object.freeze({ kind: dirty.kind });
  }
  return Object.freeze({
    kind: "regions",
    ranges: Object.freeze(
      dirty.ranges.map((range) => Object.freeze({ ...range }))
    )
  });
}
