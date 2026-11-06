import type { SceneSnapshot } from "../scene";
import type { CellRange } from "../types";
import type { SelectionModel, SelectionSnapshot } from "./SelectionModel";

export interface PointerControllerOptions {
  readonly selection: SelectionModel;
  readonly onSelectionChange?: (snapshot: SelectionSnapshot) => void;
  /**
   * Resolves the merged range (if any) owning a cell. `down()`/`move()`
   * resolve physical row/column via `scene.resolveRow`/`resolveColumn`,
   * bypassing `SceneSnapshot.hitTestCellGeometry()`'s own merge-anchor
   * normalization entirely — so click/drag selection needs this independent
   * normalization to land on the merge's anchor instead of the physical cell.
   */
  readonly findMergeAt?: (row: number, column: number) => CellRange | undefined;
}

export interface PointerModifiers {
  readonly shift: boolean;
  /** 0 for the primary button and 2 for the secondary button. */
  readonly button?: number;
  /** Ctrl on Windows/Linux or Cmd on macOS. */
  readonly ctrlOrMeta?: boolean;
}

export interface PointerPoint {
  readonly x: number;
  readonly y: number;
}

function isWithinRange(
  row: number,
  column: number,
  range: SelectionSnapshot["range"]
): boolean {
  return (
    row >= range.startRow &&
    row <= range.endRow &&
    column >= range.startColumn &&
    column <= range.endColumn
  );
}

/** Incidental pointer jitter below this distance remains a click. */
export const DRAG_THRESHOLD_PX = 4;

/** Resolves body-local coordinates through the immutable Scene snapshot. */
export function resolveRow(y: number, scene: SceneSnapshot): number {
  return scene.resolveRow(y);
}

export function resolveColumn(x: number, scene: SceneSnapshot): number {
  return scene.resolveColumn(x);
}

/**
 * Body-only selection gesture state machine. DOM routing and surface
 * ownership live outside this class; every event supplies its current Scene
 * and structure generation explicitly.
 */
export class PointerController {
  readonly #selection: SelectionModel;
  readonly #onSelectionChange: ((snapshot: SelectionSnapshot) => void) | undefined;
  readonly #findMergeAt: ((row: number, column: number) => CellRange | undefined) | undefined;
  #structureVersion: number | null = null;
  #dragStartPoint: PointerPoint | null = null;
  #dragThresholdExceeded = false;

  constructor(options: PointerControllerOptions) {
    this.#selection = options.selection;
    this.#onSelectionChange = options.onSelectionChange;
    this.#findMergeAt = options.findMergeAt;
  }

  down(
    point: PointerPoint,
    scene: SceneSnapshot,
    structureVersion: number,
    modifiers: PointerModifiers = { shift: false }
  ): void {
    this.cancel();
    const physicalRow = resolveRow(point.y, scene);
    const physicalColumn = resolveColumn(point.x, scene);
    const merge = this.#findMergeAt?.(physicalRow, physicalColumn);
    const row = merge ? merge.startRow : physicalRow;
    const column = merge ? merge.startColumn : physicalColumn;

    // Context-menu invocation inside the active range preserves the current
    // selection and does not begin a drag gesture.
    if (
      modifiers.button === 2 &&
      this.#selection.snapshot.ranges.some(range => isWithinRange(row, column, range))
    ) {
      return;
    }

    if (modifiers.shift) {
      this.#selection.extendTo(row, column);
    } else if (modifiers.ctrlOrMeta && modifiers.button !== 2) {
      this.#selection.addRange(row, column);
    } else {
      this.#selection.selectCell(row, column);
    }

    this.#structureVersion = structureVersion;
    this.#dragStartPoint = point;
    this.#notify();
  }

  move(point: PointerPoint, scene: SceneSnapshot, structureVersion: number): void {
    if (!this.#hasCurrentStructure(structureVersion)) {
      return;
    }
    if (!this.#dragThresholdExceeded) {
      const start = this.#dragStartPoint as PointerPoint;
      if (Math.hypot(point.x - start.x, point.y - start.y) < DRAG_THRESHOLD_PX) {
        return;
      }
      this.#dragThresholdExceeded = true;
    }

    const physicalRow = resolveRow(point.y, scene);
    const physicalColumn = resolveColumn(point.x, scene);
    const merge = this.#findMergeAt?.(physicalRow, physicalColumn);
    const row = merge ? merge.startRow : physicalRow;
    const column = merge ? merge.startColumn : physicalColumn;
    this.#selection.extendTo(row, column);
    this.#notify();
  }

  up(structureVersion: number): void {
    this.#hasCurrentStructure(structureVersion);
    this.cancel();
  }

  cancel(): void {
    this.#structureVersion = null;
    this.#dragStartPoint = null;
    this.#dragThresholdExceeded = false;
  }

  #hasCurrentStructure(structureVersion: number): boolean {
    if (this.#structureVersion === null) {
      return false;
    }
    if (this.#structureVersion !== structureVersion) {
      this.cancel();
      return false;
    }
    return true;
  }

  #notify(): void {
    this.#onSelectionChange?.(this.#selection.snapshot);
  }
}
