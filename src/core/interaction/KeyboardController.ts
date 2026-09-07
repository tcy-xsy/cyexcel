import type { CellRange, Disposer } from "../types";
import type { SelectionModel, SelectionSnapshot } from "./SelectionModel";

export interface KeyboardControllerOptions {
  readonly selection: SelectionModel;
  readonly onSelectionChange?: (snapshot: SelectionSnapshot) => void;
  /** Resolves the merged range (if any) owning a cell, for arrow-key stepping over merges. */
  readonly findMergeAt?: (row: number, column: number) => CellRange | undefined;
}

interface KeyboardModifiers {
  readonly shift: boolean;
}

const ARROW_DELTAS: Readonly<Record<string, readonly [number, number]>> = {
  ArrowUp: [-1, 0],
  ArrowDown: [1, 0],
  ArrowLeft: [0, -1],
  ArrowRight: [0, 1]
};

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

/**
 * Keyboard input for selection. Listeners are always attached to a specific
 * element (never window/document) — that's the entire mechanism behind "an
 * unfocused table doesn't hijack page shortcuts" and "only the focused
 * instance responds": native DOM event routing already guarantees both.
 */
export class KeyboardController {
  readonly #selection: SelectionModel;
  readonly #onSelectionChange: ((snapshot: SelectionSnapshot) => void) | undefined;
  readonly #findMergeAt: ((row: number, column: number) => CellRange | undefined) | undefined;

  constructor(options: KeyboardControllerOptions) {
    this.#selection = options.selection;
    this.#onSelectionChange = options.onSelectionChange;
    this.#findMergeAt = options.findMergeAt;
  }

  handleKeyDown(key: string, modifiers: KeyboardModifiers = { shift: false }): boolean {
    const arrowDelta = ARROW_DELTAS[key];
    if (arrowDelta) {
      this.#moveActiveStep(arrowDelta[0], arrowDelta[1], modifiers.shift);
      this.#notify();
      return true;
    }

    switch (key) {
      case "Tab":
        this.#selection.moveActive(0, modifiers.shift ? -1 : 1, false);
        this.#notify();
        return true;
      case "Enter":
        this.#selection.moveActive(modifiers.shift ? -1 : 1, 0, false);
        this.#notify();
        return true;
      case "Home":
        this.#selection.moveActiveTo(this.#selection.snapshot.active.row, 0, modifiers.shift);
        this.#notify();
        return true;
      case "End":
        this.#selection.moveActiveTo(
          this.#selection.snapshot.active.row,
          this.#selection.bounds.columnCount - 1,
          modifiers.shift
        );
        this.#notify();
        return true;
      default:
        return false;
    }
  }

  attach(element: HTMLElement): Disposer {
    const onKeyDown = (event: KeyboardEvent): void => {
      if (this.handleKeyDown(event.key, { shift: event.shiftKey })) {
        event.preventDefault();
      }
    };

    element.addEventListener("keydown", onKeyDown);
    return () => element.removeEventListener("keydown", onKeyDown);
  }

  /**
   * Arrow-key stepping over merges: steps from the merge's far edge (the
   * edge facing the direction of travel), not from the anchor or the active
   * cell's own position — stepping from the anchor would land back inside
   * the same merge and get re-normalized to that same anchor, making the
   * arrow key a no-op.
   */
  #moveActiveStep(dr: number, dc: number, shift: boolean): void {
    const active = this.#selection.snapshot.active;
    const merge = this.#findMergeAt?.(active.row, active.column);
    const fromRow = merge && dr !== 0 ? (dr > 0 ? merge.endRow : merge.startRow) : active.row;
    const fromColumn = merge && dc !== 0 ? (dc > 0 ? merge.endColumn : merge.startColumn) : active.column;
    const bounds = this.#selection.bounds;
    const targetRow = clamp(fromRow + dr, 0, bounds.rowCount - 1);
    const targetColumn = clamp(fromColumn + dc, 0, bounds.columnCount - 1);
    const landing = this.#findMergeAt?.(targetRow, targetColumn);
    const anchorRow = landing ? landing.startRow : targetRow;
    const anchorColumn = landing ? landing.startColumn : targetColumn;
    if (shift) {
      this.#selection.extendTo(anchorRow, anchorColumn);
    } else {
      this.#selection.selectCell(anchorRow, anchorColumn);
    }
  }

  #notify(): void {
    this.#onSelectionChange?.(this.#selection.snapshot);
  }
}
