import type { Disposer } from "../types";
import type { SelectionModel, SelectionSnapshot } from "./SelectionModel";

export interface KeyboardControllerOptions {
  readonly selection: SelectionModel;
  readonly onSelectionChange?: (snapshot: SelectionSnapshot) => void;
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

/**
 * Keyboard input for selection. Listeners are always attached to a specific
 * element (never window/document) — that's the entire mechanism behind "an
 * unfocused table doesn't hijack page shortcuts" and "only the focused
 * instance reacts": native DOM event routing already guarantees both.
 */
export class KeyboardController {
  readonly #selection: SelectionModel;
  readonly #onSelectionChange: ((snapshot: SelectionSnapshot) => void) | undefined;

  constructor(options: KeyboardControllerOptions) {
    this.#selection = options.selection;
    this.#onSelectionChange = options.onSelectionChange;
  }

  handleKeyDown(key: string, modifiers: KeyboardModifiers = { shift: false }): boolean {
    const arrowDelta = ARROW_DELTAS[key];
    if (arrowDelta) {
      this.#selection.moveActive(arrowDelta[0], arrowDelta[1], modifiers.shift);
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

  #notify(): void {
    this.#onSelectionChange?.(this.#selection.snapshot);
  }
}
