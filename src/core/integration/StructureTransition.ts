export interface StructureTransitionOptions {
  /** A numeric editor anchor still points at a pre-transition row/column. */
  readonly cancelActiveEdit: () => void;
  readonly cancelBodyGesture: () => void;
  readonly cancelHeaderGesture: () => void;
  readonly closeContextMenu: () => void;
}

/**
 * Owns the structure version and everything that must be torn down when it
 * moves.
 *
 * The counter and the teardown are one thing on purpose. They were previously
 * an open-coded block at the single call site, which meant nothing stopped a
 * future caller from bumping the version without cancelling the gestures
 * anchored to the old one — or from cancelling without bumping, leaving stale
 * intents able to pass a version check they should have failed.
 *
 * Everything that captured `version` before the transition is expected to
 * compare against it afterwards and abandon its work; that comparison is the
 * only thing protecting a queued intent from applying to rows that have since
 * moved.
 */
export class StructureTransition {
  readonly #options: StructureTransitionOptions;
  #version = 0;

  constructor(options: StructureTransitionOptions) {
    this.#options = options;
  }

  /** The current structure version; captured by intents and re-checked on apply. */
  get version(): number {
    return this.#version;
  }

  /**
   * Advances the version and tears down everything anchored to the old one.
   *
   * Call after a commit whose changes altered the row or column identity
   * space. Undo/redo reaches this without going through the public structure
   * facade, so it cannot be done inside those facade methods.
   */
  apply(): void {
    const next = this.#version + 1;
    if (!Number.isSafeInteger(next)) {
      throw new RangeError("Structure generation exhausted");
    }
    // Cancel the edit before publishing the new version: a stale numeric
    // anchor must not be able to commit into the newly published order.
    this.#options.cancelActiveEdit();
    this.#version = next;
    this.#options.cancelBodyGesture();
    this.#options.cancelHeaderGesture();
    this.#options.closeContextMenu();
  }
}
