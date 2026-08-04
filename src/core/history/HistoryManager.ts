import type { Command, Patch } from "../types";

export interface HistoryEntry {
  readonly commands: readonly Command[];
  readonly patches: readonly Patch[];
}

export interface HistoryManagerOptions {
  /** Maximum number of entries kept in the undo stack. Default: 100. */
  readonly maxEntries?: number | undefined;
  /**
   * Approximate byte budget for the undo stack, estimated via
   * `JSON.stringify(entry.patches).length` per entry. This is a cheap
   * heuristic, not an exact memory measurement. Default: ~5,000,000
   * (roughly 5MB of stringified patch data).
   */
  readonly maxBytes?: number | undefined;
}

const DEFAULT_MAX_ENTRIES = 100;
const DEFAULT_MAX_BYTES = 5_000_000;

/**
 * Pure undo/redo stack bookkeeping — knows nothing about how a Command is
 * executed or undone. CommandManager owns that; this only tracks which
 * HistoryEntry objects exist and evicts the oldest ones under budget
 * pressure.
 */
export class HistoryManager {
  #undoStack: HistoryEntry[] = [];
  #redoStack: HistoryEntry[] = [];
  readonly #maxEntries: number;
  readonly #maxBytes: number;

  constructor(options: HistoryManagerOptions = {}) {
    this.#maxEntries = options.maxEntries ?? DEFAULT_MAX_ENTRIES;
    this.#maxBytes = options.maxBytes ?? DEFAULT_MAX_BYTES;
  }

  get size(): number {
    return this.#undoStack.length;
  }

  canUndo(): boolean {
    return this.#undoStack.length > 0;
  }

  canRedo(): boolean {
    return this.#redoStack.length > 0;
  }

  push(entry: HistoryEntry): void {
    this.#undoStack.push(entry);
    this.#redoStack = [];
    this.#evict();
  }

  undo(): HistoryEntry | undefined {
    const entry = this.#undoStack.pop();
    if (entry) {
      this.#redoStack.push(entry);
    }
    return entry;
  }

  redo(): HistoryEntry | undefined {
    const entry = this.#redoStack.pop();
    if (entry) {
      this.#undoStack.push(entry);
    }
    return entry;
  }

  clear(): void {
    this.#undoStack = [];
    this.#redoStack = [];
  }

  #evict(): void {
    while (
      this.#undoStack.length > 1 &&
      (this.#undoStack.length > this.#maxEntries ||
        this.#estimateBytes() > this.#maxBytes)
    ) {
      this.#undoStack.shift();
    }
  }

  #estimateBytes(): number {
    let total = 0;
    for (const entry of this.#undoStack) {
      total += JSON.stringify(entry.patches).length;
    }
    return total;
  }
}
