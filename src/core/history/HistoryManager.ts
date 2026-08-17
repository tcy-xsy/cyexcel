import type { CommitReceipt } from "../mutations";

export type HistoryEntry = CommitReceipt;

export interface HistoryManagerOptions {
  readonly maxEntries?: number | undefined;
  readonly maxBytes?: number | undefined;
}

const DEFAULT_MAX_ENTRIES = 100;
const DEFAULT_MAX_BYTES = 5_000_000;

/** Receipt-only history. Stack ownership changes only after replay succeeds. */
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

  peekUndo(): HistoryEntry | undefined {
    return this.#undoStack[this.#undoStack.length - 1];
  }

  peekRedo(): HistoryEntry | undefined {
    return this.#redoStack[this.#redoStack.length - 1];
  }

  commitUndo(entry: HistoryEntry): void {
    if (this.peekUndo() !== entry) {
      throw new Error("Undo history changed before replay completed");
    }
    this.#undoStack.pop();
    this.#redoStack.push(entry);
  }

  commitRedo(entry: HistoryEntry): void {
    if (this.peekRedo() !== entry) {
      throw new Error("Redo history changed before replay completed");
    }
    this.#redoStack.pop();
    this.#undoStack.push(entry);
  }

  clear(): void {
    this.#undoStack = [];
    this.#redoStack = [];
  }

  #evict(): void {
    let bytes = this.#undoStack.reduce((total, entry) => total + entry.historyBytes, 0);
    while (
      this.#undoStack.length > 1 &&
      (this.#undoStack.length > this.#maxEntries || bytes > this.#maxBytes)
    ) {
      const removed = this.#undoStack.shift();
      bytes -= removed?.historyBytes ?? 0;
    }
  }
}
