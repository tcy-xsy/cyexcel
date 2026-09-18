import type { SearchMatch } from "./computeMatches";

/**
 * Tracks the current match set and which one is "active" (the one
 * `findNext()`/`findPrevious()` last landed on). Kept separate from
 * `CyExcelCore` — same split as `computeSortOrder`/`SortState` vs.
 * `sortByColumn()` — because the navigation logic (wrap-around index math)
 * has nothing to do with the canvas/DOM wiring and is worth testing on its
 * own, without a real `CyExcelCore` instance.
 */
export class SearchState {
  #matches: readonly SearchMatch[] = [];
  #activeIndex = -1;

  get matches(): readonly SearchMatch[] {
    return this.#matches;
  }

  get activeMatch(): SearchMatch | null {
    return this.#matches[this.#activeIndex] ?? null;
  }

  /** Replaces the match set and deactivates — a fresh search has no "current" match yet. */
  setMatches(matches: readonly SearchMatch[]): void {
    this.#matches = matches;
    this.#activeIndex = -1;
  }

  next(): SearchMatch | null {
    if (this.#matches.length === 0) {
      return null;
    }
    this.#activeIndex = (this.#activeIndex + 1) % this.#matches.length;
    return this.activeMatch;
  }

  previous(): SearchMatch | null {
    if (this.#matches.length === 0) {
      return null;
    }
    // -1 (nothing active yet) wraps to the last match, not "index -2".
    this.#activeIndex =
      this.#activeIndex <= 0 ? this.#matches.length - 1 : this.#activeIndex - 1;
    return this.activeMatch;
  }

  clear(): void {
    this.#matches = [];
    this.#activeIndex = -1;
  }
}
