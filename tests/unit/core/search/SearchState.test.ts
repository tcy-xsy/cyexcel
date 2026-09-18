import { describe, expect, it } from "vitest";

import { SearchState } from "../../../../src/core/search";

describe("SearchState", () => {
  it("starts with no matches and no active match", () => {
    const state = new SearchState();
    expect(state.matches).toEqual([]);
    expect(state.activeMatch).toBeNull();
  });

  it("setMatches resets the active pointer so nothing is active yet", () => {
    const state = new SearchState();
    state.setMatches([{ row: 0, column: 0 }, { row: 1, column: 0 }]);
    expect(state.matches).toEqual([{ row: 0, column: 0 }, { row: 1, column: 0 }]);
    expect(state.activeMatch).toBeNull();
  });

  it("next() advances through matches in order and wraps around", () => {
    const state = new SearchState();
    state.setMatches([{ row: 0, column: 0 }, { row: 1, column: 0 }, { row: 2, column: 0 }]);

    expect(state.next()).toEqual({ row: 0, column: 0 });
    expect(state.next()).toEqual({ row: 1, column: 0 });
    expect(state.next()).toEqual({ row: 2, column: 0 });
    expect(state.next()).toEqual({ row: 0, column: 0 });
  });

  it("previous() walks backward and wraps around", () => {
    const state = new SearchState();
    state.setMatches([{ row: 0, column: 0 }, { row: 1, column: 0 }, { row: 2, column: 0 }]);

    expect(state.previous()).toEqual({ row: 2, column: 0 });
    expect(state.previous()).toEqual({ row: 1, column: 0 });
    expect(state.previous()).toEqual({ row: 0, column: 0 });
    expect(state.previous()).toEqual({ row: 2, column: 0 });
  });

  it("next() and previous() return null when there are no matches", () => {
    const state = new SearchState();
    state.setMatches([]);
    expect(state.next()).toBeNull();
    expect(state.previous()).toBeNull();
  });

  it("activeMatch reflects the last cell next()/previous() landed on", () => {
    const state = new SearchState();
    state.setMatches([{ row: 0, column: 0 }, { row: 1, column: 0 }]);
    state.next();
    expect(state.activeMatch).toEqual({ row: 0, column: 0 });
    state.next();
    expect(state.activeMatch).toEqual({ row: 1, column: 0 });
  });

  it("clear() empties matches and the active pointer", () => {
    const state = new SearchState();
    state.setMatches([{ row: 0, column: 0 }]);
    state.next();

    state.clear();

    expect(state.matches).toEqual([]);
    expect(state.activeMatch).toBeNull();
  });
});
