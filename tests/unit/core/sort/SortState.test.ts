import { describe, expect, it } from "vitest";

import { nextSortDirection } from "../../../../src/core/sort";

describe("nextSortDirection", () => {
  it("starts a never-sorted column at ascending", () => {
    expect(nextSortDirection(null, 0)).toBe("asc");
  });

  it("cycles the same column from ascending to descending", () => {
    expect(nextSortDirection({ column: 2, direction: "asc" }, 2)).toBe("desc");
  });

  it("cycles the same column from descending back to ascending", () => {
    expect(nextSortDirection({ column: 2, direction: "desc" }, 2)).toBe("asc");
  });

  it("never returns to a third 'unsorted' state on repeated clicks", () => {
    let state: { column: number; direction: "asc" | "desc" } | null = null;
    const seen = new Set<string>();
    for (let i = 0; i < 5; i += 1) {
      const direction = nextSortDirection(state, 1);
      seen.add(direction);
      state = { column: 1, direction };
    }
    expect(seen).toEqual(new Set(["asc", "desc"]));
  });

  it("starts a different column fresh at ascending, even mid-descending on the old column", () => {
    expect(nextSortDirection({ column: 0, direction: "desc" }, 1)).toBe("asc");
  });

  it("starts a different column fresh at ascending when the old column was ascending", () => {
    expect(nextSortDirection({ column: 0, direction: "asc" }, 1)).toBe("asc");
  });
});
