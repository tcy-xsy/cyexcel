import { describe, expect, it } from "vitest";

import { splitByFrozenBoundary } from "../../../../src/core/layout";

describe("splitByFrozenBoundary", () => {
  it("puts everything in the scrollable half when nothing is frozen", () => {
    const result = splitByFrozenBoundary({ start: 0, count: 10 }, 0);

    expect(result.frozen).toEqual({ start: 0, count: 0 });
    expect(result.scrollable).toEqual({ start: 0, count: 10 });
  });

  it("splits a range straddling the frozen boundary into both halves", () => {
    const result = splitByFrozenBoundary({ start: 0, count: 10 }, 3);

    expect(result.frozen).toEqual({ start: 0, count: 3 });
    expect(result.scrollable).toEqual({ start: 3, count: 7 });
  });

  it("returns an empty frozen half when the range starts after the frozen boundary", () => {
    const result = splitByFrozenBoundary({ start: 5, count: 3 }, 3);

    expect(result.frozen).toEqual({ start: 3, count: 0 });
    expect(result.scrollable).toEqual({ start: 5, count: 3 });
  });

  it("returns an empty scrollable half when the range is entirely inside the frozen band", () => {
    const result = splitByFrozenBoundary({ start: 0, count: 2 }, 5);

    expect(result.frozen).toEqual({ start: 0, count: 2 });
    expect(result.scrollable).toEqual({ start: 5, count: 0 });
  });

  it("handles a frozenCount of 0 and an empty input range without throwing", () => {
    const result = splitByFrozenBoundary({ start: 0, count: 0 }, 0);

    expect(result.frozen).toEqual({ start: 0, count: 0 });
    expect(result.scrollable).toEqual({ start: 0, count: 0 });
  });
});
