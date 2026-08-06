import { describe, expect, it } from "vitest";

import { computeFillValues } from "../../../../src/core/interaction/fillValues";

describe("computeFillValues", () => {
  it("returns an empty array when targetCount is 0", () => {
    expect(computeFillValues([1, 2, 3], 0)).toEqual([]);
  });

  it("fills with undefined when there are no known source values", () => {
    expect(computeFillValues([], 3)).toEqual([undefined, undefined, undefined]);
  });

  it("copies a single known value for every new slot", () => {
    expect(computeFillValues(["x"], 4)).toEqual(["x", "x", "x", "x"]);
  });

  it("linearly extrapolates a purely-numeric sequence", () => {
    expect(computeFillValues([1, 2, 3], 3)).toEqual([4, 5, 6]);
  });

  it("extrapolates with a negative step", () => {
    expect(computeFillValues([10, 8, 6], 2)).toEqual([4, 2]);
  });

  it("averages the step for a non-evenly-spaced numeric sequence", () => {
    // first=1, last=10, length=3 -> step=(10-1)/2=4.5
    expect(computeFillValues([1, 5, 10], 2)).toEqual([14.5, 19]);
  });

  it("cycles a non-numeric sequence rather than extrapolating", () => {
    expect(computeFillValues(["a", "b", "c"], 4)).toEqual(["a", "b", "c", "a"]);
  });

  it("cycles a mixed number/string sequence rather than extrapolating", () => {
    expect(computeFillValues([1, "b"], 3)).toEqual([1, "b", 1]);
  });
});
