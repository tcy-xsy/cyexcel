import { describe, expect, it } from "vitest";

import { toPhysicalPixelSize } from "../../../../src/core/layout";

describe("toPhysicalPixelSize", () => {
  it("scales the logical size by dpr", () => {
    expect(toPhysicalPixelSize(100, 1)).toBe(100);
    expect(toPhysicalPixelSize(100, 2)).toBe(200);
  });

  it("caps dpr at maxDpr", () => {
    expect(toPhysicalPixelSize(100, 5, 3)).toBe(300);
    expect(toPhysicalPixelSize(100, 2, 3)).toBe(200);
  });

  it("uses a default maxDpr of 3 when not provided", () => {
    expect(toPhysicalPixelSize(100, 10)).toBe(300);
  });

  it("rounds the result", () => {
    expect(toPhysicalPixelSize(101, 1.5)).toBe(Math.round(101 * 1.5));
  });
});
