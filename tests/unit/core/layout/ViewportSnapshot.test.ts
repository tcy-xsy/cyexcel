import { describe, expect, it } from "vitest";

import { ViewportSnapshot } from "../../../../src/core/layout";

describe("ViewportSnapshot", () => {
  it("captures viewport state and overscan as an immutable value object", () => {
    const viewport = new ViewportSnapshot({
      version: 3,
      scrollTop: 120,
      scrollLeft: 40,
      width: 800,
      height: 600,
      overscanRows: 5,
      overscanColumns: 2
    });

    expect(viewport).toEqual({
      version: 3,
      scrollTop: 120,
      scrollLeft: 40,
      width: 800,
      height: 600,
      overscanRows: 5,
      overscanColumns: 2
    });
    expect(Object.isFrozen(viewport)).toBe(true);
  });

  it("defaults overscan to zero", () => {
    const viewport = new ViewportSnapshot({
      version: 0,
      scrollTop: 0,
      scrollLeft: 0,
      width: 1,
      height: 1
    });

    expect(viewport.overscanRows).toBe(0);
    expect(viewport.overscanColumns).toBe(0);
  });

  it("rejects invalid viewport metrics", () => {
    expect(
      () =>
        new ViewportSnapshot({
          version: 0,
          scrollTop: -1,
          scrollLeft: 0,
          width: 1,
          height: 1
        })
    ).toThrow(RangeError);

    expect(
      () =>
        new ViewportSnapshot({
          version: 0,
          scrollTop: 0,
          scrollLeft: 0,
          width: 1,
          height: 1,
          overscanRows: 1.5
        })
    ).toThrow(RangeError);
  });
});
