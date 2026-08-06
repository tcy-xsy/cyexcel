import { describe, expect, it } from "vitest";

import { toViewportRect } from "../../../../src/core";

describe("toViewportRect", () => {
  it("subtracts the current scroll offset from an absolute content rect", () => {
    const rect = { x: 50, y: 60, width: 100, height: 20 };
    const viewport = { scrollTop: 45, scrollLeft: 30, width: 200, height: 100 };

    expect(toViewportRect(rect, viewport)).toEqual({ x: 20, y: 15, width: 100, height: 20 });
  });

  it("is a no-op at scrollTop/scrollLeft 0", () => {
    const rect = { x: 50, y: 60, width: 100, height: 20 };
    const viewport = { scrollTop: 0, scrollLeft: 0, width: 200, height: 100 };

    expect(toViewportRect(rect, viewport)).toEqual(rect);
  });

  it("leaves a frozen row's y unshifted while still translating x", () => {
    const rect = { x: 50, y: 60, width: 100, height: 20 };
    const viewport = { scrollTop: 45, scrollLeft: 30, width: 200, height: 100 };

    expect(toViewportRect(rect, viewport, { row: true })).toEqual({
      x: 20,
      y: 60,
      width: 100,
      height: 20
    });
  });

  it("leaves a frozen column's x unshifted while still translating y", () => {
    const rect = { x: 50, y: 60, width: 100, height: 20 };
    const viewport = { scrollTop: 45, scrollLeft: 30, width: 200, height: 100 };

    expect(toViewportRect(rect, viewport, { column: true })).toEqual({
      x: 50,
      y: 15,
      width: 100,
      height: 20
    });
  });

  it("leaves both axes unshifted for the frozen corner", () => {
    const rect = { x: 50, y: 60, width: 100, height: 20 };
    const viewport = { scrollTop: 45, scrollLeft: 30, width: 200, height: 100 };

    expect(toViewportRect(rect, viewport, { row: true, column: true })).toEqual(rect);
  });
});
