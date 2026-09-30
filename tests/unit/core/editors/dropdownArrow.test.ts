import { describe, expect, it } from "vitest";

import { isInDropdownArrowZone } from "../../../../src/core/editors";
import { DROPDOWN_ARROW_ZONE } from "../../../../src/core/render/theme";

const rect = { x: 100, y: 40, width: 120, height: 24 };

describe("isInDropdownArrowZone", () => {
  it("hits inside the rightmost DROPDOWN_ARROW_ZONE pixels of the cell, full height", () => {
    const left = rect.x + rect.width - DROPDOWN_ARROW_ZONE;
    expect(isInDropdownArrowZone(rect, { x: left, y: rect.y })).toBe(true);
    expect(isInDropdownArrowZone(rect, { x: rect.x + rect.width, y: rect.y + rect.height })).toBe(true);
    expect(isInDropdownArrowZone(rect, { x: rect.x + rect.width - 3, y: rect.y + 12 })).toBe(true);
  });

  it("misses everywhere else in the cell and outside it", () => {
    const left = rect.x + rect.width - DROPDOWN_ARROW_ZONE;
    expect(isInDropdownArrowZone(rect, { x: left - 1, y: rect.y + 12 })).toBe(false);
    expect(isInDropdownArrowZone(rect, { x: rect.x + 10, y: rect.y + 12 })).toBe(false);
    expect(isInDropdownArrowZone(rect, { x: rect.x + rect.width + 1, y: rect.y + 12 })).toBe(false);
    expect(isInDropdownArrowZone(rect, { x: rect.x + rect.width - 3, y: rect.y - 1 })).toBe(false);
    expect(isInDropdownArrowZone(rect, { x: rect.x + rect.width - 3, y: rect.y + rect.height + 1 })).toBe(false);
  });

  it("uses the whole cell when it is narrower than the zone", () => {
    const narrow = { x: 0, y: 0, width: 10, height: 10 };
    expect(isInDropdownArrowZone(narrow, { x: 0, y: 5 })).toBe(true);
    expect(isInDropdownArrowZone(narrow, { x: 11, y: 5 })).toBe(false);
  });
});
