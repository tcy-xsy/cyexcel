import { describe, expect, it } from "vitest";

import {
  CAPSULE_HEIGHT,
  CAPSULE_ICON_PLACE,
  CAPSULE_MARGIN_H,
  CAPSULE_PADDING_H,
  computeCapsule,
  isInDropdownArrowZone,
  isInDropdownHitArea
} from "../../../../src/core/render/dropdownGeometry";
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

describe("computeCapsule", () => {
  const cell = { x: 0, y: 0, width: 100, height: 24 };

  it("hugs the left edge, is as wide as padding + text + arrow + padding, and is vertically centred", () => {
    const geometry = computeCapsule(cell, 30)!;

    expect(geometry.capsule).toEqual({
      x: CAPSULE_MARGIN_H,
      y: (24 - CAPSULE_HEIGHT) / 2,
      width: CAPSULE_PADDING_H + 30 + CAPSULE_ICON_PLACE + CAPSULE_PADDING_H,
      height: CAPSULE_HEIGHT
    });
    expect(geometry.textX).toBe(CAPSULE_MARGIN_H + CAPSULE_PADDING_H);
    expect(geometry.textMaxWidth).toBe(30);
    // 箭头在胶囊右端、垂直居中
    expect(geometry.arrowCenter).toEqual({
      x: CAPSULE_MARGIN_H + geometry.capsule.width - CAPSULE_PADDING_H - CAPSULE_ICON_PLACE / 2,
      y: 12
    });
  });

  it("stops growing at the cell width minus both margins, and shrinks the text area instead", () => {
    const geometry = computeCapsule(cell, 500)!;
    expect(geometry.capsule.width).toBe(100 - CAPSULE_MARGIN_H * 2);
    expect(geometry.textMaxWidth).toBe(100 - CAPSULE_MARGIN_H * 2 - CAPSULE_PADDING_H * 2 - CAPSULE_ICON_PLACE);
  });

  it("is just the arrow for an empty value", () => {
    const geometry = computeCapsule(cell, 0)!;
    expect(geometry.capsule.width).toBe(CAPSULE_PADDING_H * 2 + CAPSULE_ICON_PLACE);
    expect(geometry.textMaxWidth).toBe(0);
  });

  it("does not grow taller than a line even in a tall cell", () => {
    expect(computeCapsule({ x: 0, y: 0, width: 100, height: 80 }, 10)!.capsule.height).toBe(CAPSULE_HEIGHT);
  });

  it("gives up (null) when the cell is too narrow or too short", () => {
    expect(computeCapsule({ x: 0, y: 0, width: 30, height: 24 }, 10)).toBeNull();
    expect(computeCapsule({ x: 0, y: 0, width: 100, height: 6 }, 10)).toBeNull();
  });
});

describe("isInDropdownHitArea", () => {
  const cell = { x: 0, y: 0, width: 100, height: 24 };

  it("text mode never hits", () => {
    expect(isInDropdownHitArea({ mode: "text" }, cell, 10, { x: 95, y: 12 })).toBe(false);
  });

  it("arrow mode hits the right-hand zone", () => {
    expect(isInDropdownHitArea({ mode: "arrow" }, cell, 10, { x: 95, y: 12 })).toBe(true);
    expect(isInDropdownHitArea({ mode: "arrow" }, cell, 10, { x: 20, y: 12 })).toBe(false);
  });

  it("custom mode hits the whole capsule — text and arrow alike — but not the empty part of the cell", () => {
    const capsule = computeCapsule(cell, 30)!.capsule;
    const custom = { mode: "custom" } as const;
    expect(isInDropdownHitArea(custom, cell, 30, { x: capsule.x + 2, y: 12 })).toBe(true); // 文字一侧
    expect(isInDropdownHitArea(custom, cell, 30, { x: capsule.x + capsule.width - 2, y: 12 })).toBe(true); // 箭头一侧
    expect(isInDropdownHitArea(custom, cell, 30, { x: 95, y: 12 })).toBe(false); // 胶囊右边的空白
    expect(isInDropdownHitArea(custom, cell, 30, { x: capsule.x + 2, y: 0 })).toBe(false); // 胶囊上方
  });

  it("custom mode falls back to the arrow zone when the capsule does not fit", () => {
    const narrow = { x: 0, y: 0, width: 30, height: 24 };
    expect(isInDropdownHitArea({ mode: "custom" }, narrow, 10, { x: 25, y: 12 })).toBe(true);
    expect(isInDropdownHitArea({ mode: "custom" }, narrow, 10, { x: 2, y: 12 })).toBe(false);
  });
});
