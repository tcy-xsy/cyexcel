import { describe, expect, it } from "vitest";

import { remapColumnIndex } from "../../../../src/core/layout";

describe("remapColumnIndex — insert", () => {
  it("leaves columns before the insertion point unchanged", () => {
    expect(remapColumnIndex(3, { operation: "insert", start: 5, count: 2 })).toBe(3);
  });

  it("shifts columns at or after the insertion point right by count", () => {
    expect(remapColumnIndex(5, { operation: "insert", start: 5, count: 2 })).toBe(7);
    expect(remapColumnIndex(10, { operation: "insert", start: 5, count: 2 })).toBe(12);
  });
});

describe("remapColumnIndex — delete", () => {
  it("leaves columns before the deleted band unchanged", () => {
    expect(remapColumnIndex(2, { operation: "delete", start: 5, count: 3 })).toBe(2);
  });

  it("collapses columns inside the deleted band onto the deletion point", () => {
    expect(remapColumnIndex(5, { operation: "delete", start: 5, count: 3 })).toBe(5);
    expect(remapColumnIndex(6, { operation: "delete", start: 5, count: 3 })).toBe(5);
    expect(remapColumnIndex(7, { operation: "delete", start: 5, count: 3 })).toBe(5);
  });

  it("shifts columns after the deleted band left by count", () => {
    expect(remapColumnIndex(8, { operation: "delete", start: 5, count: 3 })).toBe(5);
    expect(remapColumnIndex(10, { operation: "delete", start: 5, count: 3 })).toBe(7);
  });
});

describe("remapColumnIndex — move (later in the sheet)", () => {
  // [A,B,C,D,E] -> moveColumns(1, 2, 4) -> [A,D,B,C,E]
  const change = { operation: "move" as const, start: 1, count: 2, target: 4 };

  it("leaves columns before the moved block unchanged", () => {
    expect(remapColumnIndex(0, change)).toBe(0); // A stays at 0
  });

  it("carries columns inside the moved block by their own offset to the new location", () => {
    expect(remapColumnIndex(1, change)).toBe(2); // B: 1 -> 2
    expect(remapColumnIndex(2, change)).toBe(3); // C: 2 -> 3
  });

  it("shifts columns between the old block and the target left by count", () => {
    expect(remapColumnIndex(3, change)).toBe(1); // D: 3 -> 1
  });

  it("leaves columns after the target unchanged", () => {
    expect(remapColumnIndex(4, change)).toBe(4); // E stays at 4
  });
});

describe("remapColumnIndex — move (earlier in the sheet)", () => {
  // [A,B,C,D,E] -> moveColumns(3, 1, 0) -> [D,A,B,C,E]
  const change = { operation: "move" as const, start: 3, count: 1, target: 0 };

  it("carries the moved column itself to the target", () => {
    expect(remapColumnIndex(3, change)).toBe(0); // D: 3 -> 0
  });

  it("shifts columns between the target and the old block right by count", () => {
    expect(remapColumnIndex(0, change)).toBe(1); // A: 0 -> 1
    expect(remapColumnIndex(1, change)).toBe(2); // B: 1 -> 2
    expect(remapColumnIndex(2, change)).toBe(3); // C: 2 -> 3
  });

  it("leaves columns after the old block unchanged", () => {
    expect(remapColumnIndex(4, change)).toBe(4); // E stays at 4
  });
});

describe("remapColumnIndex — move defaults target to start (a no-op move)", () => {
  it("leaves every column exactly where it was", () => {
    const change = { operation: "move" as const, start: 2, count: 2 };
    for (let column = 0; column < 6; column += 1) {
      expect(remapColumnIndex(column, change)).toBe(column);
    }
  });
});
