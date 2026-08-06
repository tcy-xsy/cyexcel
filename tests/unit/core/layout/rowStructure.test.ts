import { describe, expect, it } from "vitest";

import { remapRowIndex } from "../../../../src/core";

describe("remapRowIndex — insert", () => {
  it("leaves rows before the insertion point unchanged", () => {
    expect(remapRowIndex(3, { operation: "insert", start: 5, count: 2 })).toBe(3);
  });

  it("shifts rows at or after the insertion point down by count", () => {
    expect(remapRowIndex(5, { operation: "insert", start: 5, count: 2 })).toBe(7);
    expect(remapRowIndex(10, { operation: "insert", start: 5, count: 2 })).toBe(12);
  });
});

describe("remapRowIndex — delete", () => {
  it("leaves rows before the deleted band unchanged", () => {
    expect(remapRowIndex(2, { operation: "delete", start: 5, count: 3 })).toBe(2);
  });

  it("collapses rows inside the deleted band onto the deletion point", () => {
    expect(remapRowIndex(5, { operation: "delete", start: 5, count: 3 })).toBe(5);
    expect(remapRowIndex(6, { operation: "delete", start: 5, count: 3 })).toBe(5);
    expect(remapRowIndex(7, { operation: "delete", start: 5, count: 3 })).toBe(5);
  });

  it("shifts rows after the deleted band up by count", () => {
    expect(remapRowIndex(8, { operation: "delete", start: 5, count: 3 })).toBe(5);
    expect(remapRowIndex(10, { operation: "delete", start: 5, count: 3 })).toBe(7);
  });
});

describe("remapRowIndex — move (later in the sheet)", () => {
  // [A,B,C,D,E] -> moveRows(1, 2, 4) -> [A,D,B,C,E]
  const change = { operation: "move" as const, start: 1, count: 2, target: 4 };

  it("leaves rows before the moved block unchanged", () => {
    expect(remapRowIndex(0, change)).toBe(0); // A stays at 0
  });

  it("carries rows inside the moved block by their own offset to the new location", () => {
    expect(remapRowIndex(1, change)).toBe(2); // B: 1 -> 2
    expect(remapRowIndex(2, change)).toBe(3); // C: 2 -> 3
  });

  it("shifts rows between the old block and the target up by count", () => {
    expect(remapRowIndex(3, change)).toBe(1); // D: 3 -> 1
  });

  it("leaves rows after the target unchanged", () => {
    expect(remapRowIndex(4, change)).toBe(4); // E stays at 4
  });
});

describe("remapRowIndex — move (earlier in the sheet)", () => {
  // [A,B,C,D,E] -> moveRows(3, 1, 0) -> [D,A,B,C,E]
  const change = { operation: "move" as const, start: 3, count: 1, target: 0 };

  it("carries the moved row itself to the target", () => {
    expect(remapRowIndex(3, change)).toBe(0); // D: 3 -> 0
  });

  it("shifts rows between the target and the old block down by count", () => {
    expect(remapRowIndex(0, change)).toBe(1); // A: 0 -> 1
    expect(remapRowIndex(1, change)).toBe(2); // B: 1 -> 2
    expect(remapRowIndex(2, change)).toBe(3); // C: 2 -> 3
  });

  it("leaves rows after the old block unchanged", () => {
    expect(remapRowIndex(4, change)).toBe(4); // E stays at 4
  });
});

describe("remapRowIndex — move a multi-row block", () => {
  it("preserves the internal order of a moved block larger than one row", () => {
    // [A,B,C,D,E,F] -> moveRows(0, 3, 6) -> [D,E,F,A,B,C]
    const change = { operation: "move" as const, start: 0, count: 3, target: 6 };
    expect(remapRowIndex(0, change)).toBe(3); // A
    expect(remapRowIndex(1, change)).toBe(4); // B
    expect(remapRowIndex(2, change)).toBe(5); // C
    expect(remapRowIndex(3, change)).toBe(0); // D
    expect(remapRowIndex(4, change)).toBe(1); // E
    expect(remapRowIndex(5, change)).toBe(2); // F
  });
});

describe("remapRowIndex — move defaults target to start (a no-op move)", () => {
  it("leaves every row exactly where it was", () => {
    const change = { operation: "move" as const, start: 2, count: 2 };
    for (let row = 0; row < 6; row += 1) {
      expect(remapRowIndex(row, change)).toBe(row);
    }
  });
});
