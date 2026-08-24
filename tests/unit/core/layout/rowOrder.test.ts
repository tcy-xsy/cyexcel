import { describe, expect, it } from "vitest";

import { applyRowOrder, assertValidRowOrder, invertRowOrder } from "../../../../src/core/layout";

describe("assertValidRowOrder", () => {
  it("accepts a permutation of exactly the expected length", () => {
    expect(() => assertValidRowOrder(Int32Array.from([2, 0, 1]), 3)).not.toThrow();
  });

  it("rejects a length mismatch", () => {
    expect(() => assertValidRowOrder(Int32Array.from([0, 1]), 3)).toThrow(/length/);
  });

  it("rejects a repeated index, which would duplicate one row and drop another", () => {
    expect(() => assertValidRowOrder(Int32Array.from([0, 0, 2]), 3)).toThrow(/not a permutation/);
  });

  it("rejects an out-of-range index", () => {
    expect(() => assertValidRowOrder(Int32Array.from([0, 1, 3]), 3)).toThrow(/not a permutation/);
  });
});

describe("applyRowOrder", () => {
  it("reads order[newPosition] as the old position", () => {
    // order[0] === 2 means "the row now first is the one that was third".
    expect(applyRowOrder(["a", "b", "c"], Int32Array.from([2, 0, 1]))).toEqual(["c", "a", "b"]);
  });
});

describe("invertRowOrder", () => {
  it("produces the order that undoes the original", () => {
    const order = Int32Array.from([2, 0, 1]);
    const inverse = invertRowOrder(order);

    expect(applyRowOrder(applyRowOrder(["a", "b", "c"], order), inverse)).toEqual(["a", "b", "c"]);
  });

  it("is not equal to the original for an asymmetric permutation", () => {
    // Guards the classic bug: an inverse implemented as a copy passes the
    // round-trip test only for self-inverse permutations.
    expect([...invertRowOrder(Int32Array.from([2, 0, 1]))]).toEqual([1, 2, 0]);
  });
});
