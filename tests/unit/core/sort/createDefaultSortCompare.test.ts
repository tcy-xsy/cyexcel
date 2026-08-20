import { describe, expect, it } from "vitest";

import { computeSortOrder, createDefaultSortCompare } from "../../../../src/core/sort";

const sign = (value: number): number => (value < 0 ? -1 : value > 0 ? 1 : 0);

describe("createDefaultSortCompare", () => {
  it("compares numbers numerically, not as strings", () => {
    const compare = createDefaultSortCompare("asc");

    // "10" < "9" as strings, so a string comparison would get this backwards.
    expect(sign(compare(9, 10))).toBe(-1);
    expect(sign(compare(10, 9))).toBe(1);
    expect(sign(compare(10, 10))).toBe(0);
  });

  it("compares everything else as strings", () => {
    const compare = createDefaultSortCompare("asc");

    expect(sign(compare("apple", "banana"))).toBe(-1);
    expect(sign(compare("banana", "apple"))).toBe(1);
    expect(sign(compare("apple", "apple"))).toBe(0);
    expect(sign(compare(false, true))).toBe(-1);
  });

  it("falls back to a string comparison when only one side is a number", () => {
    const compare = createDefaultSortCompare("asc");

    // "10" vs "9" as strings.
    expect(sign(compare(10, "9"))).toBe(-1);
  });

  it("treats null and undefined as equal to each other", () => {
    const compare = createDefaultSortCompare("asc");

    expect(sign(compare(null, undefined))).toBe(0);
    expect(sign(compare(undefined, null))).toBe(0);
  });

  it("sorts null and undefined last under asc", () => {
    const keys = [3, null, 1, undefined, 2];
    const order = computeSortOrder({
      rowCount: keys.length,
      getSortKey: (row) => keys[row],
      compare: createDefaultSortCompare("asc"),
      direction: "asc"
    });

    expect([...order!]).toEqual([2, 4, 0, 1, 3]);
  });

  it("sorts null and undefined last under desc too", () => {
    const keys = [3, null, 1, undefined, 2];
    const order = computeSortOrder({
      rowCount: keys.length,
      getSortKey: (row) => keys[row],
      compare: createDefaultSortCompare("desc"),
      direction: "desc"
    });

    // Nullish stays at the bottom; only the defined keys reverse. The two
    // nullish rows keep their original relative order (1 before 3).
    expect([...order!]).toEqual([0, 4, 2, 1, 3]);
  });
});

describe("createDefaultSortCompare with a cell type's comparator", () => {
  // Chronological, like the "date" built-in: the string order of these is
  // Mon < Sat < Wed, so the fallback could never produce it.
  const wedJan01 = new Date(Date.UTC(2020, 0, 1, 12));
  const monJan06 = new Date(Date.UTC(2020, 0, 6, 12));
  const satJan11 = new Date(Date.UTC(2020, 0, 11, 12));
  const chronological = (left: unknown, right: unknown): number =>
    (left as Date).getTime() - (right as Date).getTime();

  it("delegates two defined values to it instead of comparing them as strings", () => {
    const compare = createDefaultSortCompare("asc", chronological);

    expect(sign(compare(wedJan01, monJan06))).toBe(-1);
    expect(sign(compare(satJan11, wedJan01))).toBe(1);
    expect(sign(compare(wedJan01, wedJan01))).toBe(0);
  });

  it("never hands it a null or undefined side", () => {
    const seen: unknown[] = [];
    const compare = createDefaultSortCompare("asc", (left, right) => {
      seen.push(left, right);
      return chronological(left, right);
    });

    compare(null, wedJan01);
    compare(wedJan01, undefined);
    compare(undefined, null);

    expect(seen).toEqual([]);
  });

  it("still keeps empty cells last in both directions", () => {
    const keys = [satJan11, null, wedJan01, undefined, monJan06];
    const request = {
      rowCount: keys.length,
      getSortKey: (row: number) => keys[row]
    };

    expect([
      ...computeSortOrder({
        ...request,
        compare: createDefaultSortCompare("asc", chronological),
        direction: "asc"
      })!
    ]).toEqual([2, 4, 0, 1, 3]);

    // The cell type orders the real dates; the empties stay at the bottom
    // rather than flipping to the top, which is what a bare
    // `column.compare` would do here.
    expect([
      ...computeSortOrder({
        ...request,
        compare: createDefaultSortCompare("desc", chronological),
        direction: "desc"
      })!
    ]).toEqual([0, 4, 2, 1, 3]);
  });
});
