import { describe, expect, it } from "vitest";

import { computeSortOrder, createDefaultSortCompare } from "../../../../src/core/sort";
import type { CellValue } from "../../../../src/core/sheet/types";

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

  it("treats two empty cells as equal to each other", () => {
    const compare = createDefaultSortCompare("asc");

    expect(sign(compare(undefined, undefined))).toBe(0);
  });

  it("still guards null, which only an untyped host can hand in", () => {
    // `CellValue` has no `null` (see sheet/types.ts), so this can arrive only
    // from a JS caller. The guard is what makes it degrade to "empty" rather
    // than be stringified into "null" and sort in between "n" and "o".
    const compare = createDefaultSortCompare("asc") as (left: unknown, right: unknown) => number;

    expect(sign(compare(null, undefined))).toBe(0);
    expect(sign(compare(null, "apple"))).toBe(1);
    expect(sign(compare("apple", null))).toBe(-1);
  });

  it("sorts empty cells last under asc", () => {
    const keys: (CellValue | undefined)[] = [3, undefined, 1, undefined, 2];
    const order = computeSortOrder({
      rowCount: keys.length,
      getSortKey: (row) => keys[row],
      compare: createDefaultSortCompare("asc"),
      direction: "asc"
    });

    expect([...order!]).toEqual([2, 4, 0, 1, 3]);
  });

  it("sorts empty cells last under desc too", () => {
    const keys: (CellValue | undefined)[] = [3, undefined, 1, undefined, 2];
    const order = computeSortOrder({
      rowCount: keys.length,
      getSortKey: (row) => keys[row],
      compare: createDefaultSortCompare("desc"),
      direction: "desc"
    });

    // Empties stay at the bottom; only the defined keys reverse. The two
    // empty rows keep their original relative order (1 before 3).
    expect([...order!]).toEqual([0, 4, 2, 1, 3]);
  });
});

describe("createDefaultSortCompare with a cell type's comparator", () => {
  // Case-insensitive, like a text column that wants "apple" next to "Apple":
  // the codepoint order of these three is Banana < Cherry < apple (uppercase
  // sorts first), so the fallback could never produce the order below.
  //
  // 不再用 `Date` 举例：日期现在是序列号（数字），默认的数字比较天然就是时间序，
  // 那个例子被数据模型消解掉了——见 cells/builtins.ts 的 "date"。
  const chronologicalStandIn = (left: CellValue, right: CellValue): number => {
    const leftText = String(left).toLowerCase();
    const rightText = String(right).toLowerCase();
    return leftText < rightText ? -1 : leftText > rightText ? 1 : 0;
  };
  const apple = "apple";
  const banana = "Banana";
  const cherry = "Cherry";

  it("delegates two defined values to it instead of comparing them as strings", () => {
    const compare = createDefaultSortCompare("asc", chronologicalStandIn);

    // Premise: the codepoint order really does disagree, so these can only
    // pass through the supplied comparator.
    expect(sign(createDefaultSortCompare("asc")(apple, banana))).toBe(1);

    expect(sign(compare(apple, banana))).toBe(-1);
    expect(sign(compare(cherry, apple))).toBe(1);
    expect(sign(compare(apple, apple))).toBe(0);
  });

  it("never hands it an empty side", () => {
    const seen: CellValue[] = [];
    const compare = createDefaultSortCompare("asc", (left, right) => {
      seen.push(left, right);
      return chronologicalStandIn(left, right);
    });

    compare(undefined, apple);
    compare(apple, undefined);
    compare(undefined, undefined);

    expect(seen).toEqual([]);
  });

  it("still keeps empty cells last in both directions", () => {
    const keys: (CellValue | undefined)[] = [cherry, undefined, apple, undefined, banana];
    const request = {
      rowCount: keys.length,
      getSortKey: (row: number) => keys[row]
    };

    expect([
      ...computeSortOrder({
        ...request,
        compare: createDefaultSortCompare("asc", chronologicalStandIn),
        direction: "asc"
      })!
    ]).toEqual([2, 4, 0, 1, 3]);

    // The cell type orders the real values; the empties stay at the bottom
    // rather than flipping to the top, which is what a bare
    // `column.compare` would do here.
    expect([
      ...computeSortOrder({
        ...request,
        compare: createDefaultSortCompare("desc", chronologicalStandIn),
        direction: "desc"
      })!
    ]).toEqual([0, 4, 2, 1, 3]);
  });
});
