import { describe, expect, it, vi } from "vitest";

import { computeSortOrder } from "../../../../src/core/sort";

const numbers = (values: readonly number[]) => ({
  rowCount: values.length,
  getSortKey: (row: number) => values[row],
  compare: (left: unknown, right: unknown) => (left as number) - (right as number)
});

describe("computeSortOrder", () => {
  it("produces order[newPosition] = oldPosition for ascending", () => {
    expect([...computeSortOrder({ ...numbers([30, 10, 20]), direction: "asc" })!]).toEqual([1, 2, 0]);
  });

  it("reverses for descending", () => {
    expect([...computeSortOrder({ ...numbers([30, 10, 20]), direction: "desc" })!]).toEqual([0, 2, 1]);
  });

  it("is stable: equal keys keep their original relative order", () => {
    const order = computeSortOrder({ ...numbers([5, 1, 5, 1]), direction: "asc" })!;
    // The two 1s keep 1 before 3; the two 5s keep 0 before 2.
    expect([...order]).toEqual([1, 3, 0, 2]);
  });

  it("is stable in descending: equal keys keep their original relative order", () => {
    const order = computeSortOrder({ ...numbers([5, 1, 5, 1]), direction: "desc" })!;
    // The two 5s keep 0 before 2; the two 1s keep 1 before 3.
    expect([...order]).toEqual([0, 2, 1, 3]);
  });

  it("returns null when the signal is already aborted", () => {
    const controller = new AbortController();
    controller.abort();

    expect(computeSortOrder({ ...numbers([3, 1, 2]), direction: "asc", signal: controller.signal })).toBeNull();
  });

  it("reports progress at least once and finishes at 1", () => {
    const onProgress = vi.fn();
    computeSortOrder({ ...numbers([3, 1, 2]), direction: "asc", onProgress });

    expect(onProgress).toHaveBeenCalled();
    expect(onProgress.mock.calls.at(-1)?.[0]).toBe(1);
  });

  it("reads each row's sort key exactly once", () => {
    // Sorting a million rows must not call a user-supplied formatter O(n log n)
    // times; keys are extracted up front.
    const getSortKey = vi.fn((row: number) => [30, 10, 20][row]);
    computeSortOrder({ rowCount: 3, getSortKey, compare: (a, b) => (a as number) - (b as number), direction: "asc" });

    expect(getSortKey).toHaveBeenCalledTimes(3);
  });

  it("returns null when aborted during key extraction", () => {
    // Use enough rows that sort's compare would definitely be invoked.
    const values = [9, 8, 7, 6, 5, 4, 3, 2, 1, 0];
    const controller = new AbortController();
    const getSortKey = (row: number) => {
      if (row === values.length - 1) {
        controller.abort();
      }
      return values[row];
    };
    const compare = vi.fn((a: unknown, b: unknown) => (a as number) - (b as number));

    // First, verify compare IS called in the non-aborted case.
    computeSortOrder({
      rowCount: values.length,
      getSortKey: (row: number) => values[row],
      compare,
      direction: "asc"
    });
    const compareCalls = compare.mock.calls.length;
    expect(compareCalls).toBeGreaterThan(0);

    // Now test the aborted case: compare should not be called because
    // the post-extraction checkpoint prevents sort from running.
    const compareSpyAborted = vi.fn((a: unknown, b: unknown) => (a as number) - (b as number));
    const result = computeSortOrder({
      rowCount: values.length,
      getSortKey,
      compare: compareSpyAborted,
      direction: "asc",
      signal: controller.signal
    });

    expect(result).toBeNull();
    expect(compareSpyAborted).not.toHaveBeenCalled();
  });

  it("returns null when aborted during sort", () => {
    const controller = new AbortController();
    let compareCalls = 0;
    const compare = (left: unknown, right: unknown) => {
      compareCalls += 1;
      if (compareCalls === 2) {
        controller.abort();
      }
      return (left as number) - (right as number);
    };
    const result = computeSortOrder({
      ...numbers([30, 10, 20]),
      direction: "asc",
      compare,
      signal: controller.signal
    });

    expect(result).toBeNull();
  });
});
