import { describe, expect, it, vi } from "vitest";

import { computeMatches } from "../../../../src/core/search";
import type { CellValue } from "../../../../src/core/sheet/types";

interface Cell {
  readonly row: number;
  readonly column: number;
  readonly value: CellValue | undefined;
}

const cells = (values: readonly [number, number, CellValue][]): readonly Cell[] =>
  values.map(([row, column, value]) => ({ row, column, value }));

const asText = (_row: number, _column: number, value: CellValue | undefined) => String(value);

describe("computeMatches", () => {
  it("finds substring matches, case-insensitive by default", async () => {
    const occupiedCells = cells([
      [0, 0, "Hello World"],
      [1, 0, "goodbye"],
      [2, 0, "hello there"]
    ]);

    const matches = await computeMatches({
      occupiedCells,
      getDisplayText: asText,
      query: "hello"
    });

    expect(matches).toEqual([
      { row: 0, column: 0 },
      { row: 2, column: 0 }
    ]);
  });

  it("respects matchCase: true", async () => {
    const occupiedCells = cells([
      [0, 0, "Hello"],
      [1, 0, "hello"]
    ]);

    const matches = await computeMatches({
      occupiedCells,
      getDisplayText: asText,
      query: "hello",
      matchCase: true
    });

    expect(matches).toEqual([{ row: 1, column: 0 }]);
  });

  it("treats query as a regular expression when useRegex is true", async () => {
    const occupiedCells = cells([
      [0, 0, "foo123"],
      [1, 0, "foobar"]
    ]);

    const matches = await computeMatches({
      occupiedCells,
      getDisplayText: asText,
      query: "foo\\d+",
      useRegex: true
    });

    expect(matches).toEqual([{ row: 0, column: 0 }]);
  });

  it("rejects when useRegex is true and the query does not compile", async () => {
    await expect(
      computeMatches({
        occupiedCells: cells([[0, 0, "x"]]),
        getDisplayText: asText,
        query: "(unterminated",
        useRegex: true
      })
    ).rejects.toThrow();
  });

  it("returns no matches for an empty query", async () => {
    const matches = await computeMatches({
      occupiedCells: cells([[0, 0, "anything"]]),
      getDisplayText: asText,
      query: ""
    });

    expect(matches).toEqual([]);
  });

  it("preserves the iteration order of occupiedCells", async () => {
    const occupiedCells = cells([
      [5, 2, "match"],
      [1, 0, "match"],
      [3, 9, "match"]
    ]);

    const matches = await computeMatches({
      occupiedCells,
      getDisplayText: asText,
      query: "match"
    });

    expect(matches).toEqual([
      { row: 5, column: 2 },
      { row: 1, column: 0 },
      { row: 3, column: 9 }
    ]);
  });

  it("returns null when the signal is already aborted", async () => {
    const controller = new AbortController();
    controller.abort();

    const matches = await computeMatches({
      occupiedCells: cells([[0, 0, "match"]]),
      getDisplayText: asText,
      query: "match",
      signal: controller.signal
    });

    expect(matches).toBeNull();
  });

  it("returns null and stops scanning once the signal aborts mid-scan", async () => {
    const controller = new AbortController();
    const occupiedCells = cells([
      [0, 0, "match"],
      [1, 0, "match"],
      [2, 0, "match"],
      [3, 0, "match"]
    ]);
    const scanned: number[] = [];

    const matches = await computeMatches({
      occupiedCells,
      getDisplayText: (row, column, value) => {
        scanned.push(row);
        if (row === 1) {
          controller.abort();
        }
        return String(value);
      },
      query: "match",
      signal: controller.signal,
      batchSize: 1
    });

    expect(matches).toBeNull();
    expect(scanned).toEqual([0, 1]);
  });

  it("reports progress as cells are scanned", async () => {
    const onProgress = vi.fn();
    const occupiedCells = cells([
      [0, 0, "a"],
      [1, 0, "b"],
      [2, 0, "c"]
    ]);

    await computeMatches({
      occupiedCells,
      getDisplayText: asText,
      query: "a",
      onProgress,
      batchSize: 1
    });

    expect(onProgress).toHaveBeenCalledWith(1);
    expect(onProgress).toHaveBeenCalledWith(2);
    expect(onProgress).toHaveBeenCalledWith(3);
  });
});
