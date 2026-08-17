import { describe, expect, it } from "vitest";

import { decodeTsv, encodeTsv, escapeTsvCell } from "../../../../src/core/clipboard";

describe("escapeTsvCell", () => {
  it("leaves plain text untouched", () => {
    expect(escapeTsvCell("hello")).toBe("hello");
    expect(escapeTsvCell("")).toBe("");
  });

  it("wraps a cell containing a tab in quotes", () => {
    expect(escapeTsvCell("a\tb")).toBe('"a\tb"');
  });

  it("wraps a cell containing a newline in quotes", () => {
    expect(escapeTsvCell("a\nb")).toBe('"a\nb"');
    expect(escapeTsvCell("a\r\nb")).toBe('"a\r\nb"');
  });

  it("wraps a cell containing a quote and doubles the internal quote", () => {
    expect(escapeTsvCell('say "hi"')).toBe('"say ""hi"""');
  });
});

describe("encodeTsv", () => {
  it("joins cells with tabs and rows with CRLF", () => {
    const tsv = encodeTsv([
      ["a", "b"],
      ["c", "d"]
    ]);
    expect(tsv).toBe("a\tb\r\nc\td");
  });

  it("escapes cells that need it inline with the rest of the row", () => {
    const tsv = encodeTsv([["plain", "has\ttab", 'has"quote']]);
    expect(tsv).toBe('plain\t"has\ttab"\t"has""quote"');
  });
});

describe("decodeTsv", () => {
  it("round-trips a plain grid through encode/decode", () => {
    const grid = [
      ["a", "b", "c"],
      ["1", "2", "3"]
    ];
    expect(decodeTsv(encodeTsv(grid))).toEqual(grid);
  });

  it("round-trips cells containing tabs, newlines, and quotes", () => {
    const grid = [["has\ttab", "has\nnewline", 'has"quote"here'], ["plain", "", "x"]];
    expect(decodeTsv(encodeTsv(grid))).toEqual(grid);
  });

  it("splits on bare LF as well as CRLF (real clipboard input isn't guaranteed CRLF)", () => {
    expect(decodeTsv("a\tb\nc\td")).toEqual([
      ["a", "b"],
      ["c", "d"]
    ]);
  });

  it("treats an empty string as a single empty cell", () => {
    expect(decodeTsv("")).toEqual([[""]]);
  });

  it("parses a single cell with no separators", () => {
    expect(decodeTsv("solo")).toEqual([["solo"]]);
  });
});
