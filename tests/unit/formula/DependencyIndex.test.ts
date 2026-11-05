import { expect, it } from "vitest";
import { DependencyIndex } from "../../../src/formula/DependencyIndex";

it("matches overlapping rectangles at boundaries and replaces or removes stale dependencies", () => {
  const index = new DependencyIndex();
  const ranges = Array.from({ length: 200 }, (_, i) => ({
    key: String(i), startRow: (i * 17) % 90, endRow: (i * 17) % 90 + i % 13,
    startColumn: (i * 11) % 40, endColumn: (i * 11) % 40 + i % 7
  }));
  for (const range of ranges) index.set(range.key, [range]);
  for (let row = 0; row < 105; row += 3) {
    for (let column = 0; column < 48; column += 2) {
      expect(index.query(row, column)).toEqual(new Set(ranges.filter(range =>
        range.startRow <= row && row <= range.endRow && range.startColumn <= column && column <= range.endColumn).map(r => r.key)));
    }
  }
  index.set("0", [{ key: "0", startRow: 1000, endRow: 100000, startColumn: 0, endColumn: 0 }]);
  expect(index.query(0, 0).has("0")).toBe(false);
  expect(index.query(100000, 0)).toEqual(new Set(["0"]));
  index.delete("0");
  expect(index.query(100000, 0)).toEqual(new Set());
  index.clear();
  expect(index.query(30, 20)).toEqual(new Set());
});
