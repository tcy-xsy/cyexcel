import { describe, expect, it } from "vitest";

import { RowMetrics } from "../../../../src/core/layout";

function expectMatches(metrics: RowMetrics, oracle: readonly number[]): void {
  expect(metrics.rowCount).toBe(oracle.length);
  expect(metrics.exportDenseValues()).toEqual(new Float64Array(oracle));
  expect(metrics.totalHeight).toBe(oracle.reduce((sum, value) => sum + value, 0));
  let offset = 0;
  for (let row = 0; row < oracle.length; row += 1) {
    const height = oracle[row] as number;
    expect(metrics.heightAt(row)).toBe(height);
    expect(metrics.offsetAt(row)).toBe(offset);
    expect(metrics.rowAtOffset(offset)).toBe(row);
    expect(metrics.rowAtOffset(offset + height - 0.01)).toBe(row);
    offset += height;
  }
  expect(metrics.offsetAt(oracle.length)).toBe(offset);
  expect(metrics.rowAtOffset(offset)).toBe(oracle.length);
}

describe("RowMetrics", () => {
  it("keeps a million uniform rows compact while path-copying point updates", () => {
    const metrics = RowMetrics.uniform(1_000_000, 24);
    const updated = metrics.withHeight(500_000, 40);

    expect(metrics.heightAt(500_000)).toBe(24);
    expect(updated.heightAt(500_000)).toBe(40);
    expect(updated.offsetAt(500_001)).toBe(500_000 * 24 + 40);
    expect(updated.rowAtOffset(500_000 * 24 + 39)).toBe(500_000);
  });

  it("normalizes non-finite offsets and rejects aggregate overflow", () => {
    const metrics = RowMetrics.fromDenseValues([10, 20]);

    expect(metrics.rowAtOffset(Number.NaN)).toBe(0);
    expect(metrics.rowAtOffset(Number.NEGATIVE_INFINITY)).toBe(0);
    expect(metrics.rowAtOffset(Number.POSITIVE_INFINITY)).toBe(2);
    expect(() => RowMetrics.uniform(2, Number.MAX_VALUE)).toThrow(RangeError);
    expect(() =>
      RowMetrics.uniform(Number.MAX_SAFE_INTEGER, Number.MIN_VALUE).insert(
        Number.MAX_SAFE_INTEGER,
        RowMetrics.uniform(1, Number.MIN_VALUE)
      )
    ).toThrow(RangeError);
  });

  it("keeps decimal prefix boundaries half-open", () => {
    const values = [28.37469, 43.514002, 3.866126, 22.08899, 35.943708];
    const metrics = RowMetrics.fromDenseValues(values);

    for (let row = 0; row <= values.length; row += 1) {
      expect(metrics.rowAtOffset(metrics.offsetAt(row))).toBe(row);
    }
  });

  it("supports persistent insert, delete, move, and fragment reuse", () => {
    const original = RowMetrics.fromDenseValues([10, 20, 30, 40, 50]);
    const fragment = RowMetrics.fromDenseValues([7, 8]);
    const inserted = original.insert(2, fragment);
    const deleted = inserted.delete(1, 3);
    const restored = deleted.metrics.insert(1, deleted.deleted);
    const moved = restored.move(1, 3, 7);

    expectMatches(original, [10, 20, 30, 40, 50]);
    expectMatches(inserted, [10, 20, 7, 8, 30, 40, 50]);
    expectMatches(deleted.metrics, [10, 30, 40, 50]);
    expectMatches(deleted.deleted, [20, 7, 8]);
    expectMatches(restored, [10, 20, 7, 8, 30, 40, 50]);
    expectMatches(moved, [10, 30, 40, 50, 20, 7, 8]);
  });

  it("matches a dense oracle through deterministic mixed edits", () => {
    let metrics = RowMetrics.uniform(300, 20);
    const oracle = Array.from({ length: 300 }, () => 20);
    let seed = 0x5eed;
    const random = (limit: number): number => {
      seed = (seed * 1664525 + 1013904223) >>> 0;
      return limit === 0 ? 0 : seed % limit;
    };

    for (let step = 0; step < 150; step += 1) {
      const operation = random(4);
      if (operation === 0 || oracle.length === 0) {
        const row = random(Math.max(1, oracle.length));
        if (oracle.length === 0) {
          metrics = metrics.insert(0, RowMetrics.uniform(1, 20));
          oracle.push(20);
        } else {
          const height = random(50) + 1;
          metrics = metrics.withHeight(row, height);
          oracle[row] = height;
        }
      } else if (operation === 1) {
        const start = random(oracle.length + 1);
        const values = [random(50) + 1, random(50) + 1];
        metrics = metrics.insert(start, RowMetrics.fromDenseValues(values));
        oracle.splice(start, 0, ...values);
      } else if (operation === 2 && oracle.length > 1) {
        const start = random(oracle.length);
        const count = Math.min(random(4) + 1, oracle.length - start);
        metrics = metrics.delete(start, count).metrics;
        oracle.splice(start, count);
      } else if (oracle.length > 1) {
        const start = random(oracle.length);
        const count = Math.min(random(4) + 1, oracle.length - start);
        let target = random(oracle.length + 1);
        if (target > start && target < start + count) {
          target = start;
        }
        const moved = oracle.splice(start, count);
        const adjustedTarget = target > start ? target - count : target;
        oracle.splice(adjustedTarget, 0, ...moved);
        metrics = metrics.move(start, count, target);
      }
      expect(metrics.exportDenseValues()).toEqual(new Float64Array(oracle));
    }
    expectMatches(metrics, oracle);
  });
});

describe("RowMetrics.reorder", () => {
  it("moves each row's height to its new position", () => {
    const metrics = RowMetrics.create(3, 24, new Map([[0, 60], [2, 40]]));
    const reordered = metrics.reorder(Int32Array.from([2, 0, 1]));

    expect(reordered.heightAt(0)).toBe(40);
    expect(reordered.heightAt(1)).toBe(60);
    expect(reordered.heightAt(2)).toBe(24);
  });

  it("keeps the total height unchanged", () => {
    const metrics = RowMetrics.create(3, 24, new Map([[0, 60], [2, 40]]));

    expect(metrics.reorder(Int32Array.from([2, 0, 1])).totalHeight).toBe(metrics.totalHeight);
  });

  it("recomputes offsets in the new order", () => {
    const metrics = RowMetrics.create(3, 24, new Map([[0, 60], [2, 40]]));
    const reordered = metrics.reorder(Int32Array.from([2, 0, 1]));

    expect(reordered.offsetAt(0)).toBe(0);
    expect(reordered.offsetAt(1)).toBe(40);
    expect(reordered.offsetAt(2)).toBe(100);
  });

  it("returns the same instance when every row is the default height", () => {
    // The million-row common case: a uniform tree is permutation-invariant,
    // so reordering must not materialise a dense array.
    const rowCount = 1_000_000;
    const metrics = RowMetrics.uniform(rowCount, 24);
    const order = new Int32Array(rowCount);
    for (let index = 0; index < rowCount; index += 1) {
      order[index] = rowCount - 1 - index;
    }

    expect(metrics.reorder(order)).toBe(metrics);
  });

  it("rejects an order whose length does not match the row count", () => {
    expect(() => RowMetrics.uniform(3, 24).reorder(Int32Array.from([1, 0]))).toThrow(/length/);
  });
});

describe("RowMetrics zero height", () => {
  it("allows exactly 0 as a height, but still rejects negative/NaN/Infinity", () => {
    expect(RowMetrics.uniform(5, 0).heightAt(0)).toBe(0);
    expect(() => RowMetrics.uniform(5, -1)).toThrow(RangeError);
    expect(() => RowMetrics.uniform(5, Number.NaN)).toThrow(RangeError);
    expect(() => RowMetrics.uniform(5, Number.POSITIVE_INFINITY)).toThrow(RangeError);
  });

  it("withHeight accepts 0", () => {
    const metrics = RowMetrics.uniform(3, 20).withHeight(1, 0);
    expect(metrics.heightAt(0)).toBe(20);
    expect(metrics.heightAt(1)).toBe(0);
    expect(metrics.heightAt(2)).toBe(20);
    expect(metrics.totalHeight).toBe(40);
  });

  it("fromDenseValues accepts 0 among positive heights", () => {
    const metrics = RowMetrics.fromDenseValues([0, 5, 0, 3]);
    expect(metrics.totalHeight).toBe(8);
    expect([...metrics.exportDenseValues()]).toEqual([0, 5, 0, 3]);
  });
});

describe("RowMetrics.withFilteredRows", () => {
  it("zeroes exactly the given rows and leaves the rest untouched", () => {
    const metrics = RowMetrics.fromDenseValues([10, 20, 30, 40, 50]);
    const filtered = metrics.withFilteredRows(new Set([1, 3]));

    expect([...filtered.exportDenseValues()]).toEqual([10, 0, 30, 0, 50]);
    expect(filtered.totalHeight).toBe(90);
    // The original is untouched (purity / no write path to the source tree).
    expect([...metrics.exportDenseValues()]).toEqual([10, 20, 30, 40, 50]);
  });

  it("is a pure no-op (returns this) for an empty set", () => {
    const metrics = RowMetrics.fromDenseValues([10, 20, 30]);
    expect(metrics.withFilteredRows(new Set())).toBe(metrics);
  });

  it("collapses a large contiguous filtered run so rowAtOffset still walks past it in O(log n) reachability", () => {
    const rowCount = 5000;
    const metrics = RowMetrics.uniform(rowCount, 20);
    const filteredRows = new Set<number>();
    for (let row = 100; row < 4900; row += 1) {
      filteredRows.add(row);
    }
    const filtered = metrics.withFilteredRows(filteredRows);

    // Rows before/after the filtered run keep their real height and offset.
    expect(filtered.heightAt(50)).toBe(20);
    expect(filtered.heightAt(4950)).toBe(20);
    // The whole filtered run contributes 0 to total height and offsets.
    expect(filtered.offsetAt(100)).toBe(100 * 20);
    expect(filtered.offsetAt(4900)).toBe(100 * 20);
    expect(filtered.totalHeight).toBe((rowCount - 4800) * 20);
    // rowAtOffset lands exactly on the first visible row after the run.
    expect(filtered.rowAtOffset(100 * 20)).toBe(4900);
  });

  it("throws for an out-of-range row", () => {
    const metrics = RowMetrics.uniform(3, 20);
    expect(() => metrics.withFilteredRows(new Set([5]))).toThrow(RangeError);
  });
});


describe("RowMetrics non-zero traversal", () => {
  it("skips compressed runs of a billion zero-height rows and respects bounds", () => {
    const metrics = RowMetrics.uniform(1_000_000_000, 0)
      .withHeights(new Map([[0, 20], [512, 30], [999_999_999, 40]]));
    expect([...metrics.nonZeroRows()]).toEqual([0, 512, 999_999_999]);
    expect([...metrics.nonZeroRows(1, 512)]).toEqual([512]);
    expect([...metrics.nonZeroRows(513, 999_999_486)]).toEqual([]);
    expect([...metrics.nonZeroRows(999_999_999, 0)]).toEqual([]);
  });
});
