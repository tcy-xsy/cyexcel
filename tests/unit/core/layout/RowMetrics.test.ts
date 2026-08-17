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
