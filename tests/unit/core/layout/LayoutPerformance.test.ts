import { describe, expect, it } from "vitest";

import { LayoutEngine, SkeletonSnapshot } from "../../../../src/core/layout";

function median(samples: readonly number[]): number {
  const sorted = [...samples].sort((left, right) => left - right);
  return sorted[Math.floor(sorted.length / 2)] as number;
}

function measure(operation: () => void, iterations: number): number {
  for (let index = 0; index < 10; index += 1) {
    operation();
  }
  const samples: number[] = [];
  for (let sample = 0; sample < 5; sample += 1) {
    const start = performance.now();
    for (let index = 0; index < iterations; index += 1) {
      operation();
    }
    samples.push((performance.now() - start) / iterations);
  }
  return median(samples);
}

function createLayout(rowCount: number): LayoutEngine {
  return new LayoutEngine({
    rowCount,
    columnCount: 20,
    rowHeight: 24,
    columnWidths: Array.from({ length: 20 }, () => 100)
  });
}

describe("persistent layout scaling guards", () => {
  it("keeps point updates, column roots, and skeleton publication independent of row count", () => {
    const small = createLayout(10_000);
    const large = createLayout(1_000_000);
    let sequence = 0;

    const scenarios = [
      {
        name: "set row height",
        iterations: 100,
        small: () => {
          small.prepareSetRowHeight(sequence++ % 10_000, 25 + (sequence % 5));
        },
        large: () => {
          large.prepareSetRowHeight(sequence++ % 1_000_000, 25 + (sequence % 5));
        }
      },
      {
        name: "set column width",
        iterations: 500,
        small: () => { small.prepareSetColumnWidth(5, 120); },
        large: () => { large.prepareSetColumnWidth(5, 120); }
      },
      {
        name: "publish skeleton",
        iterations: 500,
        small: () => { SkeletonSnapshot.fromRoot(small.root, 1); },
        large: () => { SkeletonSnapshot.fromRoot(large.root, 1); }
      }
    ];

    for (const scenario of scenarios) {
      const smallMs = measure(scenario.small, scenario.iterations);
      const largeMs = measure(scenario.large, scenario.iterations);
      expect(
        largeMs,
        `${scenario.name}: 1m=${largeMs.toFixed(4)}ms, 10k=${smallMs.toFixed(4)}ms`
      ).toBeLessThan(smallMs * 8 + 0.05);
    }
  });

  it("keeps single-row structure edits sublinear at one million rows", () => {
    const small = createLayout(10_000);
    const large = createLayout(1_000_000);
    const smallMs = measure(() => {
      small.prepareInsertRows(5_000, 1);
      small.prepareDeleteRows(5_000, 1);
      small.prepareMoveRows(5_000, 1, 5_010);
    }, 100);
    const largeMs = measure(() => {
      large.prepareInsertRows(500_000, 1);
      large.prepareDeleteRows(500_000, 1);
      large.prepareMoveRows(500_000, 1, 500_010);
    }, 100);

    expect(
      largeMs,
      `structure edits: 1m=${largeMs.toFixed(4)}ms, 10k=${smallMs.toFixed(4)}ms`
    ).toBeLessThan(smallMs * 12 + 0.1);
  });

  it("publishes skeletons in O(1) while retaining a million-column axis identity", () => {
    const createColumnLayout = (columnCount: number) => new LayoutEngine({
      rowCount: 1,
      columnCount,
      rowHeight: 24,
      columnWidths: new Array<number>(columnCount).fill(1)
    });
    const small = createColumnLayout(100);
    const large = createColumnLayout(1_000_000);
    let smallSkeleton = SkeletonSnapshot.fromRoot(small.root, 1);
    let largeSkeleton = SkeletonSnapshot.fromRoot(large.root, 1);

    const smallMs = measure(() => {
      smallSkeleton = SkeletonSnapshot.fromRoot(small.root, 1);
    }, 1_000);
    const largeMs = measure(() => {
      largeSkeleton = SkeletonSnapshot.fromRoot(large.root, 1);
    }, 1_000);

    expect(smallSkeleton.columnAxis).toBe(small.root.columnAxis);
    expect(largeSkeleton.columnAxis).toBe(large.root.columnAxis);
    expect(
      largeMs,
      `skeleton publication: 1m=${largeMs.toFixed(4)}ms, 100=${smallMs.toFixed(4)}ms`
    ).toBeLessThan(smallMs * 8 + 0.05);
  });
});
