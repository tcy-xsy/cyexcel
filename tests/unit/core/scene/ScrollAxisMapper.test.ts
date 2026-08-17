import { describe, expect, it } from "vitest";

import {
  DEFAULT_SAFE_NATIVE_EXTENT,
  ScrollAxisMapper
} from "../../../../src/core/scene";

/** A mapper whose safe window is small enough to force segmenting in a test. */
function segmentedMapper(logicalExtent = 10_000, viewportSize = 100): ScrollAxisMapper {
  const mapper = new ScrollAxisMapper({ safeNativeExtent: 1_000 });
  mapper.setMetrics(logicalExtent, viewportSize);
  return mapper;
}

describe("ScrollAxisMapper identity behaviour", () => {
  it("passes offsets through unchanged while the content fits the safe window", () => {
    const mapper = new ScrollAxisMapper({ safeNativeExtent: 1_000 });
    mapper.setMetrics(800, 100);

    expect(mapper.segmented).toBe(false);
    expect(mapper.nativeExtent).toBe(800);
    expect(mapper.segmentOrigin).toBe(0);
    expect(mapper.toLogical(0)).toBe(0);
    expect(mapper.toLogical(250)).toBe(250);
    expect(mapper.toNative(250)).toEqual({ segmentOrigin: 0, nativeOffset: 250 });
  });

  it("never recentres an unsegmented axis", () => {
    const mapper = new ScrollAxisMapper({ safeNativeExtent: 1_000 });
    mapper.setMetrics(800, 100);
    expect(mapper.recenter(0)).toBe(null);
    expect(mapper.recenter(700)).toBe(null);
  });

  it("defaults to a safe window below the smallest browser limit", () => {
    const mapper = new ScrollAxisMapper();
    mapper.setMetrics(DEFAULT_SAFE_NATIVE_EXTENT * 4, 500);
    expect(mapper.nativeExtent).toBe(DEFAULT_SAFE_NATIVE_EXTENT);
    expect(mapper.segmented).toBe(true);
  });

  it("clamps the spacer to the safe window once the content exceeds it", () => {
    const mapper = segmentedMapper();
    expect(mapper.nativeExtent).toBe(1_000);
    expect(mapper.maxLogicalScroll).toBe(9_900);
    expect(mapper.maxNativeScroll).toBe(900);
    expect(mapper.maxSegmentOrigin).toBe(9_000);
  });
});

describe("ScrollAxisMapper the logical = origin + native equation", () => {
  it("holds for every origin the mapper can choose", () => {
    const mapper = segmentedMapper();
    for (const logical of [0, 1, 450, 899, 900, 5_000, 9_899, 9_900]) {
      const position = mapper.toNative(logical);
      mapper.commit(position);
      expect(position.segmentOrigin + position.nativeOffset).toBeCloseTo(logical, 10);
      expect(mapper.toLogical(position.nativeOffset)).toBeCloseTo(logical, 10);
    }
  });

  it("keeps the native offset inside the scroller's own range", () => {
    const mapper = segmentedMapper();
    for (const logical of [0, 100, 4_000, 9_900]) {
      const position = mapper.toNative(logical);
      expect(position.nativeOffset).toBeGreaterThanOrEqual(0);
      expect(position.nativeOffset).toBeLessThanOrEqual(mapper.maxNativeScroll);
      expect(position.segmentOrigin).toBeGreaterThanOrEqual(0);
      expect(position.segmentOrigin).toBeLessThanOrEqual(mapper.maxSegmentOrigin);
    }
  });

  it("reaches the very last logical pixel — the whole point of segmenting", () => {
    const mapper = segmentedMapper();
    const position = mapper.toNative(mapper.maxLogicalScroll);
    mapper.commit(position);
    expect(mapper.toLogical(position.nativeOffset)).toBe(9_900);
  });

  it("clamps a logical offset past the end instead of overshooting", () => {
    const mapper = segmentedMapper();
    const position = mapper.toNative(99_999);
    expect(position.segmentOrigin + position.nativeOffset).toBe(mapper.maxLogicalScroll);
  });

  it("clamps a native offset past the end", () => {
    const mapper = segmentedMapper();
    expect(mapper.toLogical(999_999)).toBe(mapper.maxLogicalScroll);
  });
});

describe("ScrollAxisMapper recentring", () => {
  it("does nothing while the offset sits in the middle of the window", () => {
    const mapper = segmentedMapper();
    mapper.commit(mapper.toNative(5_000));
    expect(mapper.recenter(mapper.maxNativeScroll / 2)).toBe(null);
  });

  it("shifts the origin without moving the logical position", () => {
    const mapper = segmentedMapper();
    mapper.commit(mapper.toNative(5_000));
    const before = mapper.toLogical(mapper.maxNativeScroll - 1);

    const recentred = mapper.recenter(mapper.maxNativeScroll - 1);
    expect(recentred).not.toBe(null);
    mapper.commit(recentred!);

    expect(recentred!.segmentOrigin + recentred!.nativeOffset).toBeCloseTo(before, 10);
    expect(mapper.toLogical(recentred!.nativeOffset)).toBeCloseTo(before, 10);
  });

  it("does not recentre at the very top, where there is nowhere lower to go", () => {
    const mapper = segmentedMapper();
    mapper.commit(mapper.toNative(0));
    expect(mapper.segmentOrigin).toBe(0);
    expect(mapper.recenter(0)).toBe(null);
  });

  it("does not recentre at the very bottom", () => {
    const mapper = segmentedMapper();
    mapper.commit(mapper.toNative(mapper.maxLogicalScroll));
    expect(mapper.segmentOrigin).toBe(mapper.maxSegmentOrigin);
    expect(mapper.recenter(mapper.maxNativeScroll)).toBe(null);
  });

  it("survives repeated cross-segment travel without drifting", () => {
    const mapper = segmentedMapper();
    let logical = 0;
    mapper.commit(mapper.toNative(logical));
    let native = mapper.toNative(logical).nativeOffset;

    // Walk the whole axis down and back up in 37px steps, recentring on the
    // way, and check the logical position never diverges from the ground truth.
    for (const direction of [1, -1, 1, -1]) {
      for (let step = 0; step < 400; step += 1) {
        const nextNative = native + direction * 37;
        if (nextNative < 0 || nextNative > mapper.maxNativeScroll) {
          break;
        }
        native = nextNative;
        logical = mapper.toLogical(native);
        const recentred = mapper.recenter(native);
        if (recentred) {
          expect(recentred.segmentOrigin + recentred.nativeOffset).toBeCloseTo(logical, 9);
          mapper.commit(recentred);
          native = recentred.nativeOffset;
          expect(mapper.toLogical(native)).toBeCloseTo(logical, 9);
        }
      }
    }
    expect(logical).toBeGreaterThanOrEqual(0);
    expect(logical).toBeLessThanOrEqual(mapper.maxLogicalScroll);
  });

  it("leaves usable slack on both sides, so the next scroll does not immediately recentre", () => {
    const mapper = segmentedMapper();
    const position = mapper.toNative(5_000);
    mapper.commit(position);

    // The whole reason toNative() centres rather than parking the offset at an
    // edge: after landing, the user must be able to scroll a meaningful
    // distance either way before another recentre is needed.
    expect(position.nativeOffset).toBeCloseTo(mapper.maxNativeScroll / 2, 6);
    expect(mapper.recenter(position.nativeOffset)).toBe(null);

    const guard = mapper.maxNativeScroll / 4;
    expect(mapper.recenter(position.nativeOffset - guard / 2)).toBe(null);
    expect(mapper.recenter(position.nativeOffset + guard / 2)).toBe(null);
  });

  it("a recentre lands far enough from both edges to be stable", () => {
    const mapper = segmentedMapper();
    mapper.commit(mapper.toNative(5_000));
    const recentred = mapper.recenter(mapper.maxNativeScroll - 1);
    mapper.commit(recentred!);

    const guard = mapper.maxNativeScroll / 4;
    expect(recentred!.nativeOffset).toBeGreaterThan(guard);
    expect(recentred!.nativeOffset).toBeLessThan(mapper.maxNativeScroll - guard);
  });

  it("recentring is idempotent — the result never needs recentring again", () => {
    const mapper = segmentedMapper();
    mapper.commit(mapper.toNative(5_000));
    const recentred = mapper.recenter(mapper.maxNativeScroll - 1);
    mapper.commit(recentred!);
    expect(mapper.recenter(recentred!.nativeOffset)).toBe(null);
  });
});

describe("ScrollAxisMapper metric changes", () => {
  it("re-clamps the origin when the content shrinks under it", () => {
    const mapper = segmentedMapper();
    mapper.commit(mapper.toNative(mapper.maxLogicalScroll));
    expect(mapper.segmentOrigin).toBe(9_000);

    mapper.setMetrics(2_000, 100);

    expect(mapper.segmentOrigin).toBeLessThanOrEqual(mapper.maxSegmentOrigin);
    expect(mapper.segmentOrigin).toBe(1_000);
  });

  it("drops back to identity when the content shrinks inside the safe window", () => {
    const mapper = segmentedMapper();
    mapper.commit(mapper.toNative(5_000));
    expect(mapper.segmented).toBe(true);

    mapper.setMetrics(600, 100);

    expect(mapper.segmented).toBe(false);
    expect(mapper.segmentOrigin).toBe(0);
    expect(mapper.toLogical(250)).toBe(250);
  });

  it("reports whether anything actually changed", () => {
    const mapper = segmentedMapper();
    expect(mapper.setMetrics(10_000, 100)).toBe(false);
    expect(mapper.setMetrics(10_000, 120)).toBe(true);
  });

  it("bumps the version on a metric change and on an origin move", () => {
    const mapper = segmentedMapper();
    const start = mapper.version;
    mapper.setMetrics(20_000, 100);
    expect(mapper.version).toBe(start + 1);
    mapper.commit(mapper.toNative(9_000));
    expect(mapper.version).toBe(start + 2);
  });

  it("does not bump the version when committing the origin it already has", () => {
    const mapper = segmentedMapper();
    mapper.commit(mapper.toNative(5_000));
    const version = mapper.version;
    mapper.commit({ segmentOrigin: mapper.segmentOrigin, nativeOffset: 0 });
    expect(mapper.version).toBe(version);
  });
});

describe("ScrollAxisMapper degenerate inputs", () => {
  it("handles a viewport larger than the content", () => {
    const mapper = new ScrollAxisMapper({ safeNativeExtent: 1_000 });
    mapper.setMetrics(50, 200);
    expect(mapper.maxLogicalScroll).toBe(0);
    expect(mapper.segmented).toBe(false);
    expect(mapper.toLogical(0)).toBe(0);
    expect(mapper.toNative(999)).toEqual({ segmentOrigin: 0, nativeOffset: 0 });
  });

  it("handles zero-sized content", () => {
    const mapper = new ScrollAxisMapper({ safeNativeExtent: 1_000 });
    mapper.setMetrics(0, 0);
    expect(mapper.nativeExtent).toBe(0);
    expect(mapper.toLogical(0)).toBe(0);
  });

  it("rejects negative and non-finite metrics", () => {
    const mapper = new ScrollAxisMapper();
    expect(() => mapper.setMetrics(-1, 100)).toThrow(RangeError);
    expect(() => mapper.setMetrics(Number.NaN, 100)).toThrow(RangeError);
    expect(() => mapper.setMetrics(100, Number.POSITIVE_INFINITY)).toThrow(RangeError);
  });

  it("rejects a non-positive safe extent at construction", () => {
    expect(() => new ScrollAxisMapper({ safeNativeExtent: 0 })).toThrow(RangeError);
    expect(() => new ScrollAxisMapper({ safeNativeExtent: -5 })).toThrow(RangeError);
  });

  it("rejects a negative native offset", () => {
    const mapper = segmentedMapper();
    expect(() => mapper.toLogical(-1)).toThrow(RangeError);
  });
});

describe("ScrollAxisMapper fractional offsets", () => {
  it("keeps sub-pixel logical positions exact across a recentre", () => {
    const mapper = segmentedMapper();
    mapper.commit(mapper.toNative(5_000.25));
    const native = mapper.maxNativeScroll - 0.5;
    const logical = mapper.toLogical(native);

    const recentred = mapper.recenter(native);
    mapper.commit(recentred!);

    expect(mapper.toLogical(recentred!.nativeOffset)).toBeCloseTo(logical, 10);
  });

  it("does not accumulate error when the native offset is repeatedly rounded", () => {
    const mapper = segmentedMapper();
    mapper.commit(mapper.toNative(4_000));
    // Simulate a browser that reports integral scroll positions: the logical
    // value is re-derived each time rather than accumulated, so the only error
    // is the current rounding, never a growing sum of them.
    let native = Math.round(mapper.toNative(4_000).nativeOffset);
    for (let step = 0; step < 200; step += 1) {
      native = Math.round(Math.min(native + 13.7, mapper.maxNativeScroll));
      const logical = mapper.toLogical(native);
      expect(logical).toBe(mapper.segmentOrigin + native);
      const recentred = mapper.recenter(native);
      if (recentred) {
        mapper.commit(recentred);
        native = Math.round(recentred.nativeOffset);
      }
    }
  });
});
