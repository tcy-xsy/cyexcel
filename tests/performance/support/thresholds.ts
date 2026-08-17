import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { expect } from "@playwright/test";

/**
 * PR runs use a relative regression check against the checked-in baseline
 * (doc/implementation-steps.md Step 15 rule #3: "普通 PR 使用相对回归阈值").
 * Release runs additionally hold the two scenarios that have an explicit
 * PRD number (plan/PRD.md:48-57) to that absolute target, and run the full
 * scroll/DOM-renderer duration instead of the short PR-mode window.
 */
export const PERF_MODE: "pr" | "release" = process.env.CYEXCEL_PERF_MODE === "release" ? "release" : "pr";

/** 15s in PR mode (fast enough to run on every PR), the real 5 minutes only for a release run. */
export const SCROLL_DURATION_MS = PERF_MODE === "release" ? 5 * 60 * 1000 : 15 * 1000;

/**
 * Generous on purpose: this baseline was recorded in a sandboxed/headless
 * session environment, not a dedicated reference desktop (see this Step's
 * doc/progress.md entry) — a tight tolerance would make every re-run flaky
 * on hardware-driven noise alone. Tightening this is future work once a
 * baseline exists from the real reference environment.
 */
const RELATIVE_TOLERANCE = 0.35;

interface AbsoluteTarget {
  readonly max?: number;
  readonly min?: number;
}

/** plan/PRD.md §3.2 — the only two rows in that table with a literal number; everything else there is qualitative (see this module's callers for how those are checked instead). */
const ABSOLUTE_TARGETS: Readonly<Record<string, Readonly<Record<string, AbsoluteTarget>>>> = {
  firstPaint: { readyMs: { max: 100 } },
  scroll: { fps: { min: 55 } }
};

interface BaselineFile {
  readonly recordedAt: string | null;
  readonly environment: {
    readonly browser: string;
    readonly os: string;
    readonly buildMode: string;
  };
  readonly scenarios: Readonly<Record<string, Readonly<Record<string, number | null>>>>;
}

function loadBaseline(): BaselineFile {
  const path = fileURLToPath(new URL("./baseline.json", import.meta.url));
  return JSON.parse(readFileSync(path, "utf-8")) as BaselineFile;
}

const baseline = loadBaseline();

export interface AssertPerfMetricOptions {
  /**
   * A minimum absolute tolerance floor (same unit as the metric — ms for
   * timings, bytes for heap growth), in addition to the relative
   * percentage. Needed for metrics whose baseline is itself only a few ms
   * (e.g. `planMs`/`commandMs`) — a 35% relative window around 1.5ms is
   * ~0.5ms, well inside normal wall-clock/scheduler/parallel-worker-
   * contention noise — or whose baseline is exactly 0 (e.g.
   * `heapGrowthBytes` over a real 5-minute scroll: `performance.memory` is
   * a coarse, GC-timing-dependent sample, so a real run reads a small
   * non-zero number even with `domNodeGrowth` staying flat at 0; see this
   * Step's doc/progress.md entry). Leave unset (0) for metrics where the
   * relative window alone already dominates (first paint's multi-second
   * construct time) or where any deviation from a 0 baseline is meaningful
   * on its own (long-task count, DOM node growth).
   */
  readonly toleranceFloor?: number;
}

/**
 * Asserts `value` for `scenario.metric` is acceptable for the current
 * `CYEXCEL_PERF_MODE`. Always logs the raw measurement (prefixed
 * `PERF_METRIC`) so a first/record run's numbers can be read back out of
 * the test output and transcribed into baseline.json — no separate
 * baseline-recording script.
 */
export function assertPerfMetric(
  scenario: string,
  metric: string,
  value: number,
  options: AssertPerfMetricOptions = {}
): void {
  console.log(`PERF_METRIC ${JSON.stringify({ scenario, metric, value, mode: PERF_MODE })}`);

  expect(Number.isFinite(value), `${scenario}.${metric} must be a finite number`).toBe(true);
  expect(value, `${scenario}.${metric} must be >= 0`).toBeGreaterThanOrEqual(0);

  const absoluteTarget = ABSOLUTE_TARGETS[scenario]?.[metric];
  if (PERF_MODE === "release" && absoluteTarget) {
    if (absoluteTarget.max !== undefined) {
      expect(value, `${scenario}.${metric} release target: <= ${absoluteTarget.max}`).toBeLessThanOrEqual(
        absoluteTarget.max
      );
    }
    if (absoluteTarget.min !== undefined) {
      expect(value, `${scenario}.${metric} release target: >= ${absoluteTarget.min}`).toBeGreaterThanOrEqual(
        absoluteTarget.min
      );
    }
    return;
  }

  const baselineValue = baseline.scenarios[scenario]?.[metric];
  if (baselineValue === undefined || baselineValue === null) {
    // No recorded baseline yet for this metric — this run establishes it,
    // so only the sanity checks above apply.
    return;
  }

  // "fps"/"min"-style metrics regress downward; everything else (timings,
  // byte counts, node counts) regresses upward. ABSOLUTE_TARGETS already
  // encodes which direction each named metric cares about — reuse it here
  // so a metric like "fps" isn't flagged as a regression for going *up*.
  const isLowerBetter = absoluteTarget?.min === undefined;
  const tolerance = Math.max(baselineValue * RELATIVE_TOLERANCE, options.toleranceFloor ?? 0);
  if (isLowerBetter) {
    expect(
      value,
      `${scenario}.${metric} regressed vs baseline ${baselineValue} by more than ${RELATIVE_TOLERANCE * 100}%`
    ).toBeLessThanOrEqual(baselineValue + tolerance);
  } else {
    expect(
      value,
      `${scenario}.${metric} regressed vs baseline ${baselineValue} by more than ${RELATIVE_TOLERANCE * 100}%`
    ).toBeGreaterThanOrEqual(baselineValue - tolerance);
  }
}
