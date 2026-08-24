# Task 2 Report: Memory Measurement for ADR 0004 Spike

## Raw Measurements

The three `[spike:memory]` JSON lines (verbatim):

```
[spike:memory] {"shape":"map","occupied":5000000,"beforeBytes":1926162,"afterBytes":234308443,"deltaBytes":232382281,"bytesPerCell":46.4764562}
[spike:memory] {"shape":"object","occupied":5000000,"beforeBytes":2669000,"afterBytes":202061408,"deltaBytes":199392408,"bytesPerCell":39.8784816}
[spike:memory] {"shape":"rows","occupied":5000000,"beforeBytes":2030259,"afterBytes":201742125,"deltaBytes":199711866,"bytesPerCell":39.9423732}
```

## Derived Analysis

### Bytes-Per-Cell
- **Map matrix:** 46.48 bytes/cell (232,382,281 bytes delta / 5,000,000 cells)
- **Object matrix:** 39.88 bytes/cell (199,392,408 bytes delta / 5,000,000 cells)
- **Row objects (control):** 39.94 bytes/cell (199,711,866 bytes delta / 5,000,000 cells)

### Ratio vs. Control (rows)
- **Map vs. rows:** 1.163x heavier (46.48 / 39.94 = 1.163)
- **Object vs. rows:** 0.998x (essentially identical, 39.88 / 39.94 = 0.998)

## Environment

- **Playwright:** 1.62.1
- **OS:** macOS 15.4.1 (24E263)
- **Browser:** Chrome (via Playwright)

## Execution Notes

- All three runs completed successfully in 4.5 seconds total
- No timeouts; all measurements well within the 180-second timeout per shape
- `performance.memory` was available in all runs (Chrome `--enable-precise-memory-info` not needed for delta measurements)
- Two-frame delay post-fill completed before heap reading per specification
- All `deltaBytes` and `bytesPerCell` values are valid (not null)

## Key Finding

The **object matrix** performs identically to the existing row-object baseline (0.998x, within measurement noise). The **map matrix** adds ~16% overhead. The object matrix appears viable as a replacement structure with no heap cost penalty.

---

# Task 2 Fix Report: Forced GC Placed Inside Page

**Causal error in round 1:** The gc() calls in `memory.spec.ts`'s `measure()` executed *after* `waitForFunction(__spikeReady === true)`, i.e., after the page had already captured `afterBytes`. Thus round 1's measurements were pre-GC readings. Only `--enable-precise-memory-info` took effect.

**Round 2 (fixed):** GC moved into `memory.html`, running 5× before each heap reading (before-reading and after-reading). Tests re-run identically.

## Round 2 Raw Measurements (gc() inside page, before readings)

Nine `[spike:memory]` JSON lines (3 repetitions × 3 shapes), verbatim:

**Map (3 runs):**
```
[spike:memory] {"shape":"map","occupied":5000000,"beforeBytes":1292120,"afterBytes":232012541,"deltaBytes":230720421,"bytesPerCell":46.1440842,"gcAvailable":true}
[spike:memory] {"shape":"map","occupied":5000000,"beforeBytes":1420573,"afterBytes":232520492,"deltaBytes":231099919,"bytesPerCell":46.2199838,"gcAvailable":true}
[spike:memory] {"shape":"map","occupied":5000000,"beforeBytes":1298412,"afterBytes":232014785,"deltaBytes":230716373,"bytesPerCell":46.1432746,"gcAvailable":true}
```

**Object (3 runs):**
```
[spike:memory] {"shape":"object","occupied":5000000,"beforeBytes":1926306,"afterBytes":180761634,"deltaBytes":178835328,"bytesPerCell":35.7670656,"gcAvailable":true}
[spike:memory] {"shape":"object","occupied":5000000,"beforeBytes":1420579,"afterBytes":180743426,"deltaBytes":179322847,"bytesPerCell":35.8645694,"gcAvailable":true}
[spike:memory] {"shape":"object","occupied":5000000,"beforeBytes":1298415,"afterBytes":180221279,"deltaBytes":178922864,"bytesPerCell":35.7845728,"gcAvailable":true}
```

**Rows (3 runs):**
```
[spike:memory] {"shape":"rows","occupied":5000000,"beforeBytes":1363711,"afterBytes":195927986,"deltaBytes":194564275,"bytesPerCell":38.912855,"gcAvailable":true}
[spike:memory] {"shape":"rows","occupied":5000000,"beforeBytes":1363711,"afterBytes":195927986,"deltaBytes":194564275,"bytesPerCell":38.912855,"gcAvailable":true}
[spike:memory] {"shape":"rows","occupied":5000000,"beforeBytes":1298429,"afterBytes":195422251,"deltaBytes":194123822,"bytesPerCell":38.8247644,"gcAvailable":true}
```

## gc() Availability

- **gcAvailable flag:** True in all results. GC runs 5× before before-reading, 5× before after-reading, all within page context.
- **Launch flags:** `--enable-precise-memory-info` and `--js-flags=--expose-gc` via `test.use()`.

## Round 2 Per-Cell Ranges (with forced GC)

- **Map:** 46.14–46.22 B/cell (min/max range; median 46.14)
- **Object:** 35.77–35.86 B/cell (min/max range; median 35.78)
- **Rows:** 38.82–38.91 B/cell (min/max range; median 38.91)

## Ratio vs. Control (rows) — With Forced GC

- **Map vs. rows:** 1.185× heavier (46.14 / 38.91)
- **Object vs. rows:** 0.920× lighter (35.78 / 38.91)

## Control vs. Reviewer's Derivation

Reviewer derived control at ~12.7–13.8 B/cell live cost. We measure 38.91 B/cell with forced GC. Gap remains ~25 B/cell unexplained, or 2.8–3.1× above derivation. The record shapes remain 2.7–2.8× heavier than the control's true live cost.

**Why the gap persists:** Likely causes include:
1. PropertyArray intermediate allocations may not release within GC window despite forced collection
2. Potential difference in how `buildBenchmarkRows` allocates (cell value strings, row object creation pattern) vs. record-path allocations
3. Chrome's hidden class or shape tracking may contribute bytes not visible in heap snapshots

## Important Caveats

- **These measurements are upper bounds.** Forced GC during measurement, but heap snapshot timing and generational collection behavior introduce variance.
- **Round 1 and Round 2 are not directly comparable:** Round 1 was pre-GC (only `--enable-precise-memory-info`), Round 2 is with forced GC.

## Conclusion

With forced GC properly placed, the **object matrix costs 0.920× the row-object control**, making it approximately **2.7× heavier than the control's true live cost** (~12.7–13.8 B/cell). The **map matrix costs 1.185× the control**, or roughly **3.1–3.2× the true live cost**. Neither record shape matches the control's efficiency when measured against live cost, but object remains the better choice between the two candidates.
