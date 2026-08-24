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

Reviewer derived control at ~12.7–13.8 B/cell live cost. We measure 38.91 B/cell with forced GC. Gap remains ~25 B/cell, or 2.8–3.1× above derivation. **Round 3 below identifies the actual cause: the control was measured in V8 dictionary mode, not the fast mode real backend JSON produces.** (An earlier draft of this section speculated about PropertyArray allocation and hidden-class tracking as causes; that speculation was superseded by the dictionary-mode finding and has been removed — see Round 3.)

## Important Caveats

- **These measurements are upper bounds.** Forced GC during measurement, but heap snapshot timing and generational collection behavior introduce variance.
- **Round 1 and Round 2 are not directly comparable:** Round 1 was pre-GC (only `--enable-precise-memory-info`), Round 2 is with forced GC.

## Conclusion

With forced GC properly placed, the **object matrix costs 0.920× the row-object control**, making it approximately **2.7× heavier than the control's true live cost** (~12.7–13.8 B/cell). The **map matrix costs 1.185× the control**, or roughly **3.1–3.2× the true live cost**. Neither record shape matches the control's efficiency when measured against live cost, but object remains the better choice between the two candidates.

---

> **Superseded by Round 4 below.** Round 3 correctly identified that the control must be fast-mode, but `buildRowsFromJSON` fabricated a unique `` `${row},${col}` `` string for every one of the 50 columns instead of using the shared `buildCellValue` distribution (pooled region/status strings, booleans, doubles, and only 10/50 columns as fresh strings). That inflated the rows-json control ~2.6×. The **map/object/rows numbers in this round are unaffected and remain final** (see "Settled facts" at the top of Round 4); only the **rows-json numbers and every conclusion drawn from them are superseded** — use Round 4's corrected rows-json control instead.

# Task 2 Round 3: Representative Control (Fast-Mode vs. Dictionary-Mode)

**Critical finding:** The loop-assigned row objects from `buildBenchmarkRows` operate in V8 dictionary mode (HasFastProperties: false), not the fast-mode that real backend JSON produces. This inflates the control's measured cost.

## Object Mode Verification

Tested object construction methods with 50 properties using `node --allow-natives-syntax`:

```
Loop-assigned (buildBenchmarkRows pattern):
  obj = {}; for (i=0; i<50; i++) obj[`c${i}`] = value;
  Result: HasFastProperties = false (dictionary mode, NameDictionary backing)

JSON-parsed (real backend JSON response pattern):
  JSON.parse('{"c0":"...", "c1":"...", ..., "c49":"..."}')
  Result: HasFastProperties = true (fast mode, PropertyArray backing)
```

Dictionary-mode objects at 50 properties are significantly heavier than fast-mode, explaining the 60% cost gap measured below.

## Round 3 Raw Measurements (all four shapes)

Twelve `[spike:memory]` JSON lines (3 repetitions × 4 shapes), verbatim:

**Map (3 runs):**
```
[spike:memory] {"shape":"map","occupied":5000000,"beforeBytes":1307669,"afterBytes":232529541,"deltaBytes":231221872,"bytesPerCell":46.2443744,"gcAvailable":true}
[spike:memory] {"shape":"map","occupied":5000000,"beforeBytes":1307669,"afterBytes":232529581,"deltaBytes":231221912,"bytesPerCell":46.2443824,"gcAvailable":true}
[spike:memory] {"shape":"map","occupied":5000000,"beforeBytes":1372966,"afterBytes":232547877,"deltaBytes":231174911,"bytesPerCell":46.2349822,"gcAvailable":true}
```

**Object (3 runs):**
```
[spike:memory] {"shape":"object","occupied":5000000,"beforeBytes":1307672,"afterBytes":180752683,"deltaBytes":179445011,"bytesPerCell":35.8890022,"gcAvailable":true}
[spike:memory] {"shape":"object","occupied":5000000,"beforeBytes":1307672,"afterBytes":180770899,"deltaBytes":179463227,"bytesPerCell":35.8926454,"gcAvailable":true}
[spike:memory] {"shape":"object","occupied":5000000,"beforeBytes":1307672,"afterBytes":180770899,"deltaBytes":179463227,"bytesPerCell":35.8926454,"gcAvailable":true}
```

**Rows (loop-assigned, dictionary mode, 3 runs):**
```
[spike:memory] {"shape":"rows","occupied":5000000,"beforeBytes":1429832,"afterBytes":195955447,"deltaBytes":194525615,"bytesPerCell":38.905123,"gcAvailable":true}
[spike:memory] {"shape":"rows","occupied":5000000,"beforeBytes":1307670,"afterBytes":195955447,"deltaBytes":194647777,"bytesPerCell":38.9295554,"gcAvailable":true}
[spike:memory] {"shape":"rows","occupied":5000000,"beforeBytes":1372984,"afterBytes":195955447,"deltaBytes":194582463,"bytesPerCell":38.9164926,"gcAvailable":true}
```

**Rows-JSON (fast-mode via JSON.parse, 3 runs):**
```
[spike:memory] {"shape":"rows-json","occupied":5000000,"beforeBytes":1307675,"afterBytes":123677305,"deltaBytes":122369630,"bytesPerCell":24.473926,"gcAvailable":true}
[spike:memory] {"shape":"rows-json","occupied":5000000,"beforeBytes":1953773,"afterBytes":123677297,"deltaBytes":121723524,"bytesPerCell":24.3447048,"gcAvailable":true}
[spike:memory] {"shape":"rows-json","occupied":5000000,"beforeBytes":1307675,"afterBytes":123659101,"deltaBytes":122351426,"bytesPerCell":24.4702852,"gcAvailable":true}
```

## Round 3 Per-Cell Ranges (all four shapes)

- **Map:** 46.23–46.24 B/cell (median 46.24)
- **Object:** 35.89–35.89 B/cell (median 35.89)
- **Rows (dictionary mode):** 38.90–38.93 B/cell (median 38.92)
- **Rows-JSON (fast mode):** 24.34–24.47 B/cell (median 24.47) ← **representative control**

## Ratios Against Representative Control (Rows-JSON, fast-mode)

- **Map vs. rows-json:** 1.89× heavier (46.24 / 24.47)
- **Object vs. rows-json:** 1.47× heavier (35.89 / 24.47)
- **Rows (dictionary mode) vs. rows-json:** 1.59× heavier (38.92 / 24.47)

## Control Cost Confirmation

Rows-JSON (fast-mode) at **24.47 B/cell** is now much closer to the reviewer's first-principles derivation. The ~11–12 B/cell gap likely represents the cost of cell value strings (approximately 5M × "row,col" format ≈ 10–11 B/cell), which are part of the measured structure. This confirms the derivation was correct for realistic fast-mode backend data.

## Architecture Impact

- **Current control (loop-assigned, dictionary mode)** is **not representative** of real data. It overstates heap cost by ~60% due to V8 mode penalty.
- **Record matrices are 1.47–1.89× the true live cost**, not the claimed 2.7–3.2×.
- **Object matrix remains the better choice** between the two candidates, at 1.47× vs. 1.89×.
- **Real-world headline:** Against fast-mode JSON objects (representative control), the object matrix costs **1.47×**, the map matrix **1.89×**.

---

# Task 2 Round 4 (FINAL): Corrected Control — Real `buildCellValue` Distribution Through `JSON.parse`

## Settled facts carried forward, unchanged by this round

- Forced in-page `gc()` before both readings: correct, unchanged (Round 2).
- `--enable-precise-memory-info` + `--js-flags=--expose-gc` launch options: correct, unchanged.
- The loop-assigned control (`rows`) is V8 dictionary mode; `JSON.parse`-built objects are fast mode. Confirmed with `%HasFastProperties` (Round 3), reconfirmed here.
- **Record matrices' own numbers are stable across every round and unaffected by this fix: Object ≈ 35.89, Map ≈ 46.23 B/cell.**
- The dictionary-mode `rows` control (≈38.8 B/cell) is also unaffected by this fix — the bug was only in `rows-json`.

## The bug (Round 3's `rows-json`)

`buildRowsFromJSON` in `memory.html` fabricated a unique `` `${row},${col}` `` string for **all 50 columns**, never calling `buildCellValue`. Real rows are only 10/50 fresh unique strings (`` `Row ${row}` ``); the other 40/50 are pooled region/status string references, doubles, and booleans — values that cost little or nothing marginally because the pool arrays are allocated once at module load, before the "before" heap reading. Fabricating 50 unique strings/row instead of 10 inflated the control ~2.6×, which is why Round 3 concluded record matrices cost only 1.47–1.89× the control.

## The fix

`buildRowsFromJSON` now imports `buildBenchmarkRow` from `tests/performance/support/dataset.ts` (unmodified) and round-trips each row through `JSON.parse(JSON.stringify(rawRow))`:

```js
import { buildBenchmarkRows, buildBenchmarkRow } from "/tests/performance/support/dataset.ts";

function buildRowsFromJSON(rowCount, columnCount) {
  const rows = [];
  for (let row = 0; row < rowCount; row++) {
    const rawRow = buildBenchmarkRow(row, columnCount);
    const obj = JSON.parse(JSON.stringify(rawRow));
    rows.push(obj);
  }
  return rows;
}
```

This guarantees the parsed objects carry exactly `buildCellValue`'s distribution (same function, not a re-implementation), while still exercising `JSON.parse` per row to get fast-mode objects.

## Verification (required by brief, done before re-running the browser measurement)

Ran a standalone replica of this exact logic under `node --allow-natives-syntax` (`buildCellValue`/`buildBenchmarkRow` copied verbatim, `JSON.parse(JSON.stringify(row))` per row, 1,000 rows × 50 columns):

```
HasFastProperties(rows[0]): true
HasFastProperties(rows[500]): true
Comparing c2 of row 0 and row 5 - values: APAC APAC
Reference-identical (===): true
Same reference as REGIONS entry: true
Comparing c3 of row 0 and row 3 - values: active active
Reference-identical (===): true
```

- **Fast-mode: confirmed.** Every sampled parsed row has `HasFastProperties === true`.
- **Pooled-value dedupe: confirmed, and stronger than the brief anticipated.** Two different rows' `c2` (REGIONS-derived) values are not just equal but `===` reference-identical after independently going through `JSON.stringify`/`JSON.parse` — and identical to the original `REGIONS` array entry itself. Same result for `c3`/STATUSES. V8's JSON parser interns short strings into the shared string table, so pooled values collapse back to shared references even though each row was serialized and parsed independently. The control is not inflated by lost sharing.

## Round 4 Raw Measurements (all four shapes, 3 repetitions each)

Command: `pnpm exec playwright test tests/spike/adr-0004/memory.spec.ts --project=chrome --workers=1` — 12/12 passed.

```
[spike:memory] {"shape":"map","occupied":5000000,"beforeBytes":1307948,"afterBytes":232024113,"deltaBytes":230716165,"bytesPerCell":46.143233,"gcAvailable":true}
[spike:memory] {"shape":"map","occupied":5000000,"beforeBytes":1373245,"afterBytes":232530028,"deltaBytes":231156783,"bytesPerCell":46.2313566,"gcAvailable":true}
[spike:memory] {"shape":"map","occupied":5000000,"beforeBytes":1307948,"afterBytes":232530028,"deltaBytes":231222080,"bytesPerCell":46.244416,"gcAvailable":true}
[spike:memory] {"shape":"object","occupied":5000000,"beforeBytes":1373251,"afterBytes":180247239,"deltaBytes":178873988,"bytesPerCell":35.7747976,"gcAvailable":true}
[spike:memory] {"shape":"object","occupied":5000000,"beforeBytes":1430115,"afterBytes":180771166,"deltaBytes":179341051,"bytesPerCell":35.8682102,"gcAvailable":true}
[spike:memory] {"shape":"object","occupied":5000000,"beforeBytes":1307951,"afterBytes":180247239,"deltaBytes":178939288,"bytesPerCell":35.7878576,"gcAvailable":true}
[spike:memory] {"shape":"rows","occupied":5000000,"beforeBytes":1307937,"afterBytes":195431787,"deltaBytes":194123850,"bytesPerCell":38.82477,"gcAvailable":true}
[spike:memory] {"shape":"rows","occupied":5000000,"beforeBytes":1301657,"afterBytes":195937502,"deltaBytes":194635845,"bytesPerCell":38.927169,"gcAvailable":true}
[spike:memory] {"shape":"rows","occupied":5000000,"beforeBytes":1430579,"afterBytes":195429799,"deltaBytes":193999220,"bytesPerCell":38.799844,"gcAvailable":true}
[spike:memory] {"shape":"rows-json","occupied":5000000,"beforeBytes":1307954,"afterBytes":37526229,"deltaBytes":36218275,"bytesPerCell":7.243655,"gcAvailable":true}
[spike:memory] {"shape":"rows-json","occupied":5000000,"beforeBytes":1373257,"afterBytes":38033712,"deltaBytes":36660455,"bytesPerCell":7.332091,"gcAvailable":true}
[spike:memory] {"shape":"rows-json","occupied":5000000,"beforeBytes":1430589,"afterBytes":38031964,"deltaBytes":36601375,"bytesPerCell":7.320275,"gcAvailable":true}
```

## Round 4 Per-Cell Ranges (FINAL — use these)

- **Map:** 46.14–46.24 B/cell (median 46.23)
- **Object:** 35.77–35.87 B/cell (median 35.79)
- **Rows (dictionary mode):** 38.80–38.93 B/cell (median 38.82)
- **Rows-JSON (corrected control, fast mode, real distribution):** 7.24–7.33 B/cell (median 7.32)

## Ratios Against the Corrected Control (median-based)

- **Object vs. rows-json: 4.89×** heavier (35.79 / 7.32)
- **Map vs. rows-json: 6.32×** heavier (46.23 / 7.32)
- **Rows (dictionary mode) vs. rows-json: 5.30×** heavier — this is the pure dictionary-vs-fast-mode penalty, unrelated to the record-matrix question.

## This does not match the ~2.9× expectation — reported plainly, not adjusted

The task brief expected the object matrix to land near 2.9× once the control was corrected, based on an independent first-principles derivation of ~12.7–13.8 B/cell for the corrected control, and a reviewer's Node measurement of the *exact same corrected distribution/build method* at **12.49 B/cell** (reproduced independently here — see below). Chrome measures the corrected control at only **7.24–7.33 B/cell, about 1.7× lower than Node's 12.49 for the identical code**.

To rule out a bug rather than an engine difference, the same dictionary-mode `rows` control was also measured in Node with the identical `buildBenchmarkRow`/`buildBenchmarkRows` logic:

```
Node,  rows (dictionary mode):    72.81 B/cell   (HasFastProperties: false)
Chrome, rows (dictionary mode):   38.82 B/cell (median, this round)
→ Node/Chrome ratio: 1.87×

Node,  rows-json (fast mode, corrected distribution): 12.49 B/cell (HasFastProperties: true)
Chrome, rows-json (fast mode, corrected distribution): 7.32 B/cell (median, this round)
→ Node/Chrome ratio: 1.71×
```

Node consistently reports ~1.7–1.9× higher absolute bytes-per-cell than Chrome for the *same* JS logic, for both dictionary-mode and fast-mode objects alike. This is very likely a Node-vs-Chrome V8 build/embedding difference (different heap page granularity, GC tuning, or startup snapshot — not investigated further; out of scope for this spike) rather than a bug in either measurement, because: (a) the Round-4 Node script reproduces the reviewer's reported 12.4917504 ≈ 12.49 B/cell exactly when running the exact same corrected logic, confirming the fix itself is correct; and (b) the ~1.7–1.9× gap is consistent across both dictionary-mode and fast-mode shapes rather than being specific to `rows-json`, which would be the signature of a construction bug rather than an engine artifact. All four shapes in this spike (map, object, rows, rows-json) are measured in the same browser engine via the same methodology, so the **ratios among them, computed from Chrome numbers only, are the correct comparison for this spike's question** — mixing a Chrome-measured numerator with a Node-derived denominator (or a first-principles estimate) is what produced the ~2.9× expectation, and that mixing is the likely source of the mismatch.

## Final Conclusion

Using the corrected, apples-to-apples Chrome-measured control (7.24–7.33 B/cell, real `buildCellValue` distribution, fast-mode via `JSON.parse`, pooled values confirmed reference-deduped):

- **Object matrix costs ≈4.9× the corrected control** (35.79 / 7.32), not 1.47× (Round 3, control bug) and not ~2.9× (brief's expectation, based on a Node/derivation figure not measured in the same engine as the record matrices).
- **Map matrix costs ≈6.3× the corrected control** (46.23 / 7.32).
- Object remains the better of the two record-matrix candidates, as in every prior round.
- **Open question, not resolved by this spike:** why Chrome's absolute B/cell for fast-mode and dictionary-mode objects runs ~1.7–1.9× below Node's for identical code. This affects the *absolute* control number but not the internal consistency of the Round 4 ratios (map/object/rows/rows-json all measured the same way, same engine, same run). Flagging this as a concern for whoever next relies on this control number outside of a same-engine ratio comparison.
