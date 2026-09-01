# FINAL READ-ONLY FORENSIC REPORT — Solar EPC NASA POWER Resource Dataset

**Date:** 2026-09-02 · **Request ID:** `1dce0c1c-0306-4765-97f8-567a49bed34b`
**Dataset period:** 2021-01-01 → 2025-12-31 (UTC) · **Branch:** `arena/01a05bd6-solar-designed-map` · **HEAD:** `e867027` (v1.9 exists but is **PARKED — not deployed**).

**Nature of this document:** READ-ONLY forensic + design specification. No data was changed, no API parameters were changed, no D1 schema was changed, no Worker was deployed, no status was flipped, no VBA / MAP / DRAWING_DATA / SAVE / ACK / Calculation-Engine code was touched. **STOP at the end — awaiting approval.**

---

## 0. Scope and ground truth

| fact | value |
|---|---|
| requested period | 2021-01-01 → 2025-12-31 |
| calendar days | 1,826 (365+365+365+366+365) |
| expected hourly records per parameter | 43,824 (= 1,826 × 24) |
| parameters | 10 (9 meteorological + `ALLSKY_SRF_ALB`) |
| expected total values | 438,240 (43,824 × 10) |
| checked request | `1dce0c1c-0306-4765-97f8-567a49bed34b` |
| prior status | REVIEW (unchanged) |

---

## 1. CHECK #1 — per-parameter validation (read-only, from `resource_chunk_summaries` / `resource_chunks`)

Counts are derived from the stored `summary_json.metrics` (summed via `json_each`/`json_extract`) and cross-checked against `resource_chunks.record_count` / `missing_count` / `invalid_count` / `duplicate_count`.

| # | parameter | expected | valid | missing/fill | invalid | duplicates | missing % |
|---|---|---|---|---|---|---|---|
| 1 | ALLSKY_SFC_SW_DWN | 43,824 | 43,824 | 0 | 0 | 0 | 0.0000% |
| 2 | ALLSKY_SFC_SW_DNI | 43,824 | 43,824 | 0 | 0 | 0 | 0.0000% |
| 3 | ALLSKY_SFC_SW_DIFF | 43,824 | 43,824 | 0 | 0 | 0 | 0.0000% |
| 4 | T2M | 43,824 | 43,824 | 0 | 0 | 0 | 0.0000% |
| 5 | WS10M | 43,824 | 43,824 | 0 | 0 | 0 | 0.0000% |
| 6 | WD10M | 43,824 | 43,824 | 0 | 0 | 0 | 0.0000% |
| 7 | RH2M | 43,824 | 43,824 | 0 | 0 | 0 | 0.0000% |
| 8 | PS | 43,824 | 43,824 | 0 | 0 | 0 | 0.0000% |
| 9 | PRECTOTCORR | 43,824 | 43,824 | 0 | 0 | 0 | 0.0000% |
| 10 | **ALLSKY_SRF_ALB** | **43,824** | **22,971** | **20,853** | 0 | 0 | **47.5835%** |
| — | **TOTAL** | **438,240** | **417,387** | **20,853** | 0 | 0 | **4.7584%** |

Additional recorded facts:
- unit-mismatch count = **0**; invalid stored = **0**; duplicate rows = **0**.
- `T2M_MIN` / `T2M_MAX` are derived from validated hourly `T2M` (NASA hourly has no T2M_MIN/MAX) — derivation, not stored NASA fields.
- The **entire 20,853 missing count belongs exclusively to `ALLSKY_SRF_ALB`** (CHECK #1, user-verified).

---

## 2. CHECK #2 — ALBEDO vs DWN actual-row reconciliation (CONFIRMED)

Row-level inspection of every stored row for `1dce0c1c-0306-4765-97f8-567a49bed34b` (all 43,824 records; ALB = column index 10, DWN = index 1 in the stored `parameterOrder`):

| category | count |
|---|---|
| total rows checked | 43,824 |
| total ALBEDO missing (stored null) | **20,853** |
| ALBEDO missing **AND** DWN = 0 → NIGHT_NA_CANDIDATE | **20,853** |
| ALBEDO missing **AND** DWN > 0 → DAYTIME_ALBEDO_MISSING | **0** |
| ALBEDO missing **AND** DWN unavailable/invalid → DWN_UNAVAILABLE | **0** |
| **reconciliation** | **20,853 + 0 + 0 = 20,853 ✓** |

**Conclusion of CHECK #2:** every single stored albedo gap occurs at an hour where `ALLSKY_SFC_SW_DWN = 0`. There are **zero** daytime albedo gaps and **zero** records with an unavailable DWN in this dataset. Every ALBEDO gap is a sun-below-horizon hour — physically not applicable, not true missing data.

### 2.1 Rejected value — must NOT be cited as evidence

| quantity | origin | status |
|---|---|---|
| 20,853 | actual stored D1 data (CHECK #1 + CHECK #2) | **EVIDENCE** |
| 21,912 | synthetic test with a fixed rule "night = UTC hours 00 & 13–23" | **NOT EVIDENCE — rejected, never cite** |

### 2.2 Explanation of the former 20,853 ↔ 21,912 gap (1,059)

NASA POWER albedo fill hours are **seasonal**, not a fixed clock rule (verified read-only against live NASA POWER v2.10.0 responses for this cell):

| season (sample) | UTC hours with DWN = 0 | night hours/day | ALBEDO |
|---|---|---|---|
| winter (2024-01-01..03) | 00, 13–23 | 12 h/day | -999 exactly at those hours; valid 01–12 |
| summer (2024-07-01..03) | 14–23 | 10 h/day | -999 exactly at those hours; **valid at UTC 00 and 13** (DWN 3.1–4.7 / 7.0–11.2 Wh/m²) |

- Fixed 12 h/day × 1,826 days = 21,912 (synthetic).
- Actual = 20,853 → true mean night length = 20,853 / 1,826 = **11.4193 h/day**; 0.581 h/day × 1,826 ≈ 1,061 ≈ **1,059**.
- **Conclusion:** the 1,059 difference = summer (and transitional) hours that are still daylight at UTC 00 / UTC 13 where NASA returns a *valid* albedo value instead of a fill — the fixed-hour synthetic rule overcounted. There are no hidden daytime albedo gaps (confirmed: DAYTIME_ALBEDO_MISSING = 0).

---

## 3. Four distinct states — the taxonomy that governs every value

| state | definition | representation | accounting |
|---|---|---|---|
| **A. Raw NASA fill/null** | NASA response contains `-999.0` (`header.fill_value`), `null`, or an absent key for a timestamp | source-side only; never stored as `-999` | — |
| **B. Internal stored null** | the Worker normalizes every raw fill/non-finite value to `null` in `resource_chunks.data_json` rows | SQL `NULL` in the row array | counted as stored null; **never `0`** |
| **C. N/A classification** | semantic classification applied to a stored null **only when** the same-time `ALLSKY_SFC_SW_DWN = 0` → sun below horizon, albedo not defined | storage unchanged (still `NULL`); classification lives in counters/status/remarks | excluded from `missingCount`, excluded from albedo mean, excluded from REVIEW trigger |
| **D. Genuine missing daytime data** | ALBEDO stored null while `DWN > 0` at the same timestamp → the source failed to deliver a value during sun-up | storage unchanged | **must** count as missing → REVIEW |

**Why N/A albedo must remain NULL and must never become 0:**
1. **Zero is a physical value.** `ALLSKY_SRF_ALB = 0` means "surface reflects 0% of shortwave" — a real, measurable state. Substituting 0 for "no data" would inject a false physical reading.
2. **It corrupts aggregates.** The mean albedo feeds energy-balance context; zeros would drag the mean toward 0 and misrepresent the site.
3. **It destroys missing-data visibility.** A stored zero is indistinguishable from a genuine zero measurement; NULL preserves the distinction between "not applicable / not observed" and "measured value".
4. **The existing invariant is preserved:** missing → null, never zero (already established).
5. **It is auditable.** Status, per-parameter counters and remarks reflect the N/A classification; storage remains truthful to NASA payload.

---

## 4. The documented validation rule (design — implementation NOT activated)

```
ALLSKY_SRF_ALB → N/A (not-applicable, excluded from missing/review/albedo mean)
  IFF  (ALBEDO is raw fill/null AND stored as null)
   AND (ALLSKY_SFC_SW_DWN at the SAME timestamp = 0)

Otherwise, when ALBEDO is fill/null:
  IF  DWN > 0                     → GENUINE MISSING (daytime)  → missingCount++, REVIEW
  IF  DWN null / invalid / <0     → DWN_UNAVAILABLE            → REVIEW
  (DWN itself is normalized to null when NASA returns fill/invalid; null DWN is never treated as 0.)
```

Status consequences (documented rule, not applied):
- Night N/A hours: `nightAlbedoCount` (or equivalent counter) is informational; not part of missing/review.
- Any `DAYTIME_ALBEDO_MISSING > 0` → REVIEW (genuine gap).
- Any `DWN_UNAVAILABLE > 0` → REVIEW (cannot prove N/A).
- Only when no genuine missing/invalid/duplicate/unit-mismatch remains in **any** path can VALID be considered — and only after explicit user approval.

**Current dataset status: REVIEW — unchanged. v1.9 remains parked and unapproved.**

---

## 5. Invariants that must be preserved (nothing changes in this report)

| # | invariant | status |
|---|---|---|
| 9 | aggregation methodology (arithmetic means, circular wind mean, min/max extrema, daily-irradiance totals from complete 24 h days, GHI/DNI/DHI kWh/m²/day) | **unchanged** |
| 10 | stored resource values (`resource_chunks.data_json`, `resource_chunk_summaries`) | **unchanged** |
| 11 | API request parameters (10-parameter set, community RE, format JSON, time-standard UTC, lat/lon = exact centroid, period) | **unchanged** |
| 12 | D1 schema (tables, columns, keys) | **unchanged** |
| 13 | MAP / DRAWING_DATA / SAVE / ACK / VBA / Calculation Engine | **unchanged** |

---

## 6. Resource identity & traceability specification (Location column rule)

### 6.1 RESOURCE_DB column structure (Excel — already updated user-side)

The structure retains **BOTH**:
- `Location` — **additional descriptive field only** (human-readable site label)
- `Latitude (°)` — authoritative exact latitude
- `Longitude (°)` — authoritative exact longitude

No column is removed, renamed, replaced, or repurposed; `Location` is not to be duplicated (it already exists — do **not** add another).

| Project ID | Location | Latitude (°) | Longitude (°) | Source |
|---|---|---|---|---|
| Rushi | Site A | 17.990387 | 74.435250 | NASA POWER |
| Rushi | Site B | 17.688923 | 74.006176 | NASA POWER |
| Abhi | Site A | 17.990402 | 74.435105 | NASA POWER |
| Abhi | Site B | 18.180436 | 74.610050 | NASA POWER |

### 6.2 Identity rules (specification — no implementation change yet)

1. **Project ID / Project Name is a human/business identifier, NOT an identity for resource data.** It may legitimately recur across different site locations or different resource requests.
2. **The authoritative computational site location is the exact centroid:** `centroid_lat` / `centroid_lon` (Worker sends `fixed8(centroid)` to NASA and includes it in cache identity).
3. **`Location` is descriptive**; it must never replace, override, or substitute the exact coordinates.
4. **Never assume "same Project ID = same resource/weather data".** Same Project ID + different coordinates ⇒ separate resource locations/requests.
5. **Never assume "different coordinates = different weather values".** NASA POWER is gridded ⇒ different coordinates may legitimately resolve to the same underlying grid data ⇒ identical values are **not** evidence of a cache/state bug.
6. **Never merge, overwrite, reuse, or substitute resource results solely because Project ID/Name or Location is identical.**
7. **Cache/request identity must remain keyed to:** exact rounded centroid coordinates + dataset/product + period + resolution + parameter set + existing cache-identity fields + polygon/site identity where already implemented (current implementation: `{workbook, projectId, latitude: fixed8, longitude: fixed8, start, end, "HOURLY", product, parameters}` → SHA-256; `polygon_hash` stored separately).

### 6.3 Required traceability (each resource result must remain traceable to)

1. Project ID / Project Name
2. Location
3. Exact site centroid latitude
4. Exact site centroid longitude
5. NASA dataset/product
6. NASA returned/grid provenance where available
7. Period
8. Resolution
9. Parameter set
10. Existing cache identity

---

## 7. Provenance note — NASA grid resolution (previously verified)

- NASA POWER serves meteorological parameters on the **0.5° × 0.5° MERRA-2 grid** (`header.sources = ["MERRA2","POWER"]`) and albedo from the **1° SYN1DEG grid** (`["SYN1DEG","POWER"]`).
- Different site coordinates **can legitimately resolve to the same underlying NASA POWER/MERRA-2 grid data**, producing identical wind, temperature, humidity, pressure, etc.
- Verified read-only (v2.10.0): centroids 17.99038747/74.43525012 and 18.18043580/74.61004963 returned byte-identical hourly series for T2M, WS10M, WD10M, RH2M, PS, PRECTOTCORR (576 values across two seasons, 0 differences, identical grid elevation 655.86 m) while a different-grid control point returned different values and elevation.
- **Conclusion:** identical resource values at different coordinates are **NOT automatic evidence of cache/state leakage**. The app-level audit separately confirmed: cache_key includes exact coordinates (different keys for different centroids), no module-level mutable state, summaries isolated per `request_id`. No defect established; no changes made.

---

## 8. Verdict

| item | result |
|---|---|
| CHECK #1 | PASS — only `ALLSKY_SRF_ALB` has missing (20,853); everything else complete; 0 invalid / duplicate / unit mismatch |
| CHECK #2 | PASS — 20,853 ALBEDO nulls, all with DWN = 0; 0 daytime; 0 DWN-unavailable; reconciliation = 20,853 |
| N/A rule | documented (Section 4); **not activated** |
| v1.9 | **parked** — not deployed, not enabled |
| status | **REVIEW — unchanged** |
| implementation / schema / API / Excel columns | **unchanged** (Location column exists user-side; not duplicated) |

## 🛑 STOP — waiting for approval
No deployment, no status flip, no fill, no code, no schema, no MAP/DRAWING_DATA/VBA/Engine changes. Next step (only if you approve): implement/activate the N/A classification per the Section-4 rule, then re-validate; until then, nothing further is executed.
