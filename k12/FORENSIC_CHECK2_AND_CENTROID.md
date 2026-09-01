# READ-ONLY FORENSIC — CHECK #2 (ALB vs DWN) + CENTROID IDENTITY INVESTIGATION

**Date:** 2026-09-02 · **Branch:** `arena/01a05bd6-solar-designed-map` · **HEAD:** `e867027` (v1.9 parked, NOT deployed)

**Status of this document:** READ-ONLY FORENSIC. ZERO writes:
- no D1 INSERT/UPDATE/DELETE, no status change, no validation-rule change,
- no Worker deploy/code change, no VBA change, no MAP/DRAWING_DATA change,
- no Calculation-Engine change, no fill, no schema change, no cache change.

**Sandbox capability (verified):** `wrangler whoami` → *not authenticated*; no D1 credentials exist in this environment. Therefore **stored-row D1 counts below are produced by the user-side read-only commands in Part 2 — this sandbox cannot run them**. Everything else (worker source audit + live NASA POWER probes) was executed here, read-only.

---

# PART 1 — CHECK #2 : ALLSKY_SRF_ALB fill vs ALLSKY_SFC_SW_DWN

## 1.1 What is already settled
- Step 1 (user-run): **20,853 missing = 100% `ALLSKY_SRF_ALB`**; other 9 parameters complete.
- 2021-01-01 → 2025-12-31 = 1,826 days = 43,824 hourly records/parameter.

## 1.2 Method (must be run where D1 is accessible — SELECT only)
Layer A — SQL classification (no export needed):
```sql
WITH rec AS (
  SELECT je.value AS row
  FROM resource_chunks rc,
       json_each(json_extract(rc.data_json,'$.rows')) je
  WHERE rc.request_id = '1dce0c1c-0306-4765-97f8-567a49bed34b'
)
SELECT
  SUM(CASE WHEN json_extract(row,'$[10]') IS NULL THEN 1 ELSE 0 END) AS alb_missing_total,
  SUM(CASE WHEN json_extract(row,'$[10]') IS NULL AND json_extract(row,'$[1]') = 0 THEN 1 ELSE 0 END) AS night_na_candidate,
  SUM(CASE WHEN json_extract(row,'$[10]') IS NULL AND json_extract(row,'$[1]') > 0 THEN 1 ELSE 0 END) AS daytime_albedo_missing,
  SUM(CASE WHEN json_extract(row,'$[10]') IS NULL AND json_extract(row,'$[1]') IS NULL THEN 1 ELSE 0 END) AS dwn_unavailable
FROM rec;
```
Row layout (`data_json`): `[ts, DWN, DNI, DIFF, T2M, WS10M, WD10M, RH2M, PS, PRECTOTCORR, ALB]` → ALB = index 10, DWN = index 1. (`parameterOrder` in the same payload confirms this; the classifier derives indices from it defensively.)

Layer B — timestamps + continuous runs (export + classifier):
```bash
npx wrangler d1 execute <DB_NAME> --remote \
  --command "SELECT request_id,chunk_start,parameter_order_json,data_json \
             FROM resource_chunks \
             WHERE request_id='1dce0c1c-0306-4765-97f8-567a49bed34b' ORDER BY chunk_start" \
  --json > check2_chunks.json

node k12/forensic_check2.mjs check2_chunks.json 1dce0c1c-0306-4765-97f8-567a49bed34b
```
Outputs: the 3 category counts + reconciliation assert (`NIGHT + DAYTIME + DWN_UNAVAILABLE = 20,853`), first/last timestamp per category, all `DAYTIME_ALBEDO_MISSING` timestamps to `check2_daytime_albedo_missing.txt`, continuous runs (≥24 h count + longest) per category to `check2_runs.txt`.

## 1.3 Evidence collected READ-ONLY (supplementary — NOT a substitute for 1.2)

NASA POWER **v2.10.0** live probes (read-only GETs, same request shape as the Worker: `community=RE`, `format=JSON`, `time-standard=UTC`) at cell (17.99, 74.435):

**Winter, 2024-01-01 → 01-03 (72 h):**

| UTC hour | 00 | 01 | 02…11 | 12 | 13…23 |
|---|---|---|---|---|---|
| DWN (Wh/m²) | **0.0** | 5.53 | >0 | 12.05 | **0.0** |
| ALB | **-999** | 0.18 | 0.04–0.18 | 0.18 | **-999** |

→ Night = **12 h/day** (00 + 13–23); ALB fill occurs **exactly and only** at DWN = 0.

**Summer, 2024-07-01 → 07-03 (72 h):**

| UTC hour | 00 | 01…12 | 13 | 14…23 |
|---|---|---|---|---|
| DWN (Wh/m²) | **3.12 / 4.72 / 4.32** | >0 | **9.20 / 6.97 / 11.18** | **0.0** |
| ALB | **0.15 / 0.17 / 0.16** | 0.14–0.19 | **0.16 / 0.15 / 0.17** | **-999** |

→ Night = **10 h/day** (14–23 only); ALB fill occurs **exactly and only** at DWN = 0.

### 1.3.1 What this means for the 1,059 discrepancy
- The old synthetic rule "night = UTC 00 + 13–23" (12 h/day) is **seasonally wrong** — it is true in winter but not in summer, where UTC 00 and UTC 13 are still daylight (18:30 IST) with DWN > 0 and valid albedo.
- Fixed 12 h/day × 1,826 days = **21,912** (the earlier synthetic number).
- Actual = **20,853** → average actual night length = 20,853 / 1,826 = **11.4193 h/day** (i.e. 0.581 h/day less than the fixed rule; 0.581 × 1,826 = **1,061 ≈ 1,059**, within 2 on this sample model).
- **Hypothesis (NOT yet proof):** 1,059 = summer/dawn-dusk hours where DWN > 0 so NASA valid-ALB values exist instead of -999 fills — i.e. the fixed-hour rule overcounted; no true daytime albedo gap.
- **Proof gate:** the D1 classification (1.2). Expected for the hypothesis to hold: `NIGHT_NA_CANDIDATE = 20,853`, `DAYTIME_ALBEDO_MISSING = 0`, `DWN_UNAVAILABLE = 0`.
- If the D1 result shows `DAYTIME_ALBEDO_MISSING > 0` or `DWN_UNAVAILABLE > 0`, those records must be explained (they would be *real* gaps, blocking any VALID claim for the albedo path).

**Nothing is marked VALID, no status is changed, v1.9 stays parked.** The next action is the one read-only export above by the user; then STOP.

---

# PART 2 — IDENTICAL SUMMARIES ACROSS 3 CENTROIDS

## 2.1 The three observations
| label | requested centroid | user-observed identical values |
|---|---|---|
| A (request `1dce0c1c…`) | 17.99038747069333, 74.43525012125609 | T2M 24.691137 / min 7.47 / max 42.17 · WS10M 4.043636 · WD10M 315.773054 · RH2M 64.552159 · PS 93.621529 · PRECTOTCORR 3.202668 |
| B | 17.990776332796173, 74.43505256073746 | same |
| C | 18.18043580, 74.61004963 | same |

## 2.2 VERIFIED (this sandbox, read-only): NASA itself returns identical series

Read-only GETs to `power.larc.nasa.gov/api/temporal/hourly/point` (same params as Worker: T2M, WS10M, WD10M, RH2M, PS, PRECTOTCORR; community RE; UTC), machine-compared with `k12/forensic_payload_compare.mjs`:

| window | A vs C · values compared | result | A geometry | C geometry | elevation |
|---|---|---|---|---|---|
| 2024-01-01..02 (winter) | 288 (48 h × 6 params) | **0 differences — series SHA-256 identical** `a9df1b3a…` | [74.435, 17.99, **655.86**] | [74.61, 18.18, **655.86**] | identical |
| 2024-07-01..02 (monsoon) | 288 | **0 differences — series SHA-256 identical** `21b08476…` | [74.435, 17.99, **655.86**] | [74.61, 18.18, **655.86**] | identical |

- B (Jan 2024) returned the same hour-by-hour values as A/C with geometry `[74.435, 17.991, 655.86]`.
- **Controls (different cells):** (18.7, 74.5) → T2M 00h = 14.43 (A: 14.31), elevation **608.95**; (20.2, 74.5) → T2M 00h = 15.38, elevation **581.10**. Different values + different elevation ⇒ different 0.5° cells.

**Mechanism (confirmed by NASA's own response):** NASA POWER serves meteorology on a **0.5° × 0.5° MERRA-2 grid** (`header.sources = ["MERRA2","POWER"]`) and **albedo on the 1° SYN1DEG grid** (`["SYN1DEG","POWER"]`). Any requested centroid inside one cell returns the SAME gridded series; the API echoes the requested position rounded to 3 decimals (`geometry.coordinates`), so different centroids can share every value while echoing different coordinates. The three centroids all fall in the same cells (identical 655.86 m grid elevation is the giveaway). => **Identical summaries are NASA source-grid behavior, NOT a code reuse/leak bug.**

## 2.3 Code audit (read-only, `k12/worker_k12.js` ≡ `resource-range6/worker.js`)

| # | user question | finding | basis |
|---|---|---|---|
| 1 | NASA request coords used | `nasaUrl` sets `longitude=fixed8(centroid_lon)`, `latitude=fixed8(centroid_lat)` — 8-decimal values, from `resource_requests.centroid_lat/lon` | worker lines 423–424 |
| 2 | NASA returned coords/metadata | validated inside `validateAndNormalizeNasaMonth`: `geometry.coordinates` must be within **±0.01** of requested centroid (passes: NASA rounds to 3 dp); header api name/version, sources, fill_value, time_standard, start/end, parameter units all stored in `metadata_json` | lines 630–660, 850–860 |
| 3 | chunks per request | `resource_chunks`/`resource_chunk_summaries` keyed by `request_id`; every processed month/range INSERTS (idempotent `ON CONFLICT`) with `checksum_sha256` of the payload | lines 481–497, 1108–1111, 1276–1279 |
| 4 | rows per request | `record_count` per chunk; 10 params × hourly keys validated against expected month keys | lines 700–720 |
| 5 | summary isolation | `summary_state_json` and `summary_json` are **columns of `resource_requests`**, read/written with `WHERE request_id=? AND workbook_id=?`; `resourceSummary` selects the single row | lines 469, 499, 540–549 |
| 6 | accumulator/state reuse | **No module-level mutable state exists** (grep: all top-level bindings are `const`; accumulators are `newResourceSummaryState()` created per request; `mergeMonthState` only merges the request's own row state; range path builds per-request chunk summaries then merges into the request row). No `cache`/KV used anywhere in the resource path | whole-file audit |
| 7 | cache_key incl. centroid | **YES**: `cacheIdentity = {workbook, projectId, latitude: fixed8, longitude: fixed8, start, end, "HOURLY", product, parameters}` → SHA-256. The three centroids differ at 8 decimals ⇒ **three distinct cache keys**; dedupe can never serve one request for another | lines 353–366 |
| 8 | same payload reused across requests | Impossible via app cache (7); **same values are produced because NASA answers the same grid cell** — verified live (2.2). Expect identical `checksum_sha256` for identical (same-cell, same-range) real fetches — that is a fingerprint of the same NASA answer, not app reuse | 2.2 + 2.4 below |
| 9 | raw hourly comparison A vs C | Live NASA: **identical** (0 diffs, 576 values across two seasons). Stored rows: user runs SQL 2.6 / `k12/forensic_centroid_chunks.mjs` (expected identical for same month chunks) | this report |

## 2.4 Remaining D1 confirmations (user-side, SELECT only) — `k12/FORENSIC_QUERIES_READONLY.sql`
1. Find the request_id for 18.18043580, 74.61004963 (SQL 2.1).
2. Confirm 3 distinct `cache_key` + `polygon_hash` (SQL 2.2) → shows no dedupe collision.
3. Chunk/row inventory per request; `COUNT(DISTINCT checksum_sha256)` (SQL 2.3).
4. Confirm `summary_state_json`/`summary_json` exist per request row (SQL 2.5).
5. Confirm raw stored equality: SQL 2.6 (T2M/WS/RH/PS/PREC per timestamp) or export + `node k12/forensic_centroid_chunks.mjs chunks_req1.json chunks_reqC.json requests_meta.json` (SQL 2.7).

**Expected outcome, if the app is healthy:** per request = expected chunk/row counts; 3 different cache keys; summary rows isolated; stored raw values identical between request 1 and request C for matching months (because NASA returned the same cell); checksums identical for the same month chunks. **Conclusion then: identical summaries = correct NASA grid behaviour — not a defect, not state leakage.**

*Business consequence, if confirmed:* any two sites inside the same 0.5° (and 1° for albedo) NASA cell will legitimately share meteorology. If the product needs distinguishable values per site, that is a data-resolution product decision — NOT a bug, and no change is made here.

---

# PART 3 — VERDICT & NEXT ACTION

| item | status |
|---|---|
| CHECK #2 totals (NIGHT/DAYTIME/DWN_UNAVAILABLE + reconciliation) | **PENDING — one read-only export** (Part 1.2). Hypothesis: NIGHT=20,853, others=0. |
| 1,059 explanation | **Hypothesis stated** (seasonal night length: 12 h/day winter vs 10 h/day summer; 0.581 h/day × 1,826 ≈ 1,061 ≈ 1,059). **Confirmation requires the D1 classification.** |
| Centroid identity | **Mechanism confirmed** (NASA 0.5°/1° grid, same cell) + app-cache/state audit **clean**. D1 row-level check recommended (2.4) before closing. |
| v1.9 night-albedo N/A | **PARKED.** Not deployed, not enabled. Only after CHECK #2 closes with NIGHT=20,853 / rest=0 AND user approval. |
| Dataset status | **REVIEW — unchanged.** |

## 🛑 STOP — Waiting for user
1. Run Part 1.2 export + classifier (or paste SQL 1.1 output) → CHECK #2 verdict.
2. Run SQL 2.1 → request_id for 18.18043580,74.61004963, then optionally 2.2–2.7 → centroid closure.
3. No deployment / status change / fill / code change until you approve.

*Files added (diagnostic only, all read-only):* `k12/forensic_check2.mjs`, `k12/forensic_centroid_chunks.mjs`, `k12/forensic_nasa_probe.mjs`, `k12/forensic_payload_compare.mjs`, `k12/FORENSIC_QUERIES_READONLY.sql`, `k12/FORENSIC_REPORT.md`. Product files (worker, VBA, guide, zip) untouched — `git status` shows only new untracked diagnostic files.
