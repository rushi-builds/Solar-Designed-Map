# FORENSIC DIAGNOSTIC REPORT — Resource Dataset 2021-01-01 → 2025-12-31

**Status:** READ-ONLY DIAGNOSIS. NOT DEPLOYED. DATASET STATUS NOT CHANGED.
**Do not run any of the deployment steps in the other guides until this report is approved.**

---

## 0. Scope and honesty notice

- The original claim (“all 20,853 missing values are night albedo”) is **NOT proven**.
- My earlier `nightAlbedoCount = 21,912` came from a **synthetic reproduction** (assumed night = UTC hours 00 and 13–23), **not from the actual stored dataset**.
- The actual dataset lives in Cloudflare D1. This sandbox has **no D1 credentials/network**, so the definitive numbers must come from a **read-only query run by you** against your deployed D1.
- This report gives: (a) every read-only query/step, (b) the reconciliation framework for the 1,059-record difference, (c) an exact read-only Node script. **Nothing here changes stored values, fills data, adds sources, changes aggregation, changes schema, or changes status.**

---

## 1. Dataset fact sheet (from the code, not assumptions)

| Item | Value |
|---|---|
| Period | 2021-01-01 → 2025-12-31 |
| Days | 2021:365 + 2022:365 + 2023:365 + 2024:366 + 2025:365 = **1,826 days** |
| Hourly records (per parameter) | 1,826 × 24 = **43,824** |
| Parameters | 10 (ALLSKY_SFC_SW_DWN, DNI, DIFF, T2M, WS10M, WD10M, RH2M, PS, PRECTOTCORR, ALLSKY_SRF_ALB) |
| Total cell values | 43,824 × 10 = **438,240** |
| Reported missing | **20,853** → 20,853 / 438,240 = **4.757%** |
| Reported invalid / duplicates / unit mismatches | 0 / 0 / 0 |
| Original stored status | `REVIEW` (status must remain REVIEW until approved) |

---

## 2. The 1,059-record discrepancy (not yet explained — must be closed)

```
Original missing count       = 20,853
Synthetic nightAlbedoCount   = 21,912   (from my test, NOT from real data)
Difference                   =  1,059   ← unexplained
```

**Plausible explanations (each must be confirmed by the queries below):**

1. **Boundary hours.** Real sunrise/sunset in UTC for this grid are NOT exactly h=00 and h≥13. On many days hour 12 or hour 01 has `ALLSKY_SFC_SW_DWN > 0` (daylight). At those hours NASA can still return albedo `-999`. Such timestamps are **not “night”** and must remain **missing** (they are real retrieval gaps / dawn-dusk edge cases). With ~1,826 days, a net ~1,059 such hours is entirely plausible.
2. **Real non-albedo gaps.** Some of the 20,853 may be genuine missing values in another parameter (e.g. CERES solar gaps over partial days). The per-parameter query below settles this in one shot.
3. **Daytime albedo gaps.** Albedo (`ALLSKY_SRF_ALB`) is derived from GEWEX SRB + CERES SYN1deg; when DWN>0 but albedo is fill, that is a genuine missing albedo observation.

**The rule that must be validated, not assumed:** classify albedo fill as “N/A (night)” **only** when, at that same timestamp, `ALLSKY_SFC_SW_DWN == 0`. Any albedo fill with `DWN > 0` or `DWN` missing/invalid remains a true missing record.

---

## 3. Which parameter has the 20,853? (quick, read-only — run THIS first)

Every monthly summary already stores per-parameter missing/invalid counts inside `summary_json`. Nothing needs to be recomputed or written.

**Query 1 — totals from `resource_chunks` (independent store):**
```sql
SELECT COUNT(*) AS chunks,
       SUM(record_count)   AS expected_records,
       SUM(missing_count)  AS stored_missing,
       SUM(invalid_count)  AS stored_invalid,
       SUM(duplicate_count) AS stored_duplicates
FROM resource_chunks
WHERE request_id = '<REQUEST_ID>';
```

**Query 2 — totals from `resource_chunk_summaries`:**
```sql
SELECT SUM(CAST(json_extract(summary_json,'$.missingCount')      AS INTEGER)) AS missing_total,
       SUM(CAST(json_extract(summary_json,'$.invalidCount')      AS INTEGER)) AS invalid_total,
       SUM(CAST(json_extract(summary_json,'$.duplicateCount')    AS INTEGER)) AS duplicate_total,
       SUM(CAST(json_extract(summary_json,'$.unitMismatchCount') AS INTEGER)) AS unit_mismatch_total,
       SUM(CAST(json_extract(summary_json,'$.reviewCount')       AS INTEGER)) AS review_total,
       SUM(CAST(json_extract(summary_json,'$.nightAlbedoCount')  AS INTEGER)) AS night_albedo_total
FROM resource_chunk_summaries
WHERE request_id = '<REQUEST_ID>';
```
(`night_albedo_total` will be 0/NULL today because v1.9 is not deployed — expected.)

**Query 3 — per-parameter accounting (THE key query):**
```sql
SELECT je.key AS parameter,
       COUNT(*)                                          AS months,
       SUM(CAST(json_extract(je.value,'$.count')    AS INTEGER)) AS valid_count,
       SUM(CAST(json_extract(je.value,'$.missing')  AS INTEGER)) AS missing_count,
       SUM(CAST(json_extract(je.value,'$.invalid')  AS INTEGER)) AS invalid_count
FROM resource_chunk_summaries,
     json_each(json_extract(summary_json,'$.metrics')) AS je
WHERE request_id = '<REQUEST_ID>'
GROUP BY je.key
ORDER BY missing_count DESC, je.key;
```
D1 supports JSON1 (`json_extract` / `json_each`) — official Cloudflare D1 docs confirm this.

**How to run (both are read-only):**
- Cloudflare dashboard → D1 → SQL console, or
- `npx wrangler d1 execute <DB_NAME> --remote --command "<query>" --json`

**Interpretation table for Query 3:**

| Result | Meaning |
|---|---|
| `ALLSKY_SRF_ALB.missing = 20,853` and all other params `0` | All missing = albedo fill. Proceed to Section 4 (albedo cross-check) to split night vs real. |
| Several params have missing | The 20,853 is **not** all albedo. Report each; no classification change until known. |
| `ALB.missing + ALB.invalid ≈ 20,853 - (other params)` | Combined albedo evidence; still need Section 4. |

---

## 4. Albedo cross-check on the actual validated hourly rows (read-only)

The validated rows are stored in `resource_chunks.data_json` as
`{schema, timeStandard, parameterOrder, rows: [[timestamp, ...10 values], ...]}` where a value is `null` for fill/invalid (and, in old data, also night fill).

**Export (read-only, one command):**
```bash
npx wrangler d1 execute <DB_NAME> --remote \
  --command "SELECT request_id, chunk_start, data_json FROM resource_chunks WHERE request_id='<REQUEST_ID>' ORDER BY chunk_start" \
  --json > chunks_export.json
```
Then run the read-only forensic script (new file, does not touch anything):
```bash
node k12/forensic_readonly.mjs chunks_export.json "<REQUEST_ID>"
```

The script reports, **for every parameter**: expected / valid / null(missing+invalid as stored) / invalid (from existing counts) / duplicates (existing counts) / missing %.

For `ALLSKY_SRF_ALB` specifically it reports:
1. total null values in the ALB column;
2. for every null ALB timestamp: how many have `ALLSKY_SFC_SW_DWN == 0` (candidate night N/A);
3. how many have `DWN > 0` (real daytime albedo gap → must stay missing);
4. how many have `DWN == null` (DWN itself unavailable → must stay missing);
5. **exact timestamps** of every albedo-null where `DWN > 0` (also saved to `albedo_dayfill_timestamps.txt`);
6. continuous runs: count of runs, runs ≥ 24 h, and the longest run, split by category.

---

## 5. Reconciliation — exact formula the script applies

```
expected_records         = 43,824
original missing_total   = 20,853

albedo:   null_total      = ALB.missing + ALB.invalid          (from rows)
          night_count     = ALB nulls where DWN == 0
          dayfill_count   = ALB nulls where DWN > 0
          dwn_nulls       = ALB nulls where DWN == null
other_params_missing     = SUM(metrics[p].missing) for p != ALB

CHECK #1: missing_total == SUM(metrics[p].missing) over all p      (must hold)
CHECK #2: albedo null_total == ALB.missing + ALB.invalid            (must hold)
CHECK #3: null_total == night_count + dayfill_count + dwn_nulls     (must hold)

EXPECTED EXPLANATION OF THE 1,059:
  20,853 = other_params_missing + dayfill_count + dwn_nulls + night_count
  night_count  → what v1.9 will classify N/A (NOT missing)
  1,059        → 20,853 observed - (synthetic 21,912 assumed night) —
                 exactly the amount the boundary/dayfill/other-param evidence
                 must account for. The script prints all components.
```

**Verdict rule (for you, after running):**
- If `dayfill_count + dwn_nulls + other_params_missing > 0` → **status must stay REVIEW**; only the portion where `DWN == 0` may be reclassified as N/A (after approval).
- If `dayfill_count = 0`, `dwn_nulls = 0`, `other_params_missing = 0`, and `night_count = 20,853` → then (and only then) all 20,853 are truly night albedo → `VALID` may be considered.
- **No status change, no code change, no deploy, no fill, no new source, no schema change** until this reconciliation is closed and approved.

---

## 6. What I verified so far (real, but limited)

- **Live 1-day probe** (Pune grid 73.857/18.52, 2025-01-01, exact 10 params + `community=RE`, UTC): hour 00 and hours 13–23 have `ALLSKY_SFC_SW_DWN = 0` and `ALLSKY_SRF_ALB = -999`; hours 01–12 have valid DWN and valid albedo; all 9 other parameters complete. → It proves **night albedo exists**, it does **not** prove the full-set composition.
- This probe is 1 day at 1 grid cell; it cannot quantify the 5-year dataset. That quantification is Section 3–5.

---

## 7. Files

- This report: `k12/FORENSIC_REPORT.md` (read-only)
- Forensic script: `k12/forensic_readonly.mjs` (read-only; parses the export, writes only `albedo_dayfill_timestamps.txt`)
- **Unchanged / untouched:** Worker code, VBA, DRAWING_DATA, MAP/SAVE/ACK, calculation engine, D1 schema, any status value.

---

## 8. STOP — WAITING FOR APPROVAL

After the queries/script provide the reconciliation numbers, **stop and wait for approval** before any status change or code change.
