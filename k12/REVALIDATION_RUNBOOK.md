# RUNBOOK — Activate v1.9 Night-Albedo N/A & Re-Validate Request `1dce0c1c-0306-4765-97f8-567a49bed34b`

**Approved after FINAL_FORENSIC_REPORT.md:** CHECK #1 (only ALB missing, 20,853) + CHECK #2 (20,853 ALB nulls, all with DWN=0; 0 daytime; 0 DWN-unavailable) — reconciliation `20,853 + 0 + 0 = 20,853`.

**Guarantees**
- `resource_chunks` (hourly values) — **never modified**. No fill → null, never zero, unchanged.
- No schema change. No API-parameter change. No VBA / MAP / DRAWING_DATA / SAVE / ACK / Calculation-Engine change.
- Only **classification metadata** is updated: `resource_requests.summary_state_json`, `resource_requests.summary_json`, `resource_chunk_summaries.summary_json`.
- The tool **hard-aborts** (no SQL) if any ALBEDO null has DWN > 0 or DWN unavailable — such a request stays REVIEW.
- The patch is **idempotent** — safe to re-run.

---

## STEP 0 — Pre-flight (worker deploy + backups, user-side with Cloudflare auth)

> This sandbox has **no Cloudflare/D1 authentication** (`wrangler whoami` is unauthenticated). Steps 0–4 run on your machine / dashboard.

### 0.1 Deploy v1.9 Worker (so all FUTURE ingests use the N/A rule)
- Cloudflare dashboard → Workers → your resource worker → **Edit code** → replace with `k12/worker_k12.js` (or `resource-range6/worker.js` — identical v1.9 logic).
- Keep the existing **D1 binding (env name `DB`)** and existing env vars untouched (`WORKBOOK_KEY`, `RESOURCE_RANGE_MONTHS`, `RESOURCE_RANGE_ENABLED`).
- Save & Deploy.
- Verify: `GET /v1/health` → `"version":"1.9-nightalbedo"`.

### 0.2 Take a read-only backup (one command, before any write)
```bash
npx wrangler d1 execute <DB_NAME> --remote --command \
  "SELECT request_id,workbook_id,summary_state_json,summary_json,updated_at FROM resource_requests WHERE request_id='1dce0c1c-0306-4765-97f8-567a49bed34b'" \
  --json > backup_request_1dce0c1c.json

npx wrangler d1 execute <DB_NAME> --remote --command \
  "SELECT request_id,chunk_start,summary_json FROM resource_chunk_summaries WHERE request_id='1dce0c1c-0306-4765-97f8-567a49bed34b' ORDER BY chunk_start" \
  --json > backup_chunksummaries_1dce0c1c.json
```

---

## STEP 1 — Read-only exports (no writes)

```bash
npx wrangler d1 execute <DB_NAME> --remote --command \
  "SELECT request_id,chunk_start,parameter_order_json,data_json FROM resource_chunks WHERE request_id='1dce0c1c-0306-4765-97f8-567a49bed34b' ORDER BY chunk_start" \
  --json > chunks_1dce0c1c.json

npx wrangler d1 execute <DB_NAME> --remote --command \
  "SELECT * FROM resource_requests WHERE request_id='1dce0c1c-0306-4765-97f8-567a49bed34b'" \
  --json > requests_1dce0c1c.json
```

## STEP 2 — Run the re-validation replay (dry-run — writes NOTHING)

```bash
node k12/revalidate_night_albedo.mjs chunks_1dce0c1c.json requests_1dce0c1c.json 1dce0c1c-0306-4765-97f8-567a49bed34b --out . --expect-chunks 60 --expect-rows 43824
```

Expected (must ALL be true before applying):
- rows replayed = **43,824**; problems = 0
- ALBEDO null total = **20,853**; NIGHT = **20,853**; DAYTIME = **0**; DWN_UNAVAILABLE = **0**
- reconciliation `20,853 + 0 + 0 = 20,853`
- `VALUE FIELDS UNCHANGED? YES` (all means/min/max identical)
- GATES: all `OK`
- Files produced: `revalidation_report_1dce0c1c-0306-4765-97f8-567a49bed34b.txt` + `revalidate_1dce0c1c-0306-4765-97f8-567a49bed34b.sql`

If any gate is `FAIL` → **stop.** No SQL should exist for the request; it must remain REVIEW.

## STEP 3 — Review the generated report, then apply (write, metadata only)

```bash
cat revalidation_report_1dce0c1c-0306-4765-97f8-567a49bed34b.txt
npx wrangler d1 execute <DB_NAME> --remote --file revalidate_1dce0c1c-0306-4765-97f8-567a49bed34b.sql
```

## STEP 4 — Verify after apply

```bash
npx wrangler d1 execute <DB_NAME> --remote --command \
  "SELECT request_id, json_extract(summary_json,'$.status') AS data_status, \
          json_extract(summary_json,'$.missingCount') AS missing, \
          json_extract(summary_json,'$.nightAlbedoCount') AS night, \
          json_extract(summary_json,'$.invalidCount') AS invalid \
   FROM resource_requests WHERE request_id='1dce0c1c-0306-4765-97f8-567a49bed34b'" --json

npx wrangler d1 execute <DB_NAME> --remote --command \
  "SELECT chunk_start, json_extract(summary_json,'$.nightAlbedoCount') AS night \
   FROM resource_chunk_summaries WHERE request_id='1dce0c1c-0306-4765-97f8-567a49bed34b' ORDER BY chunk_start" --json

npx wrangler d1 execute <DB_NAME> --remote --command \
  "SELECT je.key AS parameter, \
          SUM(json_extract(je.value,'$.count')) valid, \
          SUM(json_extract(je.value,'$.missing')) missing, \
          SUM(json_extract(je.value,'$.invalid')) invalid \
   FROM resource_chunk_summaries, json_each(json_extract(summary_json,'$.metrics')) je \
   WHERE request_id='1dce0c1c-0306-4765-97f8-567a49bed34b' GROUP BY je.key ORDER BY je.key" --json
```

Expected:
- data_status = **VALID**, missing = **0**, night = **20,853**, invalid = **0**
- per-parameter: every `missing = 0`; ALB valid = 22,971
- Summary endpoint (`GET /v1/resource/summary/1dce0c1c-0306-4765-97f8-567a49bed34b` with workbook auth) now reports **Data Status = VALID** with `night-albedo N/A=20853`.

## STEP 5 — RESOURCE_DB / Excel
No Excel structural change needed for this fix. Keep `Location` + `Latitude (°)` + `Longitude (°)` as already updated (Location descriptive only; lat/lon authoritative). The map will read `data_status=VALID` on next refresh of that resource row.

---

## Rollback (if ever needed)
Run the Step 0.2 backups back, or simply restore `summary_state_json` / `summary_json` / per-chunk `summary_json` from `backup_*.json` (same `wrangler d1 execute --file` style). Resource rows were never touched, so there is nothing to restore there.

## What changes for FUTURE requests
- Worker v1.9 (deployed in Step 0.1) classifies night-time albedo as N/A during ingest: missing/review counts exclude it, `nightAlbedoCount` appears in state/response/remarks, `data_status` may be VALID when only night-albedo fills existed.
- Hourly albedo values remain `null` (never zero); storage semantics unchanged.
- If a future request has ALBEDO null with DWN > 0 or DWN unavailable → it stays REVIEW (genuine gap) — this is enforced by the same rule, both month and range paths.
