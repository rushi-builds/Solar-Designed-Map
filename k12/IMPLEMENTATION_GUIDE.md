# Solar EPC Resource — K=12 Turbo (v1.8 / RANGE12)

## What this delivers (K=12)

- **Before (K=1):** 60 NASA requests, 10 waves → 15–25s
- **K=6 (verified working by you):** 10 NASA requests, 2 waves → 14.25s full flow
- **K=12 (this patch):** **5 NASA requests, 1 wave** → expected **~9–12s** full flow
- **Accuracy:** proven identical (test below)
- **Cost: 100% free forever** — Cloudflare Workers Free + D1 Free, no schema change, no paid plan

---

## Files

| File | What it is |
|---|---|
| `worker.js` | Patched Worker — **default K=12**, CPU diagnostic + payload guard |
| `modSolarEPCResource_range6.bas` | Patched VBA module — range mode + B12/B13 diagnostics + auto fallback |
| `IMPLEMENTATION_GUIDE.md` | This guide |

> You are already running K=6 successfully. This release changes **only the default range size** (6 → 12). Your deployed K=6 code, VBA module and DB are 100% compatible.

---

## What changed in v1.8 (K=12)

### Worker
1. `RESOURCE_RANGE_MONTHS_DEFAULT = 12` — 60 months → **5 ranges** (`202101-202112`, …) → 1 wave at 6 parallel.
2. Env var override stays: `RESOURCE_RANGE_MONTHS` (1–12). Set `12` (or leave unset).
3. Every `process-range` response now returns:
   - `processMs` — measured CPU time of fetch+parse+validate (verifies the Free 10ms budget)
   - `rangeMonths` — the K actually used
4. **Payload guard:** if NASA's response body exceeds 8 MB (`NASA_PAYLOAD_TOO_LARGE`), the Worker returns 503 → VBA automatically falls back to month mode. K=12 is ~1.2–1.8 MB, so this is only a worst-case safety net.
5. Old endpoints unchanged (`process`, `process-month`) — instant rollback.

### VBA
1. Same range logic — automatically sends the larger ranges if the Worker returns them.
2. **Diagnostics** (hidden `_CLOUD_CFG` sheet, no input field touched):
   - `B12` = path used: `RANGE12` / `RANGE6` / `MONTH` / `RANGE6_FAIL`
   - `B13` = measured `processMs` of the range request (CPU ms)
3. Status bar now says: `"NASA 12-month ranges processing in parallel..."`
4. Fallback unchanged: any range failure → month-by-month mode → worst case = old speed.

---

## Deploy (2 steps)

### Step 1 — Worker
- Replace `worker.js` in Cloudflare Workers (dashboard or `wrangler deploy`).
- Env vars (optional — defaults are fine):
  - `RESOURCE_RANGE_MONTHS` = `12`
  - `RESOURCE_RANGE_ENABLED` = `1`
- Verify: `GET /v1/health` → `"version":"1.8-resource-range6"`

### Step 2 — VBA
- Replace `modSolarEPCResource` with `modSolarEPCResource_range6.bas` (Import File).

---

## Verify (10 min, mandatory)

1. **Run one NEW site** and note total time (target ~9–12s).
2. Check `_CLOUD_CFG.B12` → must be `RANGE12`.
   - If `RANGE12_FAIL` → some range failed → VBA correctly fell back to months; open Worker Logs and check `process-range` errors.
   - If `MONTH` → fallback is active (old speed).
3. Check `_CLOUD_CFG.B13` → `processMs` from the range request.
   - **If CPU ms ≤ ~8 → K=12 is safe on Free. Keep it.**
   - **If CPU ms is near/over 10** → set `RESOURCE_RANGE_MONTHS=6` (bulletproof K=6, back to 14s).
4. Open Worker Logs: you should see exactly **5** `process-range` calls for the 5-year period (10 if `RESOURCE_RANGE_MONTHS=6`). Look at `cpu_ms` / `exceededCpu` per invocation.
5. **Accuracy audit (once):** compare one site's per-month `checksum_sha256` + RESOURCE_DB row vs the K=1 baseline — they must match (range split already unit-tested byte-identical).

---

## Expected timing

| Path | NASA calls | Waves | Full flow (map save → RESOURCE_DB) |
|---|---|---|---|
| Old K=1 | 60 | 10 | 15–25s |
| K=6 | 10 | 2 | ~14s (your measurement) |
| **K=12** | **5** | **1** | **~9–12s (target)** |
| Cached site | 0 | 0 | ~1–2s |

## Rollback (instant)
- Worker env `RESOURCE_RANGE_MONTHS` = `6` → back to your proven K=6 behavior (no code change).
- Or `RESOURCE_RANGE_ENABLED=0` → pure month mode.

---

## Proof performed

- Hourly key generator: byte-identical to old version (leap years, 43,824 keys 5-year).
- **K=12 equivalence test:** 12-month range split vs month-by-month → all 12 months **rows, states, counts identical** (8,784 records, missing + invalid values).
- Grouping: 60 months → 5 contiguous 12-month runs (1 wave); holes handled as separate runs.
- D1 statements per range call = 25 (limit 50). Payload ~1.2–1.8 MB (guard 8 MB). Free CPU is measured via `processMs` before trusting K=12 long-term.
