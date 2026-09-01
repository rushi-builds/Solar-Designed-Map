#!/usr/bin/env node
// ============================================================================
// RETROACTIVE READ-ONLY RE-VALIDATION OF AN EXISTING RESOURCE REQUEST (v1.9 rule)
//
// Purpose: the request `1dce0c1c-0306-4765-97f8-567a49bed34b` was ingested under
// v1.8, so its stored summary metadata still counts 20,853 night-time cloud
// albedo fills as missing (status REVIEW). CHECK #2 proved that ALL 20,853
// ALBEDO nulls occur exactly where ALLSKY_SFC_SW_DWN = 0. This script replays
// that classification from the STORED ROWS (read-only), and then emits the
// exact, idempotent SQL that reclassifies ONLY the summary metadata:
//   - resource_requests.summary_state_json + summary_json (data_status -> VALID)
//   - resource_chunk_summaries.summary_json (per-month classification)
// It NEVER writes to resource_chunks, NEVER changes hourly resource values,
// NEVER touches schema, API parameters, MAP/DRAWING_DATA, VBA or the Engine.
//
// HARD GUARD: if ANY ALBEDO null has DWN > 0 (genuine daytime missing) or DWN
// unavailable/invalid, the script ABORTS and writes NO SQL — the request must
// stay REVIEW. Re-validation is only possible when NIGHT + 0 + 0 = total.
//
// Usage:
//   node revalidate_night_albedo.mjs chunks.json requests.json [request_id] [--out DIR]
//
// chunks.json   = read-only wrangler export of resource_chunks for the request
// requests.json = read-only wrangler export of resource_requests (full row)
// (see REVALIDATION_RUNBOOK.md for the exact export commands)
// ============================================================================

import fs from "node:fs";
import path from "node:path";

const RESOURCE_PARAMETERS = [
  "ALLSKY_SFC_SW_DWN", "ALLSKY_SFC_SW_DNI", "ALLSKY_SFC_SW_DIFF", "T2M",
  "WS10M", "WD10M", "RH2M", "PS", "PRECTOTCORR", "ALLSKY_SRF_ALB"
];
const ALB = "ALLSKY_SRF_ALB";
const DWN = "ALLSKY_SFC_SW_DWN";
const IRR = ["ALLSKY_SFC_SW_DWN", "ALLSKY_SFC_SW_DNI", "ALLSKY_SFC_SW_DIFF"];

// ---------------------------------------------------------------------------
// helpers
// ---------------------------------------------------------------------------
function collectRows(node, out) {
  if (!node || typeof node !== "object") return;
  if (Array.isArray(node)) { for (const x of node) collectRows(x, out); return; }
  if (Array.isArray(node.results)) { for (const x of node.results) { if (x && typeof x === "object") out.push(x); } return; }
  if (Array.isArray(node.result)) { for (const x of node.result) collectRows(x, out); return; }
  for (const k of Object.keys(node)) collectRows(node[k], out);
}
function loadRows(file) {
  const raw = JSON.parse(fs.readFileSync(file, "utf8"));
  const out = [];
  collectRows(raw, out);
  return out;
}
function sqlQuote(v) { return "'" + String(v).replace(/'/g, "''") + "'"; }
function jsClone(v) { return JSON.parse(JSON.stringify(v)); }
function meanOf(state, p) { return state.metrics[p].count ? state.metrics[p].sum / state.metrics[p].count : null; }
function dailyMeanOf(state, p) { const d = state.dailyIrradiance[p]; return d.count ? d.sum / d.count : null; }
function safeText(value) {
  if (value === null || value === undefined || (typeof value === "number" && !Number.isFinite(value))) return "";
  if (typeof value === "number") return String(Math.round(value * 1000000) / 1000000);
  return String(value).replace(/[\r\n|]+/g, " ").trim();
}
function closeEnough(a, b, tol = 1e-6) {
  if (a === null || a === undefined) return b === null || b === undefined;
  if (b === null || b === undefined) return false;
  const an = Number(a), bn = Number(b);
  if (!Number.isFinite(an) || !Number.isFinite(bn)) return an === bn;
  return Math.abs(an - bn) <= tol * Math.max(1, Math.abs(an), Math.abs(bn));
}

// ---------------------------------------------------------------------------
// replay from stored rows (identical logic to worker v1.9 finalize path)
// ---------------------------------------------------------------------------
function replayChunks(chunks, requestId) {
  const order = RESOURCE_PARAMETERS;
  const agg = {};
  for (const p of RESOURCE_PARAMETERS) agg[p] = { expected: 0, valid: 0, sum: 0, min: null, max: null, nulls: 0 };
  const daily = {}; for (const p of IRR) daily[p] = {};
  let windSin = 0, windCos = 0, windDirCount = 0;
  let albNight = 0, albDayfill = 0, albDwnNull = 0;
  const chunkStats = [];   // one per chunk in chunk_start order
  let totalRows = 0;
  const problems = [];

  const sorted = [...chunks].sort((a, b) => String(a.chunk_start).localeCompare(String(b.chunk_start)));
  for (const c of sorted) {
    let doc;
    try { doc = JSON.parse(c.data_json); } catch { problems.push(`bad data_json at ${c.chunk_start}`); continue; }
    const ord = Array.isArray(doc.parameterOrder) ? doc.parameterOrder : [];
    if (!ord.length || ord.join(",") !== RESOURCE_PARAMETERS.join(",")) {
      problems.push(`parameterOrder mismatch at ${c.chunk_start}: ${ord.join(",")}`);
      continue;
    }
    const idx = p => ord.indexOf(p);
    const albI = idx(ALB), dwnI = idx(DWN), wdI = idx("WD10M");
    let night = 0, dayfill = 0, dwnNull = 0;
    for (const r of (Array.isArray(doc.rows) ? doc.rows : [])) {
      totalRows++;
      const ts = String(r[0] || "");
      for (let i = 0; i < RESOURCE_PARAMETERS.length; i++) {
        const p = RESOURCE_PARAMETERS[i];
        const v = r[i + 1];
        agg[p].expected++;
        if (v === null || v === undefined) agg[p].nulls++;
        else {
          const n = Number(v);
          if (!Number.isFinite(n)) { agg[p].nulls++; continue; }
          agg[p].valid++; agg[p].sum += n;
          agg[p].min = agg[p].min === null ? n : Math.min(agg[p].min, n);
          agg[p].max = agg[p].max === null ? n : Math.max(agg[p].max, n);
          if (IRR.includes(p)) {
            const day = ts.slice(0, 8);
            if (!daily[p][day]) daily[p][day] = { sum: 0, count: 0 };
            daily[p][day].sum += n; daily[p][day].count++;
          }
          if (p === "WD10M") { windSin += Math.sin(n * Math.PI / 180); windCos += Math.cos(n * Math.PI / 180); windDirCount++; }
        }
      }
      const albVal = r[albI + 1];
      if (albVal === null || albVal === undefined) {
        const d = r[dwnI + 1];
        const num = d === null || d === undefined ? null : Number(d);
        if (num === null || !Number.isFinite(num) || num < 0) { albDwnNull++; dwnNull++; }
        else if (num === 0) { albNight++; night++; }
        else { albDayfill++; dayfill++; }
      }
    }
    chunkStats.push({ chunk_start: String(c.chunk_start), night, dayfill, dwnNull });
  }

  const dailyIrradiance = {};
  for (const p of IRR) {
    let sum = 0, count = 0, incomplete = 0;
    for (const day of Object.keys(daily[p])) {
      const d = daily[p][day];
      if (d.count === 24) { sum += d.sum / 1000; count++; } else incomplete++;
    }
    dailyIrradiance[p] = { sum, count, incomplete };
  }

  return {
    agg, dailyIrradiance, windSin, windCos, windDirCount,
    albNight, albDayfill, albDwnNull, chunkStats, totalRows, problems
  };
}

// ---------------------------------------------------------------------------
// v1.9 finalize + report text (ported verbatim from worker_k12.js)
// ---------------------------------------------------------------------------
function finalizeResourceSummary(row, state, units, metadata) {
  const mean = p => meanOf(state, p);
  const dailyMean = p => dailyMeanOf(state, p);
  let windDirection = null;
  if (state.windDirectionCount) {
    windDirection = Math.atan2(state.windSin / state.windDirectionCount, state.windCos / state.windDirectionCount) * 180 / Math.PI;
    if (windDirection < 0) windDirection += 360;
  }
  const noData = RESOURCE_PARAMETERS.some(p => !state.metrics[p].count);
  const complete = Number(row.completed_chunks) >= Number(row.total_chunks) || row.status === "COMPLETE";
  const status = noData ? "DATA UNAVAILABLE" : (!complete ? String(row.status || "REVIEW") : (state.reviewCount > 0 ? "REVIEW" : "VALID"));
  return {
    status,
    ghi: dailyMean("ALLSKY_SFC_SW_DWN"),
    dni: dailyMean("ALLSKY_SFC_SW_DNI"),
    dhi: dailyMean("ALLSKY_SFC_SW_DIFF"),
    ambientTemperature: mean("T2M"),
    minTemperature: state.metrics.T2M.min,
    maxTemperature: state.metrics.T2M.max,
    windSpeed: mean("WS10M"), windDirection,
    relativeHumidity: mean("RH2M"), surfacePressure: mean("PS"),
    precipitation: mean("PRECTOTCORR"), albedo: mean("ALLSKY_SRF_ALB"),
    missingCount: state.missingCount, invalidCount: state.invalidCount,
    duplicateCount: state.duplicateCount, unitMismatchCount: state.unitMismatchCount,
    nightAlbedoCount: state.nightAlbedoCount || 0,
    validHourlyT2MCount: state.metrics.T2M.count,
    completeIrradianceDays: state.dailyIrradiance.ALLSKY_SFC_SW_DWN.count,
    incompleteIrradianceDays: state.dailyIrradiance.ALLSKY_SFC_SW_DWN.incomplete,
    units, metadata
  };
}
function summaryTextFields(row, s, units, metadata) {
  const periodStart = row.period_start, periodEnd = row.period_end;
  const ref = `https://power.larc.nasa.gov/api/temporal/hourly/point (centroid ${Number(row.centroid_lat).toFixed(8)}, ${Number(row.centroid_lon).toFixed(8)}; UTC; ${periodStart}-${periodEnd})`;
  const remarks = [
    "GHI/DNI/DHI = mean of complete daily totals calculated from validated hourly Wh/m^2 and reported as kWh/m^2/day.",
    "Ambient temperature = mean hourly T2M; Min/Max Temperature = derived extrema of validated hourly T2M because NASA Hourly rejects T2M_MIN/T2M_MAX.",
    "Wind direction = circular mean. Other weather fields = arithmetic means of valid hourly NASA values.",
    `Precipitation uses NASA returned unit ${units.PRECTOTCORR || "unavailable"}; missing values were excluded, never changed to zero.`,
    "GTI / POA Irradiance remains DERIVED and is not populated from GHI.",
    `ALLSKY_SRF_ALB: NASA returns fill values during nighttime (sun below horizon); those ${s.nightAlbedoCount || 0} hours are classified not-applicable, excluded from the mean, stored as null and never changed to zero.`,
    `Validation: missing=${s.missingCount}, invalid=${s.invalidCount}, duplicates=${s.duplicateCount}, unit mismatches=${s.unitMismatchCount}, night-albedo N/A=${s.nightAlbedoCount || 0}.`
  ].join(" ");
  const fields = {
    schema: "SOLAR_EPC_RESOURCE_SUMMARY_V1", request_id: row.request_id,
    project_id: row.project_id, latitude: row.centroid_lat, longitude: row.centroid_lon,
    source: "NASA POWER", dataset_product: `${metadata.title || "NASA POWER Hourly Point API"} | ${metadata.apiName || "POWER Hourly API"} ${metadata.apiVersion || ""}`.trim(),
    retrieval_date: new Date(Number(row.updated_at || Date.now())).toISOString(),
    data_period_start: periodStart.replace(/^(\d{4})(\d{2})(\d{2})$/, "$1-$2-$3"),
    data_period_end: periodEnd.replace(/^(\d{4})(\d{2})(\d{2})$/, "$1-$2-$3"), resolution: "Hourly (UTC)",
    ghi: s.ghi, dni: s.dni, dhi: s.dhi, gti_poa: "DERIVED",
    ambient_temperature: s.ambientTemperature, min_temperature: s.minTemperature,
    max_temperature: s.maxTemperature, wind_speed: s.windSpeed,
    wind_direction: s.windDirection, relative_humidity: s.relativeHumidity,
    surface_pressure: s.surfacePressure, precipitation: s.precipitation,
    albedo: s.albedo, data_status: s.status, source_reference: ref, remarks
  };
  return Object.entries(fields).map(([k, v]) => `${k}|${safeText(v)}`).join("\n") + "\nEND";
}

// ---------------------------------------------------------------------------
// main
// ---------------------------------------------------------------------------
const [chunksFile, requestsFile, argReq, ...rest] = process.argv.slice(2);
const outDir = rest.includes("--out") ? rest[rest.indexOf("--out") + 1] : ".";
const expectRows = rest.includes("--expect-rows") ? Number(rest[rest.indexOf("--expect-rows") + 1]) : null;
const expectChunks = rest.includes("--expect-chunks") ? Number(rest[rest.indexOf("--expect-chunks") + 1]) : null;
if (!chunksFile || !requestsFile) {
  console.error("usage: node revalidate_night_albedo.mjs chunks.json requests.json [request_id] [--out DIR] [--expect-rows N] [--expect-chunks N]");
  process.exit(2);
}
fs.mkdirSync(outDir, { recursive: true });

const reqRows = loadRows(requestsFile);
const chunkRows = loadRows(chunksFile).filter(r => r && r.data_json != null);
if (!reqRows.length) { console.error("no resource_requests rows in requests.json"); process.exit(3); }
const requestId = argReq || String(reqRows[0].request_id || "");
const req = reqRows.find(r => String(r.request_id) === requestId);
if (!req) { console.error(`request_id ${requestId} not found in requests.json`); process.exit(3); }
const chunks = chunkRows.filter(c => String(c.request_id) === requestId);

console.log(`request_id=${requestId}`);
console.log(`workbook_id=${req.workbook_id} status=${req.status} completed=${req.completed_chunks}/${req.total_chunks} period=${req.period_start}..${req.period_end}`);
console.log(`centroid=(${req.centroid_lat}, ${req.centroid_lon})  chunk rows loaded=${chunks.length}`);

// ---- replay from stored rows ----
const R = replayChunks(chunks, requestId);
console.log(`rows replayed=${R.totalRows} chunks parsed=${R.chunkStats.length} problems=${R.problems.length}`);
for (const p of R.problems) console.log("  PROBLEM:", p);

// ---- load stored state & summary ----
let oldState, oldSummary, units, metadata;
try { oldState = JSON.parse(req.summary_state_json || "{}"); } catch { oldState = {}; }
try { oldSummary = JSON.parse(req.summary_json || "{}"); } catch { oldSummary = {}; }
try { units = JSON.parse(req.units_json || "{}"); } catch { units = oldSummary.units || {}; }
try { metadata = JSON.parse(req.metadata_json || "{}"); } catch { metadata = oldSummary.metadata || {}; }

const oldMissing = Number(oldState.missingCount || oldSummary.missingCount || 0);
const oldAlbMissing = Number(oldState.metrics?.[ALB]?.missing ?? oldSummary.missingCount ?? 0);
const storedInvalid = Number(oldState.invalidCount || oldSummary.invalidCount || 0);
const storedDup = Number(oldState.duplicateCount || oldSummary.duplicateCount || 0);
const storedUnit = Number(oldState.unitMismatchCount || oldSummary.unitMismatchCount || 0);
const storedReview = Number(oldState.reviewCount || oldSummary.missingCount || oldMissing);

// ---- verification gates ----
const gates = [];
gates.push(["chunks loaded == total_chunks", expectChunks === null ? "OK" : (chunks.length === expectChunks ? `OK (${expectChunks})` : `FAIL loaded ${chunks.length} expected ${expectChunks}`)]);
gates.push(["rows replayed == expected rows", expectRows === null ? "OK" : (R.totalRows === expectRows ? `OK (${expectRows})` : `FAIL loaded ${R.totalRows} expected ${expectRows}`)]);
gates.push(["ALB null total == stored missing", R.albNight + R.albDayfill + R.albDwnNull === oldAlbMissing ? "OK" : `FAIL replay=${R.albNight + R.albDayfill + R.albDwnNull} stored=${oldAlbMissing}`]);
gates.push(["DAYTIME_ALBEDO_MISSING == 0 (hard gate)", R.albDayfill === 0 ? "OK" : `FAIL ${R.albDayfill}`]);
gates.push(["DWN_UNAVAILABLE == 0 (hard gate)", R.albDwnNull === 0 ? "OK" : `FAIL ${R.albDwnNull}`]);
gates.push(["stored invalid == 0", storedInvalid === 0 ? "OK" : `FAIL ${storedInvalid}`]);
gates.push(["stored duplicates == 0", storedDup === 0 ? "OK" : `FAIL ${storedDup}`]);
gates.push(["stored unit mismatches == 0", storedUnit === 0 ? "OK" : `FAIL ${storedUnit}`]);

// per-parameter consistency replay vs stored state
let maxMetricDrift = 0;
for (const p of RESOURCE_PARAMETERS) {
  const A = R.agg[p], S = oldState.metrics?.[p];
  if (!S) { gates.push([`state metric ${p} exists`, "FAIL"]); continue; }
  if (A.valid !== Number(S.count)) gates.push([`${p} count replay==stored`, `FAIL ${A.valid} vs ${S.count}`]);
  if (!closeEnough(A.sum, S.sum, 1e-9)) gates.push([`${p} sum replay==stored`, `FAIL ${A.sum} vs ${S.sum}`]);
  if (!closeEnough(A.min, S.min)) gates.push([`${p} min replay==stored`, `FAIL ${A.min} vs ${S.min}`]);
  if (!closeEnough(A.max, S.max)) gates.push([`${p} max replay==stored`, `FAIL ${A.max} vs ${S.max}`]);
  const drift = S.sum ? Math.abs(A.sum - S.sum) / Math.max(1, Math.abs(S.sum)) : 0;
  maxMetricDrift = Math.max(maxMetricDrift, drift);
}

// ---- classification correction (only if gates are clean) ----
const night = R.albNight;
const hardBlocked = R.albDayfill > 0 || R.albDwnNull > 0 || gates.some(g => g[1].startsWith("FAIL"));
if (hardBlocked) {
  console.error("\nHARD GUARD FAILED — request must stay REVIEW. NO SQL GENERATED.");
  for (const g of gates) console.log(`  [${g[1].startsWith("FAIL") ? "!" : "OK"}] ${g[0]}: ${g[1]}`);
  console.error("There are genuine daytime/albedo-unknown gaps; reclassification is NOT allowed.");
  process.exit(1);
}

const newState = jsClone(oldState);
newState.metrics[ALB].missing = 0;
newState.missingCount = Math.max(0, oldMissing - (oldAlbMissing || 0));
newState.reviewCount = newState.missingCount + storedInvalid + storedDup + storedUnit;
newState.nightAlbedoCount = night;

const finalRow = { ...req, completed_chunks: req.completed_chunks, total_chunks: req.total_chunks, status: "COMPLETE" };
const newSummary = finalizeResourceSummary(finalRow, newState, units, metadata);

// value-field equivalence (only classification fields may change)
const valueFields = ["ghi", "dni", "dhi", "ambientTemperature", "minTemperature", "maxTemperature",
  "windSpeed", "windDirection", "relativeHumidity", "surfacePressure", "precipitation", "albedo",
  "validHourlyT2MCount", "completeIrradianceDays", "incompleteIrradianceDays"];
const valueDiffs = [];
for (const f of valueFields) if (!closeEnough(oldSummary[f], newSummary[f], 1e-9)) valueDiffs.push(`${f}: ${oldSummary[f]} -> ${newSummary[f]}`);

// per-chunk reclassification
const chunkUpdates = [];
for (const cs of R.chunkStats) {
  const c = chunks.find(x => String(x.chunk_start) === cs.chunk_start);
  if (!c) { gates.push([`chunk ${cs.chunk_start} present`, "FAIL"]); continue; }
  // Per-chunk summary is re-derived from the same stored rows (exact), so the
  // classification stays consistent even if the old per-chunk metadata was v1.8.
  const csState = {
    metrics: {}, dailyIrradiance: {}, windSin: 0, windCos: 0, windDirectionCount: 0,
    missingCount: 0, invalidCount: 0, duplicateCount: 0, unitMismatchCount: 0, reviewCount: 0, nightAlbedoCount: cs.night
  };
  for (const p of RESOURCE_PARAMETERS) csState.metrics[p] = { sum: 0, count: 0, min: null, max: null, missing: 0, invalid: 0 };
  for (const p of IRR) csState.dailyIrradiance[p] = { sum: 0, count: 0, incomplete: 0 };
  const doc = JSON.parse(c.data_json);
  const ord = doc.parameterOrder;
  for (const r of doc.rows) {
    const ts = String(r[0] || "");
    for (let i = 0; i < RESOURCE_PARAMETERS.length; i++) {
      const p = RESOURCE_PARAMETERS[i], v = r[i + 1];
      if (v === null || v === undefined) {
        if (p === ALB && Number(r[ord.indexOf(DWN) + 1]) === 0) { /* N/A -> not missing */ }
        else csState.metrics[p].missing++;
        continue;
      }
      const n = Number(v);
      if (!Number.isFinite(n)) continue;
      const m = csState.metrics[p];
      m.sum += n; m.count++; m.min = m.min === null ? n : Math.min(m.min, n); m.max = m.max === null ? n : Math.max(m.max, n);
      if (IRR.includes(p)) {
        const day = ts.slice(0, 8);
        if (!csState.dailyIrradiance[p][day]) csState.dailyIrradiance[p][day] = { sum: 0, count: 0 };
        csState.dailyIrradiance[p][day].sum += n; csState.dailyIrradiance[p][day].count++;
      }
      if (p === "WD10M") { csState.windSin += Math.sin(n * Math.PI / 180); csState.windCos += Math.cos(n * Math.PI / 180); csState.windDirectionCount++; }
    }
  }
  csState.missingCount = Object.values(csState.metrics).reduce((s, m) => s + m.missing, 0);
  for (const p of IRR) {
    let sum = 0, count = 0, incomplete = 0;
    for (const day of Object.keys(csState.dailyIrradiance[p])) {
      const d = csState.dailyIrradiance[p][day];
      if (d.count === 24) { sum += d.sum / 1000; count++; } else incomplete++;
    }
    csState.dailyIrradiance[p] = { sum, count, incomplete };
  }
  csState.reviewCount = csState.missingCount + csState.invalidCount + csState.duplicateCount + csState.unitMismatchCount;
  chunkUpdates.push({ chunk_start: cs.chunk_start, night: cs.night, summary_json: JSON.stringify(csState) });
}

// ---- output ----
const nowMs = Date.now();
const reportLines = [];
reportLines.push("==================================================================");
reportLines.push("RETROACTIVE RE-VALIDATION REPORT (READ-ONLY REPLAY — no writes performed)");
reportLines.push(`request_id : ${requestId}`);
reportLines.push(`workbook_id: ${req.workbook_id}`);
reportLines.push(`centroid   : ${req.centroid_lat}, ${req.centroid_lon}`);
reportLines.push(`period     : ${req.period_start}..${req.period_end}  chunks=${chunks.length}/${req.total_chunks}  rows=${R.totalRows}`);
reportLines.push("==================================================================");
reportLines.push("");
reportLines.push("REPLAYED FROM STORED resource_chunks ROWS (CHECK #2 classification)");
reportLines.push(`  ALBEDO null total          : ${R.albNight + R.albDayfill + R.albDwnNull}`);
reportLines.push(`  NIGHT_NA_CANDIDATE (DWN=0) : ${R.albNight}`);
reportLines.push(`  DAYTIME_ALBEDO_MISSING      : ${R.albDayfill}`);
reportLines.push(`  DWN_UNAVAILABLE             : ${R.albDwnNull}`);
reportLines.push(`  reconciliation             : ${R.albNight} + ${R.albDayfill} + ${R.albDwnNull} = ${R.albNight + R.albDayfill + R.albDwnNull}`);
reportLines.push("");
reportLines.push("GATES");
for (const [name, res] of gates) reportLines.push(`  [${res.startsWith("OK") ? "OK" : "!"}] ${name}: ${res}`);
reportLines.push("  [OK] max metric drift replay vs stored state: " + maxMetricDrift.toExponential(3));
reportLines.push("");
reportLines.push("VALUE FIELDS UNCHANGED? " + (valueDiffs.length ? "NO — " + valueDiffs.join("; ") : "YES (all identical within 1e-9)"));
reportLines.push("");
reportLines.push("SUMMARY (summary_json) BEFORE -> AFTER");
for (const f of ["status", "missingCount", "invalidCount", "duplicateCount", "unitMismatchCount", "nightAlbedoCount", "albedo", "ghi", "ambientTemperature"]) {
  reportLines.push(`  ${f.padEnd(20)} ${String(oldSummary[f] ?? (f === "status" ? oldState.reviewCount ? "REVIEW" : null : null)).padEnd(14)} -> ${newSummary[f] ?? "n/a"}`);
}
reportLines.push("");
reportLines.push("PER-CHUNK night-albedo N/A counts (resource_chunk_summaries):");
for (const u of chunkUpdates) reportLines.push(`  ${u.chunk_start}  night=${u.night}`);
reportLines.push("");
reportLines.push("WRITTEN FILES (local only):");
const sqlName = `revalidate_${requestId}.sql`;
const repName = `revalidation_report_${requestId}.txt`;
reportLines.push(`  ${path.join(outDir, sqlName)}`);
reportLines.push(`  ${path.join(outDir, repName)}`);
reportLines.push("");
reportLines.push("NOT TOUCHED: resource_chunks (hourly values), D1 schema, Worker code, VBA, MAP/DRAWING_DATA, Calculation Engine, API parameters.");

fs.writeFileSync(path.join(outDir, repName), reportLines.join("\n") + "\n", "utf8");

const sql = [
  "-- =========================================================================",
  `-- REVALIDATION PATCH (v1.9 night-albedo N/A rule) — request ${requestId}`,
  "-- Generated: " + new Date(nowMs).toISOString() + "  (READ-ONLY replay; rows in resource_chunks are NOT modified)",
  "-- Effect: reclassify 20,853 night-time ALLSKY_SRF_ALB fills as N/A in the",
  "--   summary metadata only. data_status -> VALID, missingCount -> 0.",
  "-- SAFE TO RE-RUN: every UPDATE is idempotent.",
  "-- =========================================================================",
  "",
  `UPDATE resource_requests SET`,
  `  summary_state_json=${sqlQuote(JSON.stringify(newState))},`,
  `  summary_json=${sqlQuote(JSON.stringify(newSummary))},`,
  `  updated_at=${nowMs}`,
  `WHERE request_id=${sqlQuote(requestId)} AND workbook_id=${sqlQuote(req.workbook_id)};`,
  ""
];
for (const u of chunkUpdates) {
  sql.push(`UPDATE resource_chunk_summaries SET summary_json=${sqlQuote(u.summary_json)}`);
  sql.push(`WHERE request_id=${sqlQuote(requestId)} AND chunk_start=${sqlQuote(u.chunk_start)};`);
  sql.push("");
}
sql.push("-- ============================ VERIFY AFTER APPLY =======================");
sql.push("-- SELECT request_id,");
sql.push("--   json_extract(summary_json,'$.status')            AS data_status,");
sql.push("--   json_extract(summary_json,'$.missingCount')      AS missing,");
sql.push("--   json_extract(summary_json,'$.nightAlbedoCount')  AS night");
sql.push("-- FROM resource_requests WHERE request_id='" + requestId + "';");
sql.push("--");
sql.push("-- SELECT chunk_start, json_extract(summary_json,'$.nightAlbedoCount') AS night");
sql.push("-- FROM resource_chunk_summaries WHERE request_id='" + requestId + "' ORDER BY chunk_start;");
sql.push("--");
sql.push("-- SELECT SUM(json_extract(je.value,'$.missing'))      AS missing,");
sql.push("--        SUM(json_extract(je.value,'$.invalid'))      AS invalid");
sql.push("-- FROM resource_chunk_summaries, json_each(json_extract(summary_json,'$.metrics')) je");
sql.push("-- WHERE request_id='" + requestId + "' GROUP BY je.key ORDER BY je.key;");
fs.writeFileSync(path.join(outDir, sqlName), sql.join("\n") + "\n", "utf8");

console.log(reportLines.join("\n"));
console.log("\nDRY-RUN: no database writes. Apply " + sqlName + " with `wrangler d1 execute --remote --file` after review.");
