#!/usr/bin/env node
// ============================================================================
// READ-ONLY FORENSIC — Solar EPC NASA POWER resource dataset
// Parses `chunks_export.json` (output of `wrangler d1 execute --json` for the
// resource_chunks table). Computes per-parameter accounting plus the albedo
// night-vs-dayfill cross-check. It NEVER writes to D1, never changes status,
// never fills values. It only creates ONE local text file with albedo
// timestamps that had DWN > 0 (diagnostic output).
//
// Usage:
//   node forensic_readonly.mjs chunks_export.json "<REQUEST_ID>"
//
// Expected export command (read-only):
//   npx wrangler d1 execute <DB_NAME> --remote \
//     --command "SELECT request_id,chunk_start,data_json FROM resource_chunks \
//                WHERE request_id='<REQUEST_ID>' ORDER BY chunk_start" --json \
//     > chunks_export.json
// ============================================================================

import fs from "node:fs";

const RESOURCE_PARAMETERS = [
  "ALLSKY_SFC_SW_DWN", "ALLSKY_SFC_SW_DNI", "ALLSKY_SFC_SW_DIFF", "T2M",
  "WS10M", "WD10M", "RH2M", "PS", "PRECTOTCORR", "ALLSKY_SRF_ALB"
];
const ALB = "ALLSKY_SRF_ALB";
const DWN = "ALLSKY_SFC_SW_DWN";

const file = process.argv[2];
const wantRequestId = (process.argv[3] || "").trim();
if (!file || !fs.existsSync(file)) {
  console.error("Usage: node forensic_readonly.mjs chunks_export.json <REQUEST_ID>");
  process.exit(2);
}

// ---- tolerant load of wrangler --json output -----------------------------
const raw = JSON.parse(fs.readFileSync(file, "utf8"));
function collectRows(node, out) {
  if (!node || typeof node !== "object") return;
  if (Array.isArray(node)) { for (const x of node) collectRows(x, out); return; }
  if (Array.isArray(node.results)) { for (const x of node.results) { if (x && typeof x === "object") out.push(x); } return; }
  if (Array.isArray(node.result)) { for (const x of node.result) collectRows(x, out); return; }
  for (const k of Object.keys(node)) collectRows(node[k], out);
}
const rows = [];
collectRows(raw, rows);
const chunkRows = rows.filter(r => r && (r.data_json || r.data_json === null));

if (!chunkRows.length) {
  console.error("No resource_chunks rows found in export (check request filter).");
  process.exit(3);
}
const requestId = wantRequestId || String(chunkRows[0].request_id || "");
console.log(`Loaded ${chunkRows.length} chunk rows. request_id=${requestId}`);

// ---- per-parameter accounting over stored validated rows ------------------
const stats = {};
for (const p of RESOURCE_PARAMETERS) {
  stats[p] = { expected: 0, valid: 0, nullTotal: 0, invalidStored: 0, missingStored: 0, duplicatesStored: 0, pctMissing: 0 };
}
const albedoMsgs = [];
let nullAlbedoDwnZero = 0, nullAlbedoDwnPositive = 0, nullAlbedoDwnNull = 0;

function tsToUtc(ts) { // "YYYYMMDDHH" -> Date.UTC ms
  const y = Number(ts.slice(0,4)), m = Number(ts.slice(4,6)), d = Number(ts.slice(6,8)), h = Number(ts.slice(8,10));
  return Date.UTC(y, m-1, d, h, 0, 0);
}

for (const chunk of chunkRows) {
  if (requestId && String(chunk.request_id) !== requestId) continue;
  let doc;
  try { doc = JSON.parse(chunk.data_json); } catch (e) { console.warn("Bad data_json at", chunk.chunk_start, e.message); continue; }
  const order = Array.isArray(doc.parameterOrder) ? doc.parameterOrder : [];
  const idx = p => order.indexOf(p);
  const albI = idx(ALB), dwnI = idx(DWN);
  if (albI < 0 || dwnI < 0) { console.warn("parameterOrder mismatch at", chunk.chunk_start); continue; }
  const rowsArr = Array.isArray(doc.rows) ? doc.rows : [];
  let prevTs = null;
  for (const r of rowsArr) {
    const ts = String(r[0] || "");
    const t = tsToUtc(ts);
    // continuous run tracking is done below through separate arrays; here we
    // just need per-record classification for ALB.
    for (const p of RESOURCE_PARAMETERS) {
      const i = order.indexOf(p);
      if (i < 0) continue;
      stats[p].expected++;
      const v = r[i];
      if (v === null || v === undefined) stats[p].nullTotal++;
      else stats[p].valid++;
    }
    const albVal = r[albI];
    if (albVal === null || albVal === undefined) {
      const dwnVal = r[dwnI];
      const num = dwnVal === null || dwnVal === undefined ? null : Number(dwnVal);
      if (num === 0) nullAlbedoDwnZero++;
      else if (num !== null && num > 0) { nullAlbedoDwnPositive++; albedoMsgs.push({ ts, dwn: num }); }
      else nullAlbedoDwnNull++;
    }
  }
}

// ---- cross-check with stored summary counts (if exported) ----------------
// We didn't export summaries here; totals come from the SQL queries in the
// report. This script only audits what is provable from rows.

// ---- continuous runs for ALB nulls (split by class) ----------------------
const albEvents = [];
for (const chunk of chunkRows) {
  if (requestId && String(chunk.request_id) !== requestId) continue;
  let doc; try { doc = JSON.parse(chunk.data_json); } catch { continue; }
  const order = Array.isArray(doc.parameterOrder) ? doc.parameterOrder : [];
  const albI = order.indexOf(ALB), dwnI = order.indexOf(DWN);
  if (albI < 0 || dwnI < 0) continue;
  for (const r of (Array.isArray(doc.rows) ? doc.rows : [])) {
    const ts = String(r[0] || "");
    if (r[albI] === null || r[albI] === undefined) {
      const d = r[dwnI];
      const num = d === null || d === undefined ? null : Number(d);
      const cls = num === 0 ? "night" : (num !== null && num > 0 ? "dayfill" : "dwn-null");
      albEvents.push({ ts, t: tsToUtc(ts), cls });
    }
  }
}
albEvents.sort((a,b) => a.t - b.t);
function runsByClass(events) {
  const out = [];
  let cur = null;
  for (const e of events) {
    if (cur && e.cls === cur.cls && e.t - cur.end === 3600000) cur.end = e.t, cur.len++;
    else { if (cur) out.push(cur); cur = { cls: e.cls, start: e.ts, endT: e.t, end: e.t, len: 1, first: e.ts }; }
  }
  if (cur) out.push(cur);
  return out;
}
const runs = runsByClass(albEvents);
const runSummary = {};
for (const cls of ["night", "dayfill", "dwn-null"]) {
  const c = runs.filter(r => r.cls === cls);
  const long = c.filter(r => r.len >= 24);
  runSummary[cls] = { totalRuns: c.length, runsGte24h: long.length, longestRunHours: c.reduce((m,r)=>Math.max(m,r.len),0) };
}

// ---- totals ---------------------------------------------------------------
const totalExpected = Object.values(stats).reduce((s,x)=>s+x.expected,0);
const totalNull = Object.values(stats).reduce((s,x)=>s+x.nullTotal,0);
for (const p of RESOURCE_PARAMETERS) {
  const s = stats[p];
  s.pctMissing = s.expected ? (s.nullTotal / s.expected * 100).toFixed(3) + "%" : "n/a";
}

console.log("\n================ PER-PARAMETER AUDIT (from stored rows) ================");
console.log("parameter                          expected    valid   null(total)  missing%");
for (const p of RESOURCE_PARAMETERS) {
  const s = stats[p];
  console.log(`${p.padEnd(34)} ${String(s.expected).padStart(8)} ${String(s.valid).padStart(8)} ${String(s.nullTotal).padStart(12)} ${s.pctMissing.padStart(9)}`);
}
console.log(`TOTAL all params                     ${String(totalExpected).padStart(8)} ${String(totalExpected-totalNull).padStart(8)} ${String(totalNull).padStart(12)}`);
console.log(`(records per param expected = ${stats[RESOURCE_PARAMETERS[0]].expected}; report said 43,824)`);

console.log("\n================ ALLSKY_SRF_ALB null cross-check ================");
console.log(`ALB null (missing+invalid as stored) : ${stats[ALB].nullTotal}`);
console.log(`  DWN == 0        (night candidate)  : ${nullAlbedoDwnZero}`);
console.log(`  DWN >  0        (real dayfill)     : ${nullAlbedoDwnPositive}`);
console.log(`  DWN null/invalid (unavailable)     : ${nullAlbedoDwnNull}`);
console.log(`  CHECK sum                         : ${nullAlbedoDwnZero + nullAlbedoDwnPositive + nullAlbedoDwnNull} (must equal ALB null)`);

console.log("\n================ CONTINUOUS MISSING/N-A RUNS (per class) ================");
for (const cls of ["night","dayfill","dwn-null"]) {
  const r = runSummary[cls];
  console.log(`${cls.padEnd(8)} runs=${r.totalRuns}  runs>=24h=${r.runsGte24h}  longest=${r.longestRunHours}h`);
}

console.log("\n================ DAYFILL TIMESTAMPS (ALB null & DWN>0) ================");
const dayfillList = albedoMsgs.sort((a,b)=>a.ts.localeCompare(b.ts));
console.log(`count=${dayfillList.length}`);
console.log(dayfillList.slice(0, 50).map(x=>`${x.ts} DWN=${x.dwn}`).join("\n"));
if (dayfillList.length) {
  fs.writeFileSync("albedo_dayfill_timestamps.txt",
    dayfillList.map(x=>`${x.ts}\t${x.dwn}`).join("\n") + "\n",
    "utf8");
  console.log(`(all written to albedo_dayfill_timestamps.txt)`);
}

console.log("\n================ RECONCILIATION ================");
console.log(`original missing report : 20,853`);
console.log(`synthetic night test    : 21,912 (NOT evidence)`);
console.log(`null total (rows audit) : ${totalNull}`);
console.log(`-> copy these numbers + the per-parameter summary into FORENSIC_REPORT.md`);
console.log(`\nNOTE: nullTotal here is ALREADY-normalized storage (fill+invalid+night).`);
console.log(`      Compare it with SQL Query 1/2 totals to confirm they match.`);
