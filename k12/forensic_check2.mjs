#!/usr/bin/env node
// ============================================================================
// CHECK #2 — READ-ONLY actual-row forensics for ALLSKY_SRF_ALB fills
// For every stored ALLSKY_SRF_ALB null (fill) record, inspect the matching
// ALLSKY_SFC_SW_DWN value at the same timestamp and classify into exactly one:
//   1. DWN == 0                                   -> NIGHT_NA_CANDIDATE
//   2. DWN >  0                                   -> DAYTIME_ALBEDO_MISSING
//   3. DWN null / missing / not a positive number -> DWN_UNAVAILABLE
// Reads ONLY local JSON produced by a read-only wrangler SELECT. Never writes
// to D1. Only diagnostic text files are written locally.
//
// Usage:
//   1) export (read-only):
//      npx wrangler d1 execute <DB_NAME> --remote \
//        --command "SELECT request_id,chunk_start,parameter_order_json,data_json \
//                   FROM resource_chunks \
//                   WHERE request_id='1dce0c1c-0306-4765-97f8-567a49bed34b' \
//                   ORDER BY chunk_start" --json > check2_chunks.json
//   2) run:
//      node forensic_check2.mjs check2_chunks.json 1dce0c1c-0306-4765-97f8-567a49bed34b
// ============================================================================

import fs from "node:fs";

const ALB = "ALLSKY_SRF_ALB";
const DWN = "ALLSKY_SFC_SW_DWN";
const EXPECTED_ALB_MISSING_FROM_STEP1 = 20853; // user's verified Step-1 result

const [file, wantRequestId] = [process.argv[2], (process.argv[3] || "").trim()];
if (!file || !fs.existsSync(file)) {
  console.error("usage: node forensic_check2.mjs <chunks_export.json> <request_id>");
  process.exit(2);
}

const raw = JSON.parse(fs.readFileSync(file, "utf8"));
function collect(node, out) {
  if (!node || typeof node !== "object") return;
  if (Array.isArray(node)) { for (const x of node) collect(x, out); return; }
  if (Array.isArray(node.results)) { for (const x of node.results) { if (x && typeof x === "object") out.push(x); } return; }
  if (Array.isArray(node.result)) { for (const x of node.result) collect(x, out); return; }
  for (const k of Object.keys(node)) collect(node[k], out);
}
const rows = [];
collect(raw, rows);
const chunks = rows.filter(r => r && (r.data_json != null));
console.log(`loaded ${chunks.length} resource_chunks rows (request filter applied: ${wantRequestId || "(from payload)"})`);

let albIdx = -1, dwnIdx = -1, docCount = 0;
const records = [];           // {ts, dwn, alb, cls, chunk}
let badOrder = 0;

for (const c of chunks) {
  if (wantRequestId && String(c.request_id || "") !== wantRequestId) continue;
  let doc;
  try { doc = JSON.parse(c.data_json); } catch { continue; }
  const order = Array.isArray(doc.parameterOrder) ? doc.parameterOrder : [];
  if (!order.length) { badOrder++; continue; }
  albIdx = order.indexOf(ALB);
  dwnIdx = order.indexOf(DWN);
  if (albIdx < 0 || dwnIdx < 0) { badOrder++; continue; }
  docCount++;
  for (const r of (Array.isArray(doc.rows) ? doc.rows : [])) {
    const ts = String(r[0] || "");
    const alb = r[albIdx];
    if (alb !== null && alb !== undefined) continue;     // valid albedo: not part of this check
    const d = r[dwnIdx];
    const num = (d === null || d === undefined) ? null : Number(d);
    let cls;
    if (num === null || !Number.isFinite(num) || num < 0) cls = "DWN_UNAVAILABLE";
    else if (num === 0) cls = "NIGHT_NA_CANDIDATE";
    else cls = "DAYTIME_ALBEDO_MISSING";
    records.push({ ts, dwn: num, cls, chunk: String(c.chunk_start || "") });
  }
}

console.log(`parsed docs=${docCount} parameterOrderBad=${badOrder} albIdx=${albIdx} dwnIdx=${dwnIdx}`);
const total = records.length;
const byClass = {};
for (const r of records) byClass[r.cls] = (byClass[r.cls] || 0) + 1;
const n = byClass.NIGHT_NA_CANDIDATE || 0, day = byClass.DAYTIME_ALBEDO_MISSING || 0, unavail = byClass.DWN_UNAVAILABLE || 0;

console.log("\n================ CHECK #2 RESULT ================");
console.log(`total ALB missing/fill            : ${total}`);
console.log(`  NIGHT_NA_CANDIDATE              : ${n}`);
console.log(`  DAYTIME_ALBEDO_MISSING          : ${day}`);
console.log(`  DWN_UNAVAILABLE                 : ${unavail}`);
console.log(`  reconciliation sum              : ${n + day + unavail}  ${n + day + unavail === total ? "OK" : "!! MISMATCH !!"}`);
console.log(`  vs Step-1 expected 20,853       : ${total === EXPECTED_ALB_MISSING_FROM_STEP1 ? "MATCH" : `DIFFERENCE ${total - EXPECTED_ALB_MISSING_FROM_STEP1}`}`);

const sorted = [...records].sort((a, b) => a.ts.localeCompare(b.ts));
for (const cls of ["NIGHT_NA_CANDIDATE", "DAYTIME_ALBEDO_MISSING", "DWN_UNAVAILABLE"]) {
  const arr = sorted.filter(r => r.cls === cls);
  if (!arr.length) { console.log(`\n[${cls}] none`); continue; }
  // continuous runs (UTC hour steps)
  const runs = [];
  let cur = null;
  for (const r of arr) {
    const t = Date.UTC(+r.ts.slice(0,4), +r.ts.slice(4,6)-1, +r.ts.slice(6,8), +r.ts.slice(8,10));
    if (cur && t - cur.lastT === 3600000) { cur.len++; cur.lastT = t; cur.end = r.ts; }
    else { if (cur) runs.push(cur); cur = { start: r.ts, end: r.ts, len: 1, lastT: t }; }
  }
  if (cur) runs.push(cur);
  const longest = runs.reduce((m, r) => r.len > m.len ? r : m, runs[0]);
  console.log(`\n[${cls}] count=${arr.length}`);
  console.log(`  first timestamp : ${arr[0].ts} (DWN=${arr[0].dwn})`);
  console.log(`  last  timestamp : ${arr[arr.length-1].ts} (DWN=${arr[arr.length-1].dwn})`);
  console.log(`  runs total=${runs.length}  runs>=24h=${runs.filter(r=>r.len>=24).length}`);
  console.log(`  longest run     : ${longest.len} h  ${longest.start} .. ${longest.end}`);
}

const dayfill = sorted.filter(r => r.cls === "DAYTIME_ALBEDO_MISSING");
if (dayfill.length) {
  fs.writeFileSync("check2_daytime_albedo_missing.txt",
    "timestamp\tALLSKY_SFC_SW_DWN\tchunk\n" + dayfill.map(r => `${r.ts}\t${r.dwn}\t${r.chunk}`).join("\n") + "\n", "utf8");
  fs.writeFileSync("check2_runs.txt", JSON.stringify({ byClass, runsByClass: Object.fromEntries(
    ["NIGHT_NA_CANDIDATE","DAYTIME_ALBEDO_MISSING","DWN_UNAVAILABLE"].map(cls => {
      const arr = sorted.filter(r => r.cls === cls);
      const runs = [];
      let cur = null;
      for (const r of arr) {
        const t = Date.UTC(+r.ts.slice(0,4), +r.ts.slice(4,6)-1, +r.ts.slice(6,8), +r.ts.slice(8,10));
        if (cur && t - cur.lastT === 3600000) { cur.len++; cur.lastT = t; cur.end = r.ts; }
        else { if (cur) runs.push(cur); cur = { start: r.ts, end: r.ts, len: 1, lastT: t }; }
      }
      if (cur) runs.push(cur);
      return [cls, runs];
    })
  )}, null, 2));
  console.log("\ndiagnostic files written: check2_daytime_albedo_missing.txt, check2_runs.txt");
}
console.log("\nREAD-ONLY: no D1 writes performed. STOP until user approval.");
