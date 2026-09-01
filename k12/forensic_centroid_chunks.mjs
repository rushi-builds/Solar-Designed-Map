#!/usr/bin/env node
// ============================================================================
// READ-ONLY centroid forensics: compares two `resource_chunks` exports at the
// RAW hourly value level (item 9 of the user's checklist), plus optional
// `resource_requests` metadata export for cache/polygon/isolation checks.
//
// Usage:
//   node forensic_centroid_chunks.mjs chunks_req1.json chunks_reqC.json [requests_meta.json]
//
// Export commands: see k12/FORENSIC_QUERIES_READONLY.sql section 2.7.
// ============================================================================

import fs from "node:fs";
import crypto from "node:crypto";

const [fileA, fileB, metaFile] = process.argv.slice(2);
if (!fileA || !fileB) { console.error("usage: forensic_centroid_chunks.mjs chunksA.json chunksB.json [requests_meta.json]"); process.exit(2); }

function loadRows(file) {
  const raw = JSON.parse(fs.readFileSync(file, "utf8"));
  const out = [];
  (function walk(n) {
    if (!n || typeof n !== "object") return;
    if (Array.isArray(n)) { for (const x of n) walk(x); return; }
    if (Array.isArray(n.results)) { for (const x of n.results) { if (x && typeof x === "object") out.push(x); } return; }
    if (Array.isArray(n.result)) { for (const x of n.result) walk(x); return; }
    for (const k of Object.keys(n)) walk(n[k]);
  })(raw);
  return out;
}

function analyzeChunks(rows) {
  const byReq = {};
  for (const r of rows) {
    const id = String(r.request_id || "");
    let doc;
    try { doc = JSON.parse(r.data_json); } catch { continue; }
    const order = Array.isArray(doc.parameterOrder) ? doc.parameterOrder : [];
    const params = {};
    let rowsCount = 0;
    for (const row of (Array.isArray(doc.rows) ? doc.rows : [])) {
      rowsCount++;
      const ts = String(row[0] || "");
      for (let i = 1; i < row.length; i++) {
        const p = order[i];
        if (!p) continue;
        if (!params[p]) params[p] = {};
        params[p][ts] = row[i];
      }
    }
    const key = crypto.createHash("sha256").update(JSON.stringify(params)).digest("hex").slice(0, 16);
    byReq[id] = byReq[id] || { chunks: 0, rows: 0, params: {}, checksums: new Set(), matchCount: 0, nonNullCount: 0 };
    const acc = byReq[id];
    acc.chunks++; acc.rows += rowsCount;
    acc.checksums.add(String(r.checksum_sha256 || ""));
    for (const [p, series] of Object.entries(params)) {
      if (!acc.params[p]) acc.params[p] = {};
      for (const [ts, v] of Object.entries(series)) {
        acc.params[p][ts] = v;
        if (p === "T2M" || p === "WS10M" || p === "RH2M" || p === "PS" || p === "PRECTOTCORR") {
          acc.matchCount++;
          if (v !== null && v !== undefined) acc.nonNullCount++;
        }
      }
    }
  }
  for (const id of Object.keys(byReq)) byReq[id].checksums = [...byReq[id].checksums];
  return byReq;
}

const a = analyzeChunks(loadRows(fileA));
const b = analyzeChunks(loadRows(fileB));
const ids = [...new Set([...Object.keys(a), ...Object.keys(b)])].sort();

console.log("================ REQUEST INVENTORY ================");
for (const id of ids) {
  const x = a[id] || { chunks: 0, rows: 0, checksums: [] };
  const y = b[id] || { chunks: 0, rows: 0, checksums: [] };
  console.log(`request_id=${id}`);
  console.log(`  chunks: A=${x.chunks} B=${y.chunks}   rows: A=${x.rows} B=${y.rows}`);
  console.log(`  distinct checksums: A=${x.checksums.length} B=${y.checksums.length}`);
}

console.log("\n================ RAW VALUE COMPARISON (T2M, WS10M, RH2M, PS, PRECTOTCORR) ================");
const compare = (idA, idB) => {
  if (!a[idA] || !b[idB]) { console.log(`  missing side: ${!a[idA] ? idA : idB}`); return; }
  for (const p of ["T2M", "WS10M", "RH2M", "PS", "PRECTOTCORR"]) {
    const sa = a[idA].params[p] || {}, sb = b[idB].params[p] || {};
    const ts = [...new Set([...Object.keys(sa), ...Object.keys(sb)])].sort();
    let same = 0, diff = 0, firstDiff = null;
    for (const t of ts) {
      const va = sa[t], vb = sb[t];
      const eq = (va === vb) || (va === null && vb === null) || (va === undefined && vb === undefined);
      if (eq) same++; else { diff++; if (!firstDiff) firstDiff = { t, a: va, b: vb }; }
    }
    console.log(`${p.padEnd(12)} timestamps=${String(ts.length).padStart(5)} equal=${String(same).padStart(5)} different=${String(diff).padStart(4)}${firstDiff ? ` firstDiff=${JSON.stringify(firstDiff)}` : ""}`);
  }
};
compare(ids[0], ids[ids.length - 1]);

console.log("\n================ PER-REQUEST TOTALS ================");
for (const id of ids) {
  const x = a[id] || b[id];
  if (!x) continue;
  console.log(`request_id=${id}`);
  for (const p of ["T2M", "WS10M", "RH2M", "PS", "PRECTOTCORR"]) {
    const s = x.params[p] || {};
    const vals = Object.values(s).filter(v => v !== null && v !== undefined).map(Number);
    const mean = vals.length ? vals.reduce((m, v) => m + v, 0) / vals.length : null;
    const min = vals.length ? Math.min(...vals) : null, max = vals.length ? Math.max(...vals) : null;
    console.log(`  ${p.padEnd(12)} values=${String(vals.length).padStart(5)} mean=${mean === null ? "-" : mean.toFixed(6)} min=${min} max=${max}`);
  }
}

if (metaFile && fs.existsSync(metaFile)) {
  try {
    const meta = JSON.parse(fs.readFileSync(metaFile, "utf8"));
    const rows = [];
    (function walk(n) {
      if (!n || typeof n !== "object") return;
      if (Array.isArray(n)) { for (const x of n) walk(x); return; }
      if (Array.isArray(n.results)) { for (const x of n.results) rows.push(x); return; }
      if (Array.isArray(n.result)) { for (const x of n.result) rows.push(x); return; }
      for (const k of Object.keys(n)) walk(n[k]);
    })(meta);
    console.log("\n================ REQUESTS METADATA (workbook-scoped rows) ================");
    for (const r of rows) {
      console.log(`${r.request_id} | status=${r.status} | centroid=(${r.centroid_lat},${r.centroid_lon}) | period=${r.period_start}..${r.period_end} | chunks=${r.completed_chunks}/${r.total_chunks}`);
      console.log(`  cache_key=${String(r.cache_key || "").slice(0, 24)}... polygon_hash=${String(r.polygon_hash || "").slice(0, 16)}... workbook=${r.workbook_id}`);
    }
  } catch (e) { console.log("meta parse failed:", e.message); }
}
console.log("\nREAD-ONLY: no writes performed.");
