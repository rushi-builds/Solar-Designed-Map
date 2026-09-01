#!/usr/bin/env node
// ============================================================================
// READ-ONLY NASA POWER PROBE (external API GET only — no D1 writes, no cache)
// Mirrors worker_k12.js request exactly:
//   https://power.larc.nasa.gov/api/temporal/hourly/point
//   parameters=<CSV>, community=RE, format=JSON, time-standard=UTC, lat/lon
// Purposes:
//   A) year sweep: find which period reproduces the exact summaries the user
//      observed (T2M mean/min/max, WS10M, WD10M, RH2M, PS, PRECTOTCORR).
//   B) point comparison: check whether NASA returns identical hourly series
//      for different centroids (0.5-degree grid snapping hypothesis).
//
// Usage:
//   node forensic_nasa_probe.mjs <lat> <lon> <startYYYYMMDD> <endYYYYMMDD> [--raw out.json]
//   node forensic_nasa_probe.mjs <lat> <lon> --sweep-2020-2025 [--raw-dir dir]
// ============================================================================

import fs from "node:fs";
import path from "node:path";

const BASE = "https://power.larc.nasa.gov/api/temporal/hourly/point";
const PARAMS = ["T2M", "WS10M", "WD10M", "RH2M", "PS", "PRECTOTCORR"].join(",");
const COMMUNITY = "RE";

function argValue(name) {
  const i = process.argv.indexOf(name);
  return i >= 0 ? process.argv[i + 1] : null;
}
async function fetchPoint(lat, lon, start, end) {
  const url = new URL(BASE);
  url.searchParams.set("parameters", PARAMS);
  url.searchParams.set("community", COMMUNITY);
  url.searchParams.set("longitude", String(lon));
  url.searchParams.set("latitude", String(lat));
  url.searchParams.set("start", start);
  url.searchParams.set("end", end);
  url.searchParams.set("format", "JSON");
  url.searchParams.set("time-standard", "UTC");
  const res = await fetch(url.toString(), {
    headers: { Accept: "application/json", "User-Agent": "Solar-EPC-Forensic/1.0 (read-only)" }
  });
  if (!res.ok) throw new Error(`HTTP ${res.status} ${(await res.text()).slice(0, 300)}`);
  const json = await res.json();
  return { url: url.toString(), json };
}

function analyze(json) {
  const header = json.header || {};
  const geometry = json.geometry || {};
  const data = json.properties?.parameter || {};
  const stats = {};
  for (const p of ["T2M", "WS10M", "WD10M", "RH2M", "PS", "PRECTOTCORR"]) {
    const series = data[p] || {};
    const keys = Object.keys(series);
    let fill = 0, valid = 0, nullCount = 0;
    let sum = 0, min = null, max = null;
    let sinSum = 0, cosSum = 0, dirCount = 0;
    for (const k of keys) {
      const v = series[k];
      if (v === null || v === undefined) { nullCount++; continue; }
      const n = Number(v);
      const fillValue = Number(header.fill_value);
      if (!Number.isFinite(n) || n === fillValue) { fill++; continue; }
      valid++; sum += n;
      if (min === null || n < min) min = n;
      if (max === null || n > max) max = n;
      if (p === "WD10M") { sinSum += Math.sin(n * Math.PI / 180); cosSum += Math.cos(n * Math.PI / 180); dirCount++; }
    }
    let windDir = null;
    if (dirCount) { windDir = Math.atan2(sinSum / dirCount, cosSum / dirCount) * 180 / Math.PI; if (windDir < 0) windDir += 360; }
    stats[p] = { count: keys.length, fill, nullCount, valid, mean: valid ? sum / valid : null, min, max, windDirection: windDir };
  }
  return {
    requested: { lat: Number(geometry.coordinates?.[1]), lon: Number(geometry.coordinates?.[0]) },
    geometry, header: { apiName: header.api?.name, apiVersion: header.api?.version, title: header.title, fill_value: header.fill_value, time_standard: header.time_standard, start: header.start, end: header.end, sources: header.sources },
    messages: json.messages || [], stats, rawCounts: Object.fromEntries(Object.entries(data).map(([k, v]) => [k, Object.keys(v).length]))
  };
}

async function sweep(lat, lon, fromYear, toYear, rawDir) {
  const out = [];
  for (let y = fromYear; y <= toYear; y++) {
    const start = `${y}0101`, end = `${y}1231`;
    const { json, url } = await fetchPoint(lat, lon, start, end);
    const analysis = analyze(json);
    const rawFile = path.join(rawDir || ".", `raw_${lat}_${lon}_${start.replaceAll("-", "")}_${end.replaceAll("-", "")}.json`);
    fs.writeFileSync(rawFile, JSON.stringify({ url, requestedLat: lat, requestedLon: lon, json }, null, 0));
    console.log(`YEAR ${y} requested=(${lat},${lon}) returned=(${analysis.requested.lat},${analysis.requested.lon})`);
    for (const p of Object.keys(analysis.stats)) {
      const s = analysis.stats[p];
      console.log(`  ${p.padEnd(12)} n=${String(s.count).padStart(5)} fill=${String(s.fill).padStart(4)} valid=${String(s.valid).padStart(5)} mean=${s.mean === null ? "-" : s.mean.toFixed(6)} min=${s.min === null ? "-" : s.min} max=${s.max === null ? "-" : s.max}${p === "WD10M" ? ` dir=${s.windDirection.toFixed(6)}` : ""}`);
    }
    console.log(`  api=${analysis.header.apiName} ${analysis.header.apiVersion} fill=${analysis.header.fill_value} ts=${analysis.header.time_standard} url=${url}`);
    out.push({ year: y, requested: { lat, lon }, returned: analysis.requested, stats: analysis.stats, rawFile });
  }
  return out;
}

const first = process.argv[2], second = process.argv[3], third = process.argv[4];
const sweepYears = /^--sweep-(20\d{2})-(20\d{2})$/.exec(third || "");
if (first && second && sweepYears) {
  const from = Number(sweepYears[1]), to = Number(sweepYears[2]);
  const rawDir = argValue("--raw-dir") || ".";
  fs.mkdirSync(rawDir, { recursive: true });
  await sweep(first, second, from, to, rawDir);
} else if (first && second && third && process.argv[5]) {
  const [lat, lon, start, end] = [Number(first), Number(second), process.argv[4], process.argv[5]];
  const { json, url } = await fetchPoint(lat, lon, start, end);
  const analysis = analyze(json);
  const rawFile = argValue("--raw");
  if (rawFile) fs.writeFileSync(rawFile, JSON.stringify({ url, requestedLat: lat, requestedLon: lon, json }, null, 0));
  console.log(JSON.stringify({ url, requested: { lat, lon }, returned: analysis.requested, header: analysis.header, messages: analysis.messages, stats: analysis.stats, rawCounts: analysis.rawCounts, rawFile }, null, 2));
} else {
  console.error("Usage:");
  console.error("  node forensic_nasa_probe.mjs <lat> <lon> <startYYYYMMDD> <endYYYYMMDD> [--raw out.json]");
  console.error("  node forensic_nasa_probe.mjs <lat> <lon> --sweep-2020-2025 [--raw-dir dir]");
  process.exit(2);
}
