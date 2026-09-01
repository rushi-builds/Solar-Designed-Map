#!/usr/bin/env node
// READ-ONLY payload comparison: verifies byte-level equality of NASA hourly
// series between two saved POWER JSON responses (geom/header excluded from
// value comparison but printed for provenance).
//   node forensic_payload_compare.mjs a.json b.json [more...]
import fs from "node:fs";
import crypto from "node:crypto";

const files = process.argv.slice(2);
if (files.length < 2) { console.error("usage: forensic_payload_compare.mjs a.json b.json ..."); process.exit(2); }
function load(f) { return JSON.parse(fs.readFileSync(f, "utf8")); }
function series(payload, param) { return payload.properties.parameter[param] || {}; }
function hashOf(obj) { return crypto.createHash("sha256").update(JSON.stringify(obj)).digest("hex").slice(0, 16); }

const docs = files.map(load);
const params = Object.keys(docs[0].properties.parameter);
for (let i = 1; i < docs.length; i++) {
  let allEqual = true, totalValues = 0, diffValues = 0;
  const firstDiffs = [];
  for (const p of params) {
    const a = series(docs[0], p), b = series(docs[i], p);
    const ka = Object.keys(a).sort(), kb = Object.keys(b).sort();
    const tsEqual = JSON.stringify(ka) === JSON.stringify(kb);
    let eq = tsEqual;
    if (tsEqual) {
      for (const t of ka) {
        totalValues++;
        if (Number(a[t]) !== Number(b[t])) { eq = false; diffValues++; if (firstDiffs.length < 5) firstDiffs.push({ p, t, a: a[t], b: b[t] }); }
      }
    }
    if (!eq) allEqual = false;
    console.log(`${p.padEnd(12)} keys=${String(ka.length).padStart(3)} timestampsEqual=${tsEqual} valuesEqual=${eq}`);
  }
  console.log(`FILE ${files[0]} vs ${files[i]}`);
  console.log(`  geometry A: ${JSON.stringify(docs[0].geometry.coordinates)}`);
  console.log(`  geometry B: ${JSON.stringify(docs[i].geometry.coordinates)}`);
  console.log(`  header A: ${docs[0].header.api.name} ${docs[0].header.api.version} | ${JSON.stringify(docs[0].header.sources)} | ${docs[0].header.start}..${docs[0].header.end}`);
  console.log(`  header B: ${docs[i].header.api.name} ${docs[i].header.api.version} | ${JSON.stringify(docs[i].header.sources)} | ${docs[i].header.start}..${docs[i].header.end}`);
  const seriesHashA = hashOf(docs[0].properties.parameter), seriesHashB = hashOf(docs[i].properties.parameter);
  console.log(`  series sha256(16) A=${seriesHashA} B=${seriesHashB}  -> ${seriesHashA === seriesHashB ? "SERIES IDENTICAL" : "SERIES DIFFER"}`);
  console.log(`  values compared=${totalValues} differed=${diffValues}  -> ${allEqual ? "ALL VALUES IDENTICAL" : "DIFFERENT (see first diffs)"}`);
  if (firstDiffs.length) console.log("  first diffs:", JSON.stringify(firstDiffs, null, 0));
}
