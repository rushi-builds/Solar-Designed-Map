#!/usr/bin/env bash
# Reproducible proof that the v2.2 VBA centroid math is identical to the
# deployed Worker's polygonCentroidLocalProjection().
#
#   ./run_parity_check.sh
#
# Requires: node (>=16) and python3. No network access is used.
# The 5000 generated polygon cases are re-created on every run, so nothing
# large is stored in the repository.
set -euo pipefail
cd "$(dirname "$0")"

echo "==> 1/4  building 5000 deterministic polygon cases + Worker results"
node parity.mjs

echo "==> 2/4  comparing Worker centroid vs the VBA port (lat/lon/area/vertex count)"
python3 - <<'PY'
import json, sys
sys.path.insert(0, '.')
from vba_port import exact_site_location
cases = json.load(open('parity_cases.json'))
worker = json.load(open('parity_worker.json'))
assert len(cases) == len(worker), "case/result count mismatch"
ok = value_mm = state_mm = 0
for c, w in zip(cases, worker):
    mine = exact_site_location(c['coords'])
    if w['ok'] != mine['ok']:
        state_mm += 1
        print("   accept/reject mismatch:", c['id'], w.get('error'), mine.get('error'))
        continue
    if not w['ok']:
        ok += 1
        continue
    if (w['n'] != mine['vertexCount'] or w['lat8'] != mine['lat8']
            or w['lon8'] != mine['lon8'] or w['area2'] != mine['area2']):
        value_mm += 1
        print("   value mismatch:", c['id'], w, mine)
        continue
    ok += 1
print(f"   matched {ok}/{len(cases)}   value mismatches {value_mm}   accept/reject mismatches {state_mm}")
assert value_mm == 0 and state_mm == 0, "PARITY FAILED"
PY

echo "==> 3/4  comparing the cleaned boundary text + sitePolygon JSON wire format"
node - <<'EOF'
import {validateSitePolygon} from './worker_algo.mjs';
import {readFileSync,writeFileSync} from 'fs';
const cases=JSON.parse(readFileSync('parity_cases.json'));
writeFileSync('parity_worker_boundary.json',JSON.stringify(cases.map(c=>{
  try{
    const p=validateSitePolygon(c.coords.split(";").map(x=>x.split(",").map(Number)));
    return {ok:true,boundary:p.map(q=>q[0].toFixed(8)+","+q[1].toFixed(8)).join(";"),
            json:"["+p.map(q=>"["+q[0].toFixed(8)+","+q[1].toFixed(8)+"]").join(",")+"]"};
  }catch(e){return {ok:false,error:String(e.message||e)};}
})));
EOF
python3 - <<'PY'
import json, sys
sys.path.insert(0, '.')
from vba_port import parse_polygon
cases = json.load(open('parity_cases.json'))
worker = json.load(open('parity_worker_boundary.json'))
def boundary_text(lat, lon):
    return ";".join("%.8f,%.8f" % (lat[i], lon[i]) for i in range(len(lat)))
def polygon_json(text):
    parts = []
    for pair in text.split(";"):
        a, b = pair.split(",")
        parts.append("[%.8f,%.8f]" % (float(a), float(b)))
    return "[" + ",".join(parts) + "]"
mm = 0
for c, w in zip(cases, worker):
    poly, err = parse_polygon(c['coords'])
    if not w['ok']:
        if poly is not None: mm += 1
        continue
    if poly is None:
        mm += 1; continue
    lat, lon = poly
    if boundary_text(lat, lon) != w['boundary'] or polygon_json(boundary_text(lat, lon)) != w['json']:
        mm += 1
print(f"   boundary/JSON mismatches {mm}/{len(cases)}")
assert mm == 0, "BOUNDARY PARITY FAILED"
PY

echo "==> 4/4  hand-built edge cases (closed rings, duplicate vertices, triangle,"
echo "          12-vertex site, southern hemisphere, sub-metre site)"
node cases.mjs run > /tmp/edge_worker.json
python3 - <<'PY'
import json, sys
sys.path.insert(0, '.')
from cases_worker import CASES
from vba_port import exact_site_location
worker = json.load(open('/tmp/edge_worker.json'))
bad = 0
for case, w in zip(CASES, worker):
    name, coords = case
    m = exact_site_location(coords)
    if w['ok'] != m['ok'] or (w['ok'] and (w['lat8'] != m['lat8'] or w['lon8'] != m['lon8'] or w['area2'] != m['area2'])):
        bad += 1
        print("   MISMATCH", name, w, m)
    else:
        state = f"{m['lat8']}, {m['lon8']}  {m['area2']} m2  {m['vertexCount']} corners" if m['ok'] else "rejected: " + m['error']
        print(f"   {name:<18} {state}")
assert bad == 0, "EDGE CASE PARITY FAILED"
PY

rm -f parity_cases.json parity_worker.json parity_worker_boundary.json
echo
echo "PARITY OK - the v2.2 VBA extraction is numerically identical to the Worker."
