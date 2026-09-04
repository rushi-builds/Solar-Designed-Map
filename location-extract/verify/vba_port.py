# 1:1 port of the VBA functions I am about to write, so the numbers can be
# compared against the real Worker JS before the .bas is shipped.
import math, json, sys
from cases_worker import CASES  # generated below

EARTH_R = 6378137.0
MAX_VERTICES = 1000

def vba_format8(x):
    # Format$(Value, "0.00000000") -> round-half-away-from-zero, 8 decimals
    from decimal import Decimal, ROUND_HALF_UP
    return float(Decimal(repr(x)).quantize(Decimal('1E-8'), rounding=ROUND_HALF_UP))

def strict_number(t):
    import re
    return re.fullmatch(r"-?[0-9]+(?:\.[0-9]+)?", t.strip()) is not None

def parse_polygon(coordinate_text):
    """ResourceParseSitePolygon"""
    err = ""
    raw = coordinate_text.strip()
    if raw == "":
        return None, "Coordinates cell is blank."
    pairs = raw.split(";")
    lat_list, lon_list, n = [], [], 0
    for p in pairs:
        p = p.strip()
        if p == "":
            continue
        xy = p.split(",")
        if len(xy) - 1 != 1:
            return None, "Each coordinate pair must be written as lat,lon"
        lt, ln = xy[0].strip(), xy[1].strip()
        if not strict_number(lt) or not strict_number(ln):
            return None, "Coordinate text is not a plain decimal number"
        lat, lon = float(lt), float(ln)
        if lat < -90 or lat > 90 or lon < -180 or lon > 180:
            return None, "Coordinate is outside the valid WGS84 range"
        lat_list.append(vba_format8(lat)); lon_list.append(vba_format8(lon)); n += 1
    if n > MAX_VERTICES:
        return None, "SITE polygon exceeds 1000 vertices"
    # drop consecutive duplicates (8-dp), then a repeated closing vertex
    k = 0
    olat, olon = [], []
    for i in range(n):
        if k > 0 and olat[k-1] == lat_list[i] and olon[k-1] == lon_list[i]:
            continue
        olat.append(lat_list[i]); olon.append(lon_list[i]); k += 1
    if k > 1 and olat[0] == olat[k-1] and olon[0] == olon[k-1]:
        k -= 1
    if k < 3:
        return None, "SITE polygon needs at least three distinct vertices"
    if len({(olat[i], olon[i]) for i in range(k)}) < 3:
        return None, "SITE polygon needs at least three distinct vertices"
    return (olat[:k], olon[:k]), err

def centroid_local_projection(lat, lon, n):
    """ResourceCentroidLocalProjection - identical algebra to worker_k12.js"""
    sum_lat = sum(lat[i] for i in range(n))
    sum_lon = sum(lon[i] for i in range(n))
    mean_lat = sum_lat / n
    mean_lon = sum_lon / n
    cos_lat = math.cos(mean_lat * math.pi / 180.0)
    if abs(cos_lat) < 1e-12:
        return None, "Polygon is too close to a pole for this centroid projection."
    xy_x, xy_y = [0.0]*n, [0.0]*n
    for i in range(n):
        xy_x[i] = (lon[i] - mean_lon) * math.pi / 180.0 * EARTH_R * cos_lat
        xy_y[i] = (lat[i] - mean_lat) * math.pi / 180.0 * EARTH_R
    twice_area = cx6a = cy6a = 0.0
    for i in range(n):
        j = (i + 1) % n
        cross = xy_x[i] * xy_y[j] - xy_x[j] * xy_y[i]
        twice_area += cross
        cx6a += (xy_x[i] + xy_x[j]) * cross
        cy6a += (xy_y[i] + xy_y[j]) * cross
    if abs(twice_area) < 1e-9:
        return None, "SITE polygon has zero area (all vertices are collinear)."
    cx = cx6a / (3.0 * twice_area)
    cy = cy6a / (3.0 * twice_area)
    return {"latitude": mean_lat + (cy / EARTH_R) * 180.0 / math.pi,
            "longitude": mean_lon + (cx / (EARTH_R * cos_lat)) * 180.0 / math.pi,
            "areaM2": abs(twice_area) / 2.0, "vertexCount": n}, ""

def exact_site_location(coords):
    poly, err = parse_polygon(coords)
    if poly is None:
        return {"ok": False, "error": err}
    lat, lon = poly
    c, err = centroid_local_projection(lat, lon, len(lat))
    if c is None:
        return {"ok": False, "error": err}
    if c["areaM2"] < 0.01:
        return {"ok": False, "error": "SITE polygon has zero or negligible area."}
    c["ok"] = True
    c["lat8"] = "%.8f" % c["latitude"]
    c["lon8"] = "%.8f" % c["longitude"]
    c["area2"] = "%.2f" % c["areaM2"]
    c["vertices"] = ";".join("%.8f,%.8f" % (lat[i], lon[i]) for i in range(len(lat)))
    return c

if __name__ == "__main__":
    print(json.dumps([dict(name=n, **exact_site_location(c)) for n, c in CASES], indent=1))
