# Solar EPC — Exact Drawn-Site Location (v2.2)

> **"Is file se hum ek location ko open karte hain — toh jaha humne draw kiya,
> uski EXACT location extract ho sakti hai kya?"**
>
> **Haan.** Aur ab wo Excel ke andar hi mil jaati hai — bina Worker, bina NASA,
> bina internet, bina cost.

---

## 0. Kaunsi file import karni hai? (pehle ye padho)

| Tumhe kya chahiye | Ye import karo | Kya hota hai |
|---|---|---|
| **Sab automatic + purana sab kaam karta rahe** (recommended) | **`modSolarEPCResource.bas`** — v2.4 COMPLETE+AUTO build | Purane module ka **poora replacement**: NASA pipeline, scheduler, saare macros waise hi + watcher jo har ~5s DRAWING_DATA check karke RESOURCE_DB ke **blank Latitude/Longitude me exact centroid** aur **blank Location me "Exact Area, Taluka, District"** apne aap likhta hai. Koi Alt+F8 nahi, koi popup nahi. |
| Manual toolkit (FillDrawnCentroids, `SolarEPC_Exact*` functions, Show macro) | `modSolarEPCResource_v2.2_complete.bas` | Pehle purana `modSolarEPCResource` module **remove** karo, phir import karo. |

Automatic build deliberately safe hai: **sirf BLANK cells bharta hai** (manual entry /
formula kabhi overwrite nahi), NASA numbers / Data Status / cache identity ko chhoota
hi nahi, aur tumhara existing `modSolarEPCResource` module jaisa hai waisa rahega
(module name alag hai: `modSolarEPCDrawnLocation`).

---

## 1. Exact location aati kaha se hai (poora flow)

```
 map_main.html / public/index.html          Excel VBA                        Cloud
 ─────────────────────────────────          ─────────                        ─────
 DRAW SITE → corners click → SAVE
        │
        │  buildPayload()
        │  siteCoordinates: [{point,latitude,longitude}]  ← 8 decimals
        │  + selectedLocation + area + obstacles
        ▼
 window.chrome.webview.postMessage
 { type:"SOLAR_EPC_SITE_BOUNDARY_SAVED" }
        │
        └────────────────────────────►  WebMessageReceived
                                             │
                                             ▼
                                  DRAWING_DATA.autoLWHTbl
                                  "Coordinates" = "lat,lon;lat,lon;…"  ← drawn boundary
                                  "Reference id" = "SITE-MAP-<messageId>"
                                             │
                     ┌───────────────────────┴───────────────────────┐
                     │  v2.2 (NEW, offline)                          │  v1.8+ (existing)
                     ▼                                               ▼
          ResourceExactSiteData(ProjectID)              SolarEPC_ResourceQueueImportedSite
          → exact area-weighted CENTROID                → POST /v1/resource/start
          → area m², corner count                       → Worker computes the SAME centroid
          → Google Maps link                            → NASA POWER hourly → RESOURCE_DB
```

Do cheezein "exact location" hain, aur dono ab Excel me available hain:

| # | Kya | Kaha stored | v2.2 se pehle | v2.2 me |
|---|---|---|---|---|
| a | **Drawn boundary** — har corner jo tumne click kiya (8 decimals) | `DRAWING_DATA.autoLWHTbl` → `Coordinates` | cell me pada tha, padhne ka API nahi tha | `SolarEPC_ExactSiteBoundary` |
| b | **Exact site centroid** — us polygon ka area-weighted centre. **Yahi authoritative coordinate hai** (NASA ko yahi jata hai, `RESOURCE_DB` Latitude/Longitude me yahi likha jata hai) | sirf Worker ke paas tha | Worker + internet zaroori tha | `SolarEPC_ExactSiteLocation` — **offline** |

---

## 1.5 v2.3 — AUTOMATIC fill (bina Alt+F8, bina popup)

```
 workbook open
      │  Auto_Open  (legacy auto-start, koi paste nahi karna)
      ▼
 watcher ON  +  turant ek sweep (pehle se drawn sites bhi bhar jaati hain)
      │
      │  har 5 second (AUTO_TICK_SECONDS)
      ▼
 DRAWING_DATA.autoLWHTbl scan
      │  nayi SITE row (Reference id = SITE-MAP-...) ya dobara-draw hui row
      ▼
 exact centroid (offline, Worker-identical math)
      ▼
 RESOURCE_DB me us Project ID ki row:
      • BLANK Latitude (°) / Longitude (°)  ← exact centroid (8 decimals)
      • BLANK Location                     ← "Exact Area, Taluka, District"
                                            (VILLAGE_DB → Google → Nominatim → BigDataCloud)
      ▼
 status bar: "Solar EPC: exact drawn location auto-filled -> PRJ-001 | lat lon location(...)"
 _CLOUD_CFG!A18/B18: last auto-fill record
```

### Safety rules (kabhi violate nahi hote)
- **Blank-only**: bhara hua cell, manual entry, formula cell — kabhi overwrite nahi.
- **Row match**: same Project ID jiske lat/lon blank hain, ya lat/lon already isi
  centroid ke barabar → wahi row; warna kuch ticks retry (`NOROW:n`), phir chhod deta hai.
- **NASA untouched**: GHI/DNI/…/Data Status/cache_key/polygon_hash — is module me
  unka code hai hi nahi.
- **Label retry**: internet/VILLAGE_DB na ho to lat/lon phir bhi bharte hain; label
  6 ticks tak retry hota hai, phir ruk jaata hai — galat value kabhi nahi likhi jaati.
- **Redraw**: Coordinates badle to key badalti hai → naya centroid; par blank-only
  rule ki wajah se purani bhari hui value replace NAHI hoti (manual safety).

### Control macros (zaroorat padhe to)
| Macro | Kaam |
|---|---|
| `SolarEPC_DrawnLocationAutoStart` | watcher chalu + turant sweep |
| `SolarEPC_DrawnLocationAutoStop` | watcher band (OnTime cancel) |
| `SolarEPC_DrawnLocationAutoNow` | bina wait kiye turant ek sweep |
| `SolarEPC_ResourceShowDrawnLocation` | manual dekhna ho (verification) |
| `SolarEPC_DrawnLocationDebug` | **kuch nahi bhar raha? ye chalao** — poori chain ka read-only report (watcher ON?, SITE row?, centroid?, resource_db row?, label tier?, last fill) + clipboard copy |

### v2.5 additions
- **Auto row-create**: RESOURCE_DB me us Project ID ki row hai hi nahi, to watcher
  ~10s baad row apne aap bana deta hai (Project ID + centroid + label) — draw kiya
  hai to location dikhegi hi dikhegi.
- **Status bar reasons**: NOROW / LOCPEND jaisi states ab chhupti nahi — status bar
  pe likha aata hai ki kya ho raha hai aur kyu.
- **`SolarEPC_DrawnLocationDebug`**: support chahiye ho to iska screenshot bhej do.

Workbook band hone pe `Auto_Close` scheduler saaf kar deta hai.

### Agar "ambiguous name: Auto_Open" aaye (bahut rare)
Is workbook me kahin aur Auto_Open/Auto_Close defined honge. Is module ke do
chhote stub hata do aur ThisWorkbook me ye rakh do:
```vb
Private Sub Workbook_Open()
    SolarEPC_DrawnLocationAutoStart
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    SolarEPC_DrawnLocationAutoStop
End Sub
```

---

## 2. v2.2 me kya add hua

`modSolarEPCResource.bas` (v2.1 → v2.2). **Sirf addition hai — resource flow ka ek bhi line change nahi hua.**

### Naye public macros (button / Alt+F8 se chalao)

| Macro | Kya karta hai |
|---|---|
| `SolarEPC_ResourceShowDrawnLocation` | Ek click me last drawn SITE ki exact location dikhata hai: Project ID, Reference, corner count, area (m²), **exact centroid lat/lon (8 decimals)**, descriptive location label. Phir poochta hai — **Yes** = Google Maps us exact centroid pe khul jaye, **No** = coordinates clipboard pe copy, **Cancel** = band. |
| `SolarEPC_ResourceFillDrawnCentroids` | `RESOURCE_DB` ke **blank** `Latitude (°)` / `Longitude (°)` cells ko drawn-site centroid se bhar deta hai (Project ID se match), phir existing resolver se blank `Location` label bhar deta hai. Confirm dialog ke bina kuch nahi karta. |

### Naye public functions (kisi bhi module / Immediate window se)

```vb
' 1. Exact centroid — 100% offline, no geocoding, no network
Dim la As Double, lo As Double
If SolarEPC_ExactSiteLocation("PRJ-001", la, lo) Then
    Debug.Print la, lo          ' 17.99088747   74.43575012
End If

' 2. Exact drawn area (m²) — offline
Debug.Print SolarEPC_ExactSiteArea("PRJ-001")          ' 11786.13

' 3. Ek printable line (label + coords + area + corners) — label ke liye network lag sakta hai
Debug.Print SolarEPC_ExactSiteLocationText("PRJ-001")
' Sonwadi Bk., Phaltan, Satara | 17.99088747, 74.43575012 | 11,786.13 m2 | 4 corners

' 4. Raw drawn boundary + Worker wire JSON — offline
Dim Boundary As String, WireJSON As String, Corners As Long
If SolarEPC_ExactSiteBoundary("PRJ-001", Boundary, WireJSON, Corners) Then
    Debug.Print Corners, Boundary
    Debug.Print WireJSON        ' [[17.99038747,74.43525012],[…]]  ← exactly what /v1/resource/start receives
End If
```

Sheet cell me bhi (formula se nahi — VBA se, kyunki ye table padhta hai):
`=SolarEPC_ExactSiteLocationText(A2)` type helper chahiye to ek chhoti UDF wrapper bana lena;
module functions already `Public` hain.

### Naye internal helpers

`ResourceExactSiteData`, `ResourceLatestSiteData`, `ResourceSiteTable`,
`ResourceFindSiteRowByProject`, `ResourceRowIsImportedSite`, `ResourceRowProjectID`,
`ResourceRowReferenceText`, `ResourceRowCoordinateText`, `ResourceRowCellText`,
`ResourceParseSitePolygon`, `ResourceBoundaryText`, `ResourceDistinctVertexCount`,
`ResourceCentroidLocalProjection`, `ResourceCopyToClipboard`

Naye constants: `SITE_MAX_VERTICES = 1000`, `SITE_MIN_AREA_M2 = 0.01`,
`EARTH_RADIUS_M = 6378137` — **teeno Worker ke `validateSitePolygon()` se match karte hain.**
Naya diagnostic: `_CLOUD_CFG!A18/B18` = last extracted drawn-site location.

---

## 3. Sabse important baat — numbers Worker se identical hain

`ResourceCentroidLocalProjection` Worker ke `polygonCentroidLocalProjection()`
(`k12/worker_k12.js` lines 578–601) ka **statement-by-statement port** hai:

- same Earth radius `6378137` m (WGS84 semi-major axis)
- same vertex rounding — har point **8 decimals** pe round, maths se pehle
- same duplicate rules — consecutive duplicates drop, repeated closing vertex drop
- same local equirectangular projection (mean vertex ke around) → area-weighted centroid → WGS84 wapas
- same guards — `|cos(lat)| < 1e-12` (pole), `|2A| < 1e-9` (zero area), `area < 0.01 m²` (negligible)
- same limits — 3…1000 vertices, ≥3 distinct vertices, WGS84 range check

Isliye Excel aur Worker **kabhi disagree nahi kar sakte** ki drawn site kaha hai.

### Proof (reproducible)

```bash
cd location-extract/verify
./run_parity_check.sh
```

Node pe actual Worker code chalta hai, Python pe VBA ka 1:1 port, dono ke results compare hote hain.
Network use nahi hota. Results:

```
matched 5000/5000   value mismatches 0   accept/reject mismatches 0
boundary/JSON mismatches 0/5000

sonwadi_site       17.99088747, 74.43575012  11786.13 m2   4 corners
closed_ring_dup    18.18110247, 74.61038296   5886.68 m2   3 corners
consecutive_dup    18.52093000, 73.85774000   8812.66 m2   3 corners
triangle           19.07642637, 72.87999526 181507.70 m2   3 corners
long_12pt          17.45033805, 75.12204511  70111.19 m2  12 corners
southern_hemi     -23.55102000, -46.63280800  11359.83 m2   4 corners
tiny_1m            18.00000450, 74.00000450      0.95 m2   4 corners

PARITY OK
```

5000 generated polygons (3–43 vertices, star-shaped, deterministic PRNG) + 7 hand-built edge
cases + 12 reject cases. Latitude, longitude, area, vertex-count, cleaned boundary text aur
`sitePolygon` wire JSON — sab **byte-for-byte identical**.

Structural lint bhi clean: 80 procedures, 0 duplicate names, 0 missing `GoTo` labels,
0 unbalanced blocks, 0 lines over VBA ki 1023-char limit.

Aur v2.1 ke existing code ki safety: repo ke v1.9.2 se procedure-level diff liya —
46 shared procedures me se 37 byte-identical, baaki 9 me sirf **tumhare apne v2.0/v2.1
changes** hain (English MsgBox text, `setTimeouts` casing, `zoom=18`, VILLAGE_DB + Google
geocode tiers). Maine khud se ek bhi existing line nahi badli.

---

## 4. Deploy (2 minute)

**Complete+Auto build (v2.4 — recommended):**

1. Excel → `Alt+F11` → VBE.
2. Agar `modSolarEPCDrawnLocation` module dikhe → **Remove**.
3. Purana `modSolarEPCResource` module → **Remove** (backup export kar lena).
4. Right-click project → **Import File…** → `location-extract/modSolarEPCResource.bas`.
5. Save karke workbook dobara kholo. Ab har draw+SAVE ke ~5s andar RESOURCE_DB
   apne aap bharega, aur NASA pipeline/scheduler sab pehle jaisa chalega.

> Ye file purane module ka **poora replacement** hai — isliye ThisWorkbook ke
> `Workbook_Open` (`SolarEPC_ResourceResumePending`) aur relay ke
> `SolarEPC_ResourceQueueImportedSite` calls compile hote hain. Module content
> sirf paste mat karna (Attribute line paste nahi hoti); hamesha **Import** karo.

**Complete v2.2 build (saare macros/functions):**

1. Excel → `Alt+F11` → VBE.
2. Purana `modSolarEPCResource` module → right-click → **Remove** (export karke backup rakh lo).
3. Right-click project → **Import File…** → `location-extract/modSolarEPCResource_v2.2_complete.bas`.
4. `Alt+F8` → `SolarEPC_ResourceShowDrawnLocation` → Run.

**Worker deploy nahi karna. `worker.js` / D1 schema / map HTML / cache key — kuch touch nahi hua.**
Ye module tumhare v2.1 pe based hai, isliye `SETTINGS!B11`, `VILLAGE_DB`, `_CLOUD_CFG!A17/B17`
sab waise hi kaam karte hain.

> Note: file me `Attribute VB_Name = "modSolarEPCResource"` pehli line hai, isliye import
> hone pe module ka naam sahi rahega.

---

## 5. Verify (5 minute)

1. Map kholo, ek site draw karo, **SAVE** dabao (heights bhar do).
2. `DRAWING_DATA.autoLWHTbl` me nayi row dikhegi — `Reference id = SITE-MAP-…`,
   `Coordinates = lat,lon;lat,lon;…`
3. `Alt+F8` → `SolarEPC_ResourceShowDrawnLocation`.
   - Centroid drawn shape ke beech me hona chahiye (Google Maps pe **Yes** karke dekh lo).
   - Area map ke site-stats panel wale area se match karega (dono same polygon, same projection).
4. Cross-check against the Worker: usi project ka resource run karo
   (`SolarEPC_ResourceQueueImportedSite`) → jo `RESOURCE_DB` me `Latitude (°)` /
   `Longitude (°)` aaye, wo `SolarEPC_ResourceShowDrawnLocation` wale centroid se
   **8 decimals tak same** hona chahiye. Yahi asli proof hai.

---

## 6. Safety — kya bilkul change NAHI hua

| Cheez | Status |
|---|---|
| `SolarEPC_ResourceQueueImportedSite` → Worker request body | **unchanged** (same `ResourcePolygonJSON` output) |
| Worker code, endpoints, API parameters | **unchanged, not deployed** |
| D1 schema, `centroid_lat`/`centroid_lon`, `cache_key`, `polygon_hash` | **unchanged** |
| NASA POWER request coordinates | **unchanged** — Worker ka centroid hi authoritative |
| `RESOURCE_DB` existing values, `Data Status`, resource numbers | **untouched** |
| `Location` label logic (VILLAGE_DB → Google → Nominatim → BigDataCloud) | **unchanged** |
| `tblDrawingData` (obstacles) | **never read** for the centroid, same as before |

`SolarEPC_ResourceFillDrawnCentroids` sirf **blank** coordinate cells bharta hai, confirm
dialog ke baad, aur `HasFormula` cells ko chhod deta hai. Manually entered values kabhi
overwrite nahi hote.

### Ek behaviour note (jaan-na zaroori hai)

`ResourceProjectTableRow` row reuse karta hai jab (Project ID + centroid match ho) **ya**
(Project ID match ho aur coordinates blank ho). Agar tum `SolarEPC_ResourceFillDrawnCentroids`
se coordinates pehle bhar dete ho, to baad ka NASA import us row ko **centroid se match**
karega — aur kyunki local aur Worker centroid identical hain (Section 3), wo **wahi row**
milegi. Yaani duplicate row nahi banti. Agar site dobara alag draw ki gayi, to centroid
badlega aur behaviour wahi rahega jo aaj hai (nayi row).

---

## 7. Troubleshooting

| Problem | Reason / fix |
|---|---|
| "No drawn SITE row was found in DRAWING_DATA." | Map me SAVE nahi dabaya, ya `DRAWING_DATA` sheet / `autoLWHTbl` table ka naam alag hai. Constants module ke top pe hain: `DRAWING_SHEET`, `SITE_TABLE`. |
| Coordinates dikh rahe hain par centroid nahi milta | `Coordinates` text me pair `lat,lon` format me hona chahiye, `;` se separated. `ResourceParseSitePolygon` exact reason `ErrorText` me deta hai (blank cell / non-decimal / WGS84 range / <3 distinct vertices / >1000 vertices). |
| Centroid Worker wale se differ karta hai | Possible nahi agar dono same polygon pe chale. Check karo ki DRAWING_DATA row wahi hai jo resource run me use hui (`Reference id = SITE-MAP-<messageId>`), aur site dobara draw to nahi hui. `location-extract/verify/run_parity_check.sh` chala ke algorithm verify karo. |
| Location label blank aaya | Coordinates mil gaye, label nahi — ye geocoding tier ka issue hai. `_CLOUD_CFG!A14/B14` (Google status), `A16/B16` (VILLAGE_DB match), `A17/B17` (key source) dekho. `SolarEPC_ExactSiteLocation` / `…Area` / `…Boundary` label pe depend hi nahi karte. |
| Clipboard copy kaam nahi kiya | `ResourceCopyToClipboard` `MSForms.DataObject` use karta hai. Kisi stripped environment me na mile to MsgBox me coordinates already dikhe hote hain — waha se copy kar lo. |

---

## 8. Files

| File | What it is |
|---|---|
| `location-extract/modSolarEPCResource.bas` | **v2.4 COMPLETE+AUTO** — module `modSolarEPCResource`: poora legacy module + automatic watcher. **Yahi import karo (purana module replace karta hai).** |
| `location-extract/modSolarEPCResource_v2.2_complete.bas` | **Complete v2.2** — saare macros + public functions (purana modSolarEPCResource replace karta hai) |
| `location-extract/SolarEPC_v2.2_LocationExtract_patch.zip` | Dono builds + guide + harness ek zip me |
| `location-extract/IMPLEMENTATION_GUIDE.md` | This guide |
| `location-extract/vba_lint.py` | Structural linter used on the module (blocks, labels, duplicates, line length) |
| `location-extract/verify/run_parity_check.sh` | Reproducible Worker-vs-VBA parity proof |
| `location-extract/verify/worker_algo.mjs` | Verbatim copy of the Worker's `validateSitePolygon` + `polygonCentroidLocalProjection` |
| `location-extract/verify/vba_port.py` | 1:1 Python port of the new VBA functions |
| `location-extract/verify/parity.mjs`, `cases.mjs`, `cases_worker.py` | Case generators (5000 random + 7 hand-built) |

Nothing here is deployed. `k12/`, `resource-range6/`, `map_main.html`, `public/` — untouched.

---

## Legacy users (module v2.1 / v2.2 / v2.3 jo already workbook me hai)

Apna purana `modSolarEPCResource` **mat hatao**. Uske SAATH ye add-on import karo:

**`modSolarEPCAutoLocation.bas`** (zip: `SolarEPC_v2.6_addon.zip`)

- Draw + SAVE -> ~5-15s me RESOURCE_DB ke blank lat/lon + blank Location apne aap bharne lagte hain
- Row na ho to ~10s me row apne aap ban jaati hai
- `SolarEPC_AutoLocationDebug` = poori chain ka read-only report
- Saare helpers `Private` hain, isliye purane module se koi naam nahi takrata
- Workbook khulte hi watcher ON (`Auto_Open`), band hote OFF
- **Note:** agar import ke baad "Ambiguous name detected: Auto_Open" aaye
  (matlab kisi doosre module me bhi Auto_Open hai), to add-on ki
  `Auto_Open` sub hata dena aur ThisWorkbook `Workbook_Open` me ek line
  jod dena: `SolarEPC_AutoLocationStart`
- v2.4+ complete module ke saath ye add-on MAT lagana (watcher double ho jayega)
