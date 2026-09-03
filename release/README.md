# Solar EPC Resource — v2.3

**Date:** 2026-09-03
**Branch:** `arena/01a06261-solar-designed-map`

---

## 1. Kya badla (v2.2.1 → v2.3)

**Geoapify tier hata diya.** Bas itna.

Teen Maharashtra sites par head-to-head measurement me Geoapify ne do par
Nominatim jaisa hi jawab diya aur teesre par **galat** jawab diya (usne paas ke
school se "Rotegaon Rly Stn" hamlet utha liya). Matlab wo rate-limit ki jagah
to deta tha, par accuracy kuch nahi badhata tha.

| Cheez | v2.2.1 | v2.3 |
|---|---|---|
| Geoapify tier | tha | ❌ **hataya** |
| `SETTINGS!B11` | Geoapify key | **spare Google key slot** |
| `_CLOUD_CFG` A18/B18, A19/B19 | likhe jate the | ab nahi likhe jate |
| Functions / Subs | 37 / 25 | **36 / 24** |
| Size | 85,219 B | **81,707 B** |

Baaki sab kuch unchanged.

---

## 2. Format — `Exact Area, Taluka, District`

```
Sonwadi Bk., Phaltan, Satara
Rahegaon, Vaijapur, Chhatrapati Sambhajinagar
```

| Part | Kya hai | Priority |
|---|---|---|
| **1. Exact Area** | Centroid jahan draw hua, us area ka sabse specific naam | `village → hamlet → suburb → locality → neighbourhood → town → city` |
| **2. Taluka** | Taluka / tehsil | Nominatim `county` (`"Phaltan"`, `"Paithan taluka"`) |
| **3. District** | District | Nominatim `state_district` (`"Satara"`) |

**State nahi aata.** Jo level nahi milta wo skip ho jata hai — **kabhi kuch
invent nahi hota.**

---

## 3. Resolution order

```
0)  VILLAGE_DB      tumhari apni sheet (free, offline) — abhi use nahi ho raha
0b) Google          sirf tab chalta hai jab Cloud project me billing ON ho
1)  Nominatim       free, key nahi chahiye   ← asli kaam yahi karta hai
2)  BigDataCloud    free, key nahi chahiye
```

`_CLOUD_CFG!A20/B20` me likha jata hai ki kaunse resolver ne jawab diya
(`VILLAGE_DB` / `GOOGLE` / `NOMINATIM` / `BIGDATACLOUD`).

### ⚠️ Google ke baare me — sach

Google tier **code me hai**, par abhi **kaam nahi karta**, kyunki Google
Geocoding API ko **billing-enabled Cloud project** chahiye. Test kiya:

```
REQUEST_DENIED — "You must enable Billing on the Google Cloud Project"
```

`SETTINGS!B4` aur `SETTINGS!B11` **do alag projects** hain — dono me billing
off hai. Billing on karoge to bas key paste karna, code change ki zaroorat
nahi: chain `B4 → B11 → _CLOUD_CFG!B5` pehle non-empty value leti hai.

> Google Maps **JavaScript** API (`google.maps.Geocoder`) se bhi billing nahi
> bachti — wo wahi API hai, bas JS wrapper me. Aur wo browser me chalta hai,
> Excel VBA me nahi.

---

## 4. ⭐ Sabse zaroori: manual Lokasyon **permanent** hai

Jo tum haath se likhoge wo **kabhi overwrite nahi hota** — na import par, na
back-fill par. Code me dono jagah guard hai:

```vba
If Len(Trim$(CStr(Cell.Value2))) > 0 Then Exit Sub   ' non-blank → chho do
```

**Isliye ~20-25 sites jinme gaon ka naam OSM me nahi hai, unme bas ek baar
haath se likh do — wo hamesha ke liye rehga.**

Kisi cell ko dobara auto-resolve karana ho to pehle us cell ko **clear** karo,
phir `SolarEPC_ResourceFillLocations` chalao.

### Verified results

| Site | Output | Status |
|---|---|---|
| Rahegaon | `Rahegaon, Vaijapur, Chhatrapati Sambhajinagar` | ✅ |
| Pachod | `Pachod, Paithan, Chhatrapati Sambhajinagar` | ✅ |
| Baramati | `Baramati, Pune` | ✅ |
| 17.96608273, 74.46912967 | `Phaltan, Satara` | ⚠️ **haath se likho: `Sonwadi Bk., Phaltan, Satara`** |

Tumhare ~100 sites me se **~75-80 auto-resolve** ho jayenge. Baaki 20-25 me ek
baar typing.

---

## 5. Kyun koi free geocoder `Sonwadi Bk.` nahi de sakta

Humne khud raw OSM (Overpass) query kiya 8 km radius me:

```
Vidni 3.0 km, Saskal 4.1, Vinchurni 5.6, Nirugudi 5.8, Pimprad 6.8
```

**Sonwadi Bk. OSM me hai hi nahi.** Aur ye providers sab OSM ke upar bane hain:

| Provider | Data source | Sonwadi Bk.? |
|---|---|---|
| Nominatim | OpenStreetMap | ❌ |
| **BigDataCloud** | apna data | ❌ |
| Geoapify | OSM + Who's On First | ❌ (hataya) |
| **LocationIQ** | OSM/Nominatim ka wrapper | ❌ |
| Wikidata / GeoNames | apna | ❌ |
| Google | proprietary | ✅ **par billing chahiye** |

LocationIQ technically free hai (5,000/day, card nahi chahiye), par wo
**Nominatim ki accuracy inherit karta hai** — matlab wahi galat jawab. Isliye
add karne ka koi fayda nahi.

---

## 6. Import kaise kare

1. Workbook kholo → **Alt+F11**.
2. **Purana module delete karo** (`modSolarEPCResource_k12` → right-click →
   **Remove**). ⚠️ Skip mat karna.
3. **File → Import File** → `modSolarEPCResource_k12.bas`.
4. **Ctrl+S**.
5. **Debug → Compile VBAProject** → error nahi aana chahiye.

> ⚠️ Is file me `Attribute VB_Name` nahi hai, isliye Excel module ka naam
> **filename** se leta hai. Purana module maujood hoga to naye ko
> `modSolarEPCResource_k121` naam milega → dono me same `Public Sub` hone se
> **"Ambiguous name detected"** compile error. Pehle delete, phir import.

**Purani rows bharne ke liye:** Alt+F8 → `SolarEPC_ResourceFillLocations` →
Run. Sirf **blank** cells bharenge.

---

## 7. Safety — kya change NAHI hua

| Cheez | Status |
|---|---|
| Exact centroid lat/lon | ✅ Wahi (koi nearest-town search nahi) |
| Manually typed Location | ✅ **Kabhi overwrite nahi hota** |
| Sirf blank cells fill | ✅ Wahi |
| Network fail | ✅ Location blank, import block nahi hota |
| NASA / resource / cache logic | ✅ Unchanged |
| MAP / DRAWING_DATA / SAVE / ACK | ✅ Unchanged |
| Worker (`worker_k12.js`) | ✅ Unchanged |

Sirf `.bas` files change hui hain (3 copies, sab byte-identical).

---

## 8. Diagnostics (`_CLOUD_CFG` sheet)

| Cell | Matlab |
|---|---|
| A10/B10 | last resource-start error timestamp |
| A11/B11 | last resource-start error detail |
| A12/B12 | processing mode (RANGE12 / RANGE6 / MONTH) |
| A13/B13 | last range processing time (ms) |
| A14/B14 | Google geocode status |
| A15/B15 | Google geocode detail |
| A16/B16 | VILLAGE_DB match |
| A17/B17 | Google key source |
| A20/B20 | **kaunse resolver ne jawab diya** |

A18/B18 aur A19/B19 ab nahi likhe jate (Geoapify hata diya).

---

## 9. Rollback

```bash
git checkout 41dc299 -- k12/ resource-range6/ release/   # v2.2.1
git checkout 55b4e4f -- k12/ resource-range6/ release/   # v2.2
git checkout 183a516 -- k12/ resource-range6/ release/   # v1.9.2
```

Location sirf descriptive column hai — weather/resource data par koi asar nahi.
