# Solar EPC Resource — v1.9.4

**Date:** 2026-09-02
**Commit:** (see below)
**Branch:** `arena/01a06261-solar-designed-map`

---

## 1. Format badal gaya — ab sirf `Gaon, Taluka`

| | Format | Example |
|---|---|---|
| v1.9.2 | `Town, District` | `Phaltan, Satara` |
| v1.9.3 | `Gaon, Taluka, District/State` | `Rahegaon, Vaijapur, Chhatrapati Sambhajinagar` |
| **v1.9.4** | **`Gaon, Taluka`** | **`Rahegaon, Vaijapur`** |

**District aur State ab nahi aate.** Bilkul nahi.

```
Gaon   priority (exact village): village → hamlet → suburb →
                                 locality → neighbourhood → town → city
Taluka: county  (India me Nominatim ka county = taluka)
```

- Taluka nahi mila → sirf gaon ka naam (kuch invent nahi hota).
- Gaon jo khud taluka HQ ho (jaise Baramati) → sirf ek baar: `Baramati`.

---

## 2. ⭐ Naya: Google tier — ASLI gaon ke liye (optional)

**Problem:** OSM me Indian revenue-village ke naam hote hi nahi.
`Sonwadi Bk.` Satara district me OSM me **exist hi nahi karta** (maine
bounded search se verify kiya — result empty aaya). Isliye Nominatim use
kabhi nahi de sakta.

**Solution:** Google ke paas Indian revenue-village coverage hai. Ab ek
optional Google tier sabse pehle try hota hai.

### Setup (30 second)

`_CLOUD_CFG` sheet me:

| Cell | Value |
|---|---|
| **B5** | Apni **Google Geocoding API key** |

> B5 pehle khali tha — isliye koi existing setting overwrite nahi hoti.

**Key khali rakho to kuch bhi change nahi hota** — purana
Nominatim → BigDataCloud flow chalta rahega. Zero risk.

### Google tier kya karta hai

| Google field | Matlab | Label me |
|---|---|---|
| `locality` | exact gaon | pehla part |
| `administrative_area_level_3` | taluka | dusra part |
| `administrative_area_level_2` | district | ❌ skip |
| `administrative_area_level_1` | state | ❌ skip |

Agar `locality` na mile to `sublocality` → `neighborhood` try hota hai.

---

## 3. Is folder me kya hai

| File | Kya karna hai |
|---|---|
| `modSolarEPCResource_k12.bas` | **Ye import karo** |
| `README.md` | Ye instructions |

---

## 4. Excel me import kaise kare

1. Workbook kholo → **Alt+F11**.
2. **Purana module delete karo** (`modSolarEPCResource_k12` → right-click →
   **Remove**). ⚠️ Skip mat karna.
3. **File → Import File** → `modSolarEPCResource_k12.bas`.
4. **Ctrl+S**.
5. **Debug → Compile VBAProject** → error nahi aana chahiye.

> ⚠️ **Duplicate module ka issue**
> Is file me `Attribute VB_Name` nahi hai, isliye Excel module ka naam
> **filename** se leta hai. Purana module maujood hoga to Excel naye ko
> `modSolarEPCResource_k121` naam dega → dono me same `Public Sub` hone se
> **"Ambiguous name detected"** compile error. Pehle delete, phir import.

---

## 5. Verify

**Naya site:** ek naya site run karo → `RESOURCE_DB` ka **Location** column
dekho. `Gaon, Taluka` aana chahiye.

**Purani rows:** **Alt+F8** → `SolarEPC_ResourceFillLocations` → Run.
Sirf **blank** cells bharenge.

### Expected results (real Nominatim data, Google key ke bina)

| Site | v1.9.4 output |
|---|---|
| Rahegaon, Vaijapur | `Rahegaon, Vaijapur` |
| Pachod (county = "Paithan taluka") | `Pachod, Paithan` |
| Baramati (town == taluka HQ) | `Baramati` |
| Kharadi | `Kharadi, Pune` |

---

## 6. Safety — kya change NAHI hua

| Cheez | Status |
|---|---|
| Exact centroid lat/lon | ✅ Wahi (koi nearest-town search nahi) |
| Manually typed Location | ✅ Kabhi overwrite nahi hota |
| Sirf blank cells fill | ✅ Wahi |
| Network fail | ✅ Location blank, import block nahi hota |
| NASA / resource / cache logic | ✅ Unchanged |
| MAP / DRAWING_DATA / SAVE / ACK | ✅ Unchanged |
| Worker (`worker_k12.js`) | ✅ Unchanged |

Sirf `.bas` files change hui hain (3 copies, sab byte-identical).

---

## 7. ⚠️ Important — Google output verify karna padega

Maine **parser ko** test kiya hai realistic Google response structure par —
wo sahi se `Sonwadi Bk` + `Phaltan` nikal raha hai:

```
locality                    → Sonwadi Bk
administrative_area_level_3 → Phaltan
FINAL LABEL                 → Sonwadi Bk, Phaltan   ✅
```

Par **ye maine live Google API se verify nahi kiya** — mere sandbox se
`maps.googleapis.com` reachable nahi hai, aur mere paas key bhi nahi hai.

Isliye B5 me key daalne ke baad **pehle ek site run karke check kar lena.**
Agar Google kuch ajeeb de (ya `REQUEST_DENIED` aaye), to module automatically
Nominatim par fall back ho jata hai — koi error nahi aayega.

Agar Google galat locality de to B5 khali kar do → turant purana behavior.

---

## 8. Rollback

1. Purana module import karo (v1.9.3 = commit `406658a` se pehle wala,
   ya v1.9.2 = `183a516`).
2. Ya B5 khali kar do → Google tier off.
3. Location sirf descriptive column hai — weather/resource data par koi asar
   nahi. Column clear kar do to bhi sab chalega.

---

## 9. Cleanup candidates (abhi touch nahi kiya)

- 🔶 `k12/` aur `resource-range6/` duplicate folders.
- 🔶 `map_main.html` (root) vs `public/index.html` me ~9.5 KB drift.
  Firebase `public/` serve karta hai — root file deployed source of truth
  **nahi** hai.
- 🔶 Repo root par koi README nahi hai.
