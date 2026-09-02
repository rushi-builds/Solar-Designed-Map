# Solar EPC Resource — v1.9.5

**Date:** 2026-09-02
**Commit:** `8dfa491`
**Branch:** `arena/01a06261-solar-designed-map`

---

## 1. Format — `Exact Area, Taluka, District`

```
Sonwadi Bk., Phaltan, Satara
Rahegaon, Vaijapur, Chhatrapati Sambhajinagar
```

| Part | Kya hai | Priority |
|---|---|---|
| **1. Exact Area** | Centroid jahan draw hua, us area ka sabse specific naam — gaon ho, hamlet ho, suburb ho ya colony | `village → hamlet → suburb → locality → neighbourhood → town → city` |
| **2. Taluka** | Taluka / tehsil | Nominatim `county` (`"Phaltan"`, `"Paithan taluka"`) |
| **3. District** | District | Nominatim `state_district` (`"Satara"`) |

**State nahi aata.** District zaroor aata hai.

Jo level nahi milta wo skip ho jata hai — **kabhi kuch invent nahi hota.**
Gaon jo khud taluka HQ ho (jaise Baramati) → `Baramati, Pune`.

### Version history (meri galti ki wajah se)

| Version | Format | Status |
|---|---|---|
| v1.9.2 | `Phaltan, Satara` | gaon/taluka hi nahi aata tha |
| v1.9.3 | `Rahegaon, Vaijapur, Chhatrapati Sambhajinagar` | sahi tha, par kabhi State bhi ghus jata tha |
| v1.9.4 | `Rahegaon, Vaijapur` | ❌ **galat** — maine district hata diya |
| **v1.9.5** | **`Rahegaon, Vaijapur, Chhatrapati Sambhajinagar`** | ✅ abhi wala |

---

## 2. ⭐ Google tier — ASLI gaon ke liye (optional par important)

**Problem:** OSM me Indian revenue-village ke naam hote hi nahi.
`Sonwadi Bk.` Satara district me OSM me **exist hi nahi karta** — maine
bounded Nominatim search se verify kiya, result **empty** aaya. Isliye
Nominatim se wo naam kabhi nahi milega, chahe koi bhi logic lagao.

**Solution:** Google ke paas Indian revenue-village coverage hai.

### Setup (30 second)

`_CLOUD_CFG` sheet me:

| Cell | Value |
|---|---|
| **B5** | Apni **Google Geocoding API key** |

> B5 pehle se khali tha — koi existing setting overwrite nahi hoti.
> **Key khali rakho to kuch bhi change nahi hota** — purana
> Nominatim → BigDataCloud flow chalta rahega. Zero risk.

### Google se kya uthta hai

| Google field | Matlab | Label me |
|---|---|---|
| `locality` | exact gaon | part 1 |
| `administrative_area_level_3` | taluka | part 2 |
| `administrative_area_level_2` | district | part 3 |
| `administrative_area_level_1` | state | ❌ skip |

`locality` na mile to `sublocality` → `neighborhood` try hota hai.

**Parser test result (tumhare example coordinate ke liye):**
```
Google @ 17.990387, 74.435250  →  Sonwadi Bk, Phaltan, Satara   ✅
```

---

## 3. Is folder me kya hai

| File | Kya karna hai |
|---|---|
| `modSolarEPCResource_k12.bas` | **Ye import karo** |
| `README.md` | Ye instructions |

---

## 4. Import kaise kare

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

---

## 5. Verify

**Naya site:** ek naya site run karo → `RESOURCE_DB` ka **Location** column
dekho.

**Purani rows:** **Alt+F8** → `SolarEPC_ResourceFillLocations` → Run.
Sirf **blank** cells bharenge.

### Expected (real Nominatim data, Google key ke bina)

| Site | v1.9.5 output |
|---|---|
| Rahegaon | `Rahegaon, Vaijapur, Chhatrapati Sambhajinagar` |
| Pachod (county = "Paithan taluka") | `Pachod, Paithan, Chhatrapati Sambhajinagar` |
| Baramati (town == taluka HQ) | `Baramati, Pune` |
| Phaltan coords (OSM me gaon nahi) | `MSEB Colony, Phaltan, Satara` |

Last wala case me `MSEB Colony` isliye aata hai kyunki **OSM me wahan koi
gaon mapped hi nahi** — ye sabse specific cheez hai jo OSM ko mili. Asli
`Sonwadi Bk.` ke liye **B5 me Google key daalo.**

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

## 7. ⚠️ Google output live verify nahi hua

Maine **parser** ko test kiya hai realistic Google response structure par —
wo `Sonwadi Bk` + `Phaltan` + `Satara` sahi nikal raha hai. Par **live Google
API se verify nahi kiya** — mere sandbox se `maps.googleapis.com` reachable
nahi hai aur mere paas key bhi nahi.

Isliye B5 me key daalne ke baad **pehle ek site run karke check kar lena.**
Agar Google kuch ajeeb de ya `REQUEST_DENIED` aaye, to module automatically
Nominatim par fall back ho jata hai — koi error nahi aayega. Pasand na aaye
to B5 khali kar do → turant purana behavior.

---

## 8. Rollback

1. Purana module import karo (v1.9.4 = `e90a959`, v1.9.3 = `fdf0de7`,
   v1.9.2 = `183a516`).
2. Ya B5 khali kar do → Google tier off.
3. Location sirf descriptive column hai — weather/resource data par koi asar
   nahi.

---

## 9. Cleanup candidates (abhi touch nahi kiya)

- 🔶 `k12/` aur `resource-range6/` duplicate folders.
- 🔶 `map_main.html` (root) vs `public/index.html` me ~9.5 KB drift.
  Firebase `public/` serve karta hai — root file deployed source of truth
  **nahi** hai.
- 🔶 Repo root par koi README nahi hai.
