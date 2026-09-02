# Solar EPC Resource — v1.9.3

**Date:** 2026-09-02
**Commit:** `fdf0de7`
**Branch:** `arena/01a06261-solar-designed-map`

---

## 1. Is folder me kya hai

| File | Kya karna hai |
|---|---|
| `modSolarEPCResource_k12.bas` | **Ye import karo.** Updated VBA module (v1.9.3) |
| `README.md` | Ye instructions |

---

## 2. Kya change hua (v1.9.2 → v1.9.3)

### Problem
`RESOURCE_DB` ka **Location** column sirf `Town, District` tak ruk jata tha
(jaise `Phaltan, Satara`). Village aur Taluka dono miss ho jate the.

Do wajah thi:
1. Priority ulti thi — `town → city → village → hamlet`. Town milte hi
   village ignore ho jata tha.
2. Sirf 2 level hi output hote the, isliye **Taluka** kabhi nahi aata tha.

### Fix
Ab Location **most specific locality** use karta hai:

```
Locality  priority : village → hamlet → suburb → locality →
                    neighbourhood → town → city
Admin     priority : county/taluka → district → state_district → state
Output    format   : "Village, Taluka, District"
```

Saath me:
- Nominatim `zoom` 16 → **18** (village/hamlet/suburb/neighbourhood sirf
  high zoom par milte hain).
- Admin suffix saaf: `"Phaltan taluka" → "Phaltan"`,
  `"Satara district" → "Satara"`. **Locality ka naam kabhi alter nahi hota.**
- Duplicate collapse: town jo apne taluka ka HQ ho, wo sirf ek baar.

### Before / After (live Nominatim data se verify kiya)

| Site | v1.9.2 (purana) | v1.9.3 (naya) |
|---|---|---|
| Rahegaon, Vaijapur | Rahegaon, Chhatrapati Sambhajinagar | **Rahegaon, Vaijapur, Chhatrapati Sambhajinagar** |
| Pachod (county = "Paithan taluka") | Pachod, Chhatrapati Sambhajinagar | **Pachod, Paithan, Chhatrapati Sambhajinagar** |
| Koi locality key nahi | Pune | **Kharadi, Pune** |

---

## 3. Excel me import kaise kare

1. Workbook kholo → **Alt+F11** (VB Editor).
2. **Purana module delete karo** — `modSolarEPCResource_k12` (ya jo naam hai)
   par right-click → **Remove**. ⚠️ Step 2 skip mat karna, neeche dekho.
3. **File → Import File** → `modSolarEPCResource_k12.bas` select karo.
4. **Ctrl+S** → workbook save karo.
5. VB Editor me **Debug → Compile VBAProject** → koi error nahi aana chahiye.

> ⚠️ **Duplicate module ka issue**
> Is file me `Attribute VB_Name` nahi hai, isliye Excel module ka naam
> **filename** se leta hai. Agar purane naam ka module pehle se maujood hoga,
> to Excel naye ko `modSolarEPCResource_k121` jaisa naam dega — aur dono me
> same `Public Sub` hone se **"Ambiguous name detected"** compile error aayega.
> Isliye pehle purana module delete karo, phir import karo.

---

## 4. Verify kaise kare

**Naya site (recommended):**
1. Ek naya site run karo.
2. `RESOURCE_DB` me **Location** column check karo.
3. Ab usme `Village, Taluka, District` aana chahiye (jaise
   `Rahegaon, Vaijapur, Chhatrapati Sambhajinagar`).

**Purani rows ke liye (ek saath sab fill):**
1. **Alt+F8** → `SolarEPC_ResourceFillLocations` → Run.
2. Sirf **blank** Location cells bharenge — manually type kiya hua kuch
   change nahi hoga.

---

## 5. Safety — kya change NAHI hua

| Cheez | Status |
|---|---|
| Exact centroid lat/lon | ✅ Wahi (koi nearest-town search nahi) |
| Manually typed Location | ✅ Kabhi overwrite nahi hota |
| Sirf blank cells fill | ✅ Wahi |
| Network fail hone par | ✅ Location blank rehta hai, import block nahi hota |
| NASA / resource / cache logic | ✅ Bilkul unchanged |
| MAP / DRAWING_DATA / SAVE / ACK | ✅ Bilkul unchanged |
| Worker (`worker_k12.js`) | ✅ Unchanged — is release me Worker nahi chuna gaya |

Sirf 2 files change hui hain:
- `k12/modSolarEPCResource_k12.bas`
- `resource-range6/modSolarEPCResource_range6.bas` (dono byte-identical)

---

## 6. ⚠️ Known limitation — "Sonwadi Bk." nahi milega

Test coordinate `17.990387, 74.435250` par **OSM me "Sonwadi Bk." hai hi nahi**:

- Nominatim reverse (zoom 14/16/18) → `neighbourhood=MSEB Colony,
  town=Phaltan, county=Phaltan, state_district=Satara`
- "Sonwadi" search bounded around Phaltan → **empty**
- India-wide "Sonwadi" sirf Akola / Chhatrapati Sambhajinagar / Nanded /
  Hingoli me hai — Phaltan ke paas nahi
- BigDataCloud (fallback provider) → bhi sirf Phaltan / Satara

Isliye wo point ab **`MSEB Colony, Phaltan, Satara`** deta hai
(neighbourhood → taluka → district). Ye pehle se zyada specific hai, par
`Sonwadi Bk.` nahi — kyunki **wo data OSM me exist hi nahi karta.**

`Sonwadi Bk.` ek **revenue-village** naam hai (Bk. = Budruk). Ye Indian
government/census records me hai, OpenStreetMap me nahi. Rule 7 ke mutabik
kuch bhi **fabricate/ hard-code nahi kiya gaya.**

Agar asli revenue-village naam chahiye to ek aisa provider chahiye jiske paas
Indian census/revenue coverage ho — **Google reverse geocoding** (sabse best
India village coverage; map me pehle se Google key configured hai) ya
**GeoNames**. Batao to next release me add kar deta hoon.

---

## 7. Rollback

Agar kuch gadbad lage:
1. Purana module wapas import karo (`v1.9.2`) — is branch ke previous commit
   `183a516` se mil jayega.
2. Location sirf descriptive column hai — weather/resource data par **koi
   asar nahi**. Column clear kar do to bhi sab chalega.

---

## 8. Dhyaan de (cleanup candidates)

- 🔶 `k12/` aur `resource-range6/` **duplicate folders** hain (guides identical,
  worker/bas alag). Is release me dono `.bas` files ko byte-identical rakha
  gaya hai taaki galti se purana import na ho — par inhe merge karna chahiye.
- 🔶 `map_main.html` (root) aur `public/index.html` me **drift** hai (~9.5 KB).
  Firebase `public/` serve karta hai, isliye **root file deployed source of
  truth nahi hai.**
- 🔶 Repo me koi README nahi hai (root par).
