# v3.8 — Date/Time stamp fix (site draw karne pe date-time nahi aa raha tha)

## Asli bug kya tha (root cause)

`ResourceStampLocalRunTime` (v3.5/v3.7) me ye do lines thi:

```vb
Cell.NumberFormat = "dd-mm-yyyy hh:nn:ss"   ' <-- YAHI BUG HAI
Cell.Value2 = Now
```

Problem:

1. **`"nn"` Excel ka valid number-format code NahI hai.** `nn` sirf VBA ke
   `Format$()` function me minutes ke liye hota hai. `Range.NumberFormat`
   me minutes ka code **`mm`** hota hai (Excel context se samajh leta hai
   ki month hai ya minute). Invalid format string dete hi Excel **runtime
   error 1004** throw karta hai.
2. Wo error **value likhne se PEHLE** aata tha (`NumberFormat` line
   `Value2 = Now` se upar thi). Isliye cell me kabhi kuch likha hi nahi
   jaata tha.
3. `ResourceStampLocalRunTime` ke andar `On Error Resume Next` sirf
   column-create block tak tha — format line ke pehle `On Error GoTo 0`
   ho chuka tha, isliye 1004 **caller tak propagate** hota tha
   (`FillResourceDbForSite` ka `Failed:` handler kha jaata tha). Net
   result: **lat/lon/Location bhar jaate the, Date/Time hamesha blank.**

Ek secondary wajah bhi thi: watcher me stamp sirf tab call hota tha jab
usi call me kuch *naya* likha gaya ho (`DidText` non-empty). Agar lat/lon
pehle se bhare the (e.g. NASA import pehle ho gaya) to stamp attempt hi
nahi hota tha.

## v3.8 me kya badla

`resource-v38/modSolarEPCResource.bas` — poora drop-in replacement module:

1. **Value pehle, format baad me.** `Cell.Value2 = Now` ab format se
   pehle chalta hai — formatting kabhi bhi fail ho to bhi stamp cell me
   lag chuka hota hai.
2. **Sahi format code**: `"dd-mm-yyyy hh:mm:ss"`.
3. **Stamp function poora error-sealed** (`On Error Resume Next` end tak)
   — stamping ki koi bhi dikkat lat/lon/Location fill ko fail nahi kar
   sakti. Har outcome `_CLOUD_CFG!B21` me diagnostic ke roop me record
   hota hai (OK / NOT BLANK / FORMULA CELL / WRITE FAIL / NO COLUMN).
4. **Har fill event pe stamp try hota hai** (function khud blank-only
   hai, isliye kuch overwrite nahi hota) — watcher sweep,
   `SolarEPC_ResourceFillDrawnCentroids`, aur NASA summary import teeno.
5. **`ResourceProjectRowFilled` ab Date/Time bhi dekhta hai** — lat/lon
   bhare hain par Date/Time blank hai to watcher us row ko dobara
   process karke stamp laga deta hai (workbook reopen ki zaroorat nahi).
6. **Fuzzy column match**: "Date/Time" header me Alt+Enter line break,
   non-breaking space, double space, "Date / Time", "Date Time",
   "Date & Time" — sab pakde jaate hain. Column na ho to
   "Run Date-Time (Local)" auto-create hoti hai (ListColumns.Add ab
   position argument ke bina — total-row / odd layouts pe bhi safe) aur
   uska stray calculated-column content clear kiya jaata hai.
7. **Nayi one-time repair macro**: `SolarEPC_ResourceStampMissingDates`
   — purani rows jinke lat/lon bhare hain par Date/Time blank hai
   (purane bug ke shikar), unhe abhi ke local time se stamp kar deti
   hai. Blank-only; manual/formula cells kabhi nahi chhute.
8. Debug macro (`SolarEPC_DrawnLocationDebug`) ab matched row ka
   Date/Time value bhi dikhata hai aur blank hone pe agla step batata hai.

## Install (waise hi jaise pehle)

1. VBE me purana `modSolarEPCResource` module REMOVE karo
   (aur `modSolarEPCDrawnLocation` agar ho to wo bhi).
2. Right-click project → Import File… → `resource-v38/modSolarEPCResource.bas`.
3. Save karke workbook dobara kholo (Auto_Open watcher chalu karega).
4. Purani rows ka blank Date/Time bharna ho to ek baar
   Alt+F8 → `SolarEPC_ResourceStampMissingDates`.

## Verify

- Nayi site draw + SAVE karo → 5–10s me RESOURCE_DB me lat/lon, Location
  ke saath **Date/Time bhi** aa jayega (`dd-mm-yyyy hh:mm:ss`).
- `_CLOUD_CFG!B21` me "OK row N = ..." dikhna chahiye.
- Kuch bhi na aaye to Alt+F8 → `SolarEPC_DrawnLocationDebug` — report me
  "Date/Time stamp:" line exact reason batayegi.
