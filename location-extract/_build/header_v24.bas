Attribute VB_Name = "modSolarEPCResource"
Option Explicit

'==========================================================================
' SOLAR EPC - NASA POWER HOURLY RESOURCE MODULE + AUTOMATIC LOCATION FILL
' Version 2.4
'
' YE FILE TUMHARE PURANE modSolarEPCResource KA POORA REPLACEMENT HAI.
'   - v2.1/v2.2 ka SARA kaam waise hi chalta hai: NASA POWER hourly import,
'     range/month scheduler, RESOURCE_DB write, Location auto-fill,
'     VILLAGE_DB/Google/Nominatim/BigDataCloud tiers, sab public macros
'     (ConfigurePeriod, QueueImportedSite, RetryLatestSite, FillLocations,
'     MakeVillageDb, AddVillage, Resume, Stop, ShowLastError,
'     ResumePending, ProcessNext) - isliye tumhara ThisWorkbook,
'     modSolarEPCCloudRelay aur sheet buttons sab compile hote hain.
'   - PLUS v2.3 ka AUTOMATIC watcher: workbook khulte hi chalu, har 5s
'     DRAWING_DATA.autoLWHTbl dekhta hai; nayi ya dobara-draw hui SITE row pe
'     RESOURCE_DB ke BLANK Latitude(deg)/Longitude(deg) me exact centroid
'     (Worker-identical, offline) aur BLANK Location me "Exact Area, Taluka,
'     District" apne aap likh deta hai. Koi Alt+F8 nahi, koi popup nahi -
'     sirf status bar pe ek line.
'
' IMPORT KA TAREEKA (zaroori - warna compile error):
'   1. VBE me agar "modSolarEPCDrawnLocation" naam ka module ho to REMOVE karo.
'   2. Purana "modSolarEPCResource" module REMOVE karo (backup export optional).
'   3. Right-click project -> Import File... -> ye .bas file.
'   4. Save karke workbook dobara kholo.
'
' SAFETY (auto-fill ke rules):
'   - sirf BLANK cells bharte hain; manual entry / formula / NASA ke values
'     kabhi overwrite nahi hote
'   - NASA numbers, Data Status, cache_key, polygon_hash - unchanged
'   - label offline na mile to lat/lon phir bhi bharte hain; label limited
'     retry ke baad ruk jaata hai, galat value kabhi nahi likhi jaati
'   - rokna ho: SolarEPC_DrawnLocationAutoStop ; turant sweep:
'     SolarEPC_DrawnLocationAutoNow ; manual dekhna ho:
'     SolarEPC_ResourceShowDrawnLocation
'
' CENTROID MATH = WORKER KA MATH (polygonCentroidLocalProjection port):
' 5000 generated polygons + 7 edge cases pe byte-for-byte verified
' (location-extract/verify/run_parity_check.sh).
'==========================================================================

'Watcher tuning (baaki sab constants module ke andar pehle se hain).
Private Const AUTO_TICK_SECONDS As Long = 5
Private Const AUTO_NOROW_RETRIES As Long = 12    'RESOURCE_DB row na mile to ~1 min retry
Private Const AUTO_LABEL_RETRIES As Long = 6     'label offline ho to limited retry

