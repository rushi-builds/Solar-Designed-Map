Attribute VB_Name = "modSolarEPCDrawnLocation"
Option Explicit

'==========================================================================
' SOLAR EPC - EXACT DRAWN-SITE LOCATION, AUTOMATIC  (v2.3)
'
' YE MODULE SAB KUCH APNE AAP KARTA HAI. Koi Alt+F8 nahi, koi popup nahi.
'
'   map pe DRAW SITE -> SAVE
'        -> (within ~5 seconds, khud)
'   RESOURCE_DB me us Project ID ki row:
'       Latitude (deg) / Longitude (deg)  = exact area-weighted centroid
'                                           (8 decimals, Worker-identical)
'       Location                          = "Exact Area, Taluka, District"
'                                           (VILLAGE_DB -> Google ->
'                                            Nominatim -> BigDataCloud)
'       ...sirf BLANK cells bharte hain. Bhara hua / manual / formula cell
'       kabhi overwrite NAHI hota. NASA ke numbers/status ko ye code chhoota
'       hi nahi.
'
' FEEDBACK: status bar pe ek line -
'   "Solar EPC: exact drawn location auto-filled -> PRJ-001 | lat lon location(Sonwadi Bk., Phaltan, Satara)"
'   aur hidden _CLOUD_CFG!A18/B18 me last auto-fill ka record.
'
' KAISE CHALTA HAI
'   Workbook khulte hi Auto_Open watcher chalu kar deta hai aur ek turant
'   sweep karta hai (pehle se drawn sites bhi bhar jaati hain). Phir har
'   AUTO_TICK_SECONDS (5s) me DRAWING_DATA.autoLWHTbl check hota hai; nayi ya
'   dobara-draw hui SITE row pe fill ho jata hai. Workbook band hone pe
'   Auto_Close scheduler saaf kar deta hai.
'   Chalu/rokna ho to: SolarEPC_DrawnLocationAutoStart / ...AutoStop,
'   turant sweep: SolarEPC_DrawnLocationAutoNow.
'   Manual dekhna ho to wahi ek macro bhi hai: SolarEPC_ResourceShowDrawnLocation
'   (zaroori nahi, sirf verification ke liye).
'
' OFFLINE BEHAVIOUR
'   Centroid hamesha offline fill hota hai. Location label ke liye internet
'   (ya VILLAGE_DB sheet) chahiye; na mile to lat/lon bhare rahenge aur label
'   kuch ticks retry hoke ruk jaayega - kabhi galat value nahi likhi jaati.
'
' MODULE NAME modSolarEPCDrawnLocation hai taaki tumhara existing
' modSolarEPCResource (NASA POWER resource module) jaisa hai waisa rahe -
' dono saath chalte hain, koi collision nahi.
'   NOTE 1: Auto_Open / Auto_Close legacy Excel names hain. Agar is workbook
'           me kahin aur bhi Auto_Open/Auto_Close defined ho (bahut rare),
'           VBA "ambiguous name" bolega - is module wale do chhote stubs
'           hata dena aur ThisWorkbook_Workbook_Open me
'           SolarEPC_DrawnLocationAutoStart call kar lena.
'   NOTE 2: Agar pichla minimal build (sirf ShowDrawnLocation wala) import
'           kiya tha, wo pehle remove karo - same module name hai.
'
' CENTROID MATH = WORKER KA MATH
'   polygonCentroidLocalProjection() (worker_k12.js) ka statement-by-statement
'   port: same 6378137 m radius, same 8-decimal rounding, same duplicate /
'   closing-point rules, same 0.01 m2 minimum. Isliye jo yaha likha jaata hai
'   wo NASA POWER ko jaane wale centroid se byte-for-byte same hota hai.
'   Verified: 5000 generated polygons + 7 edge cases identical
'   (location-extract/verify/run_parity_check.sh).
'==========================================================================

Private Const DRAWING_SHEET As String = "DRAWING_DATA"
Private Const SITE_TABLE As String = "autoLWHTbl"
Private Const RESOURCE_DB_TABLE As String = "resource_db"
Private Const CONFIG_SHEET As String = "_CLOUD_CFG"
Private Const SETTINGS_SHEET As String = "SETTINGS"
Private Const VILLAGE_SHEET As String = "VILLAGE_DB"
Private Const VILLAGE_MAX_KM As Double = 5#
'Worker constants - do not change.
Private Const EARTH_RADIUS_M As Double = 6378137#
Private Const SITE_MAX_VERTICES As Long = 1000
Private Const SITE_MIN_AREA_M2 As Double = 0.01
'Watcher tuning.
Private Const AUTO_TICK_SECONDS As Long = 5
Private Const AUTO_NOROW_RETRIES As Long = 12    'RESOURCE_DB row na mile to ~1 min retry
Private Const AUTO_LABEL_RETRIES As Long = 6     'label offline ho to limited retry

Private mAutoActive As Boolean
Private mAutoNextRun As Date
Private mProcessed As Object          'key = ReferenceID|Coordinates -> state
'v1.9.2: last reverse-geocoded coordinate cache (same coords -> same label).
Private mLastLocKey As String
Private mLastLocLabel As String

