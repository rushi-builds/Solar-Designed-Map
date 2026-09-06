import io

ROOT = '/home/user/Solar-Designed-Map/location-extract/'
s = io.open(ROOT + 'modSolarEPCResource.bas', encoding='utf-8').read()
miss = []

def rep(old, new):
    global s
    if s.count(old) != 1:
        miss.append(old[:70])
        return
    s = s.replace(old, new, 1)

# ---------------- HEADER BLOCK -> professional English ----------------
old_header = """' YE FILE TUMHARE PURANE modSolarEPCResource KA POORA REPLACEMENT HAI.
'   - v2.1/v2.2 ka SARA kaam waise hi chalta hai: NASA POWER hourly import,
'     range/month scheduler, RESOURCE_DB write, Location auto-fill,
'     VILLAGE_DB/Google/Nominatim/BigDataCloud tiers, sab public macros
'     (ConfigurePeriod, QueueImportedSite, RetryLatestSite, FillLocations,
'     MakeVillageDb, AddVillage, Resume, Stop, ShowLastError,
'     ResumePending, ProcessNext) - isliye tumhara ThisWorkbook,
'     modSolarEPCCloudRelay aur sheet buttons sab compile hote hain.
'   - v2.5: RESOURCE_DB me project ki row na ho to watcher wo row APNE AAP
'     bana deta hai (blank-only rules phir bhi lagute hain), status bar pe har
'     step ka reason dikhta hai, aur ek diagnostic macro hai:
'         SolarEPC_DrawnLocationDebug  -> poori chain ka report (read-only)
'   - v3.8: SolarEPC_FillAndStampNow - THE BIG RED BUTTON: ek click me
'     watcher ON + turant sweep + purani rows ka Date/Time Retrieval Date
'     (UTC) se LOCAL time me backfill + B21 report. Koi reopen nahi.
'   - v3.5: Date/Time stamp ab BLANK-ONLY hai aur HAR pehle fill pe
'     lagta hai (watcher auto-fill YA NASA import, jo pehle ho). Column
'     candidates: Date/Time, Run Date-Time (Local), Date-Time, DateTime.
'     Debug macro batata hai column table ke andar mili ya nahi.
'   - v3.4 LEAN: GeoNames tier HATAYA (unke database me Indian
'     revenue villages hain hi nahi - prove ho gaya). Ab user ka apna
'     "Date/Time" column bharta hai (local Now, real date value);
'     column na ho to "Run Date-Time (Local)" apne aap banti hai.
'     NASA fetch hamesha us row ke APNE centroid se hota hai - same
'     town/village ho tab bhi coordinates se alag result aata hai.
'   - v3.2: SolarEPC_ResourceRefreshLabels - purani bhari hui Location
'     rows ko village-level label pe safely upgrade karo (suffix-rule,
'     manual entries kabhi nahi chhute). Blank-only auto-fill waisa hi.
'   - v3.0: SINGLE MODULE - NASA pipeline + watcher + auto row-create +
'     debug, sab ek hi file me. Nayi RESOURCE_DB column
'     "Run Date-Time (Local)" har NASA summary import pe aapke computer
'     ka exact din + samay stamp karti hai (column khud ban jaati hai).
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
'=========================================================================="""
new_header = """' THIS FILE IS A COMPLETE REPLACEMENT FOR THE LEGACY modSolarEPCResource.
'   - Every legacy capability keeps working unchanged: NASA POWER hourly
'     import, range/month scheduler, RESOURCE_DB writes, Location auto-fill,
'     the VILLAGE_DB / Google / Nominatim / BigDataCloud label tiers and all
'     public macros (ConfigurePeriod, QueueImportedSite, RetryLatestSite,
'     FillLocations, MakeVillageDb, AddVillage, Resume, Stop, ShowLastError,
'     ResumePending, ProcessNext) - so ThisWorkbook, modSolarEPCCloudRelay
'     and every sheet button continue to compile.
'   - v2.5: when RESOURCE_DB has no row for a project, the watcher creates
'     it automatically (blank-only rules still apply); every step reports
'     its reason on the status bar; SolarEPC_DrawnLocationDebug produces a
'     read-only report of the whole chain.
'   - v3.8: SolarEPC_FillAndStampNow - one click starts the watcher, runs an
'     immediate sweep and back-fills Date/Time on existing rows from the
'     NASA Retrieval Date (UTC converted to local time). No reopen needed.
'   - v3.5: the Date/Time stamp is BLANK-ONLY and is applied on the FIRST
'     fill of a row (watcher auto-fill or NASA import, whichever happens
'     first). Accepted column names: Date/Time, Run Date-Time (Local),
'     Date-Time, DateTime. The debug macro reports whether the column was
'     found inside the table.
'   - v3.4: the GeoNames tier was REMOVED (its database provably contains no
'     Indian revenue villages). The user's own "Date/Time" column is filled
'     with the local run date/time; if absent, "Run Date-Time (Local)" is
'     created automatically. Every NASA fetch uses that row's OWN centroid -
'     two sites in the same town still receive independently fetched values.
'   - v3.2: SolarEPC_ResourceRefreshLabels safely upgrades previously filled
'     Location cells to village-level labels (suffix rule; manual entries
'     are never touched). Blank-only auto-fill is unchanged.
'   - v3.0: SINGLE MODULE - NASA pipeline + watcher + auto row creation +
'     diagnostics in one file.
'   - PLUS the v2.3 AUTOMATIC watcher: starts when the workbook opens and
'     inspects DRAWING_DATA.autoLWHTbl every 5 s; for every new or re-drawn
'     SITE row it writes the exact centroid (Worker-identical, computed
'     offline) into BLANK Latitude/Longitude cells and a descriptive
'     "Exact Area, Taluka, District" label into the BLANK Location cell.
'     No Alt+F8, no popups - a single status-bar line is the only feedback.
'
' IMPORT PROCEDURE (required - otherwise compile errors):
'   1. In the VBE, REMOVE any module named "modSolarEPCDrawnLocation".
'   2. REMOVE the old "modSolarEPCResource" module (export a backup first
'      if desired).
'   3. Right-click the project -> Import File... -> this .bas file.
'   4. Save and reopen the workbook.
'
' SAFETY RULES OF THE AUTO-FILL:
'   - only BLANK cells are written; manual entries, formulas and NASA
'     values are never overwritten
'   - NASA numbers, Data Status, cache_key and polygon_hash are untouched
'   - if no label can be resolved offline/online the lat/lon are still
'     filled; label retries are limited and no incorrect value is ever
'     written
'   - stop anytime: SolarEPC_DrawnLocationAutoStop ; immediate sweep:
'     SolarEPC_DrawnLocationAutoNow ; manual inspection:
'     SolarEPC_ResourceShowDrawnLocation
'
' CENTROID MATH = WORKER MATH (port of polygonCentroidLocalProjection):
' verified byte-for-byte on 5000 generated polygons + 7 edge cases
' (location-extract/verify/run_parity_check.sh).
'=========================================================================="""
rep(old_header, new_header)

PAIRS = [
("'Watcher tuning (baaki sab constants module ke andar pehle se hain).",
 "'Watcher tuning (all other constants already exist inside the module)."),
("Private Const AUTO_NOROW_RETRIES As Long = 12    'RESOURCE_DB row na mile to ~1 min retry",
 "Private Const AUTO_NOROW_RETRIES As Long = 12    'retry ~1 min when the RESOURCE_DB row is missing"),
("Private Const AUTO_LABEL_RETRIES As Long = 6     'label offline ho to limited retry",
 "Private Const AUTO_LABEL_RETRIES As Long = 6     'limited retries while the label tier is offline"),
("' v2.3 - AUTOMATIC LOCATION FILL (koi Alt+F8 nahi, koi popup nahi)\n'\n' Workbook khulte hi (Auto_Open) ye watcher chalu ho jata hai aur har\n' AUTO_TICK_SECONDS me DRAWING_DATA.autoLWHTbl dekhta hai. Nayi ya badli hui\n' SITE row milte hi:\n'   1. exact area-weighted centroid offline calculate hota hai\n'   2. RESOURCE_DB me us Project ID ki row dhundhi jaati hai\n'   3. BLANK Latitude (deg) / Longitude (deg) cells me exact centroid likha\n'      jaata hai (Worker wale centroid se identical)\n'   4. BLANK Location cell me descriptive label likha jaata hai\n'      (\"Exact Area, Taluka, District\") - VILLAGE_DB -> Google -> Nominatim\n'      -> BigDataCloud, bilkul isi workbook ke existing rules se\n'\n' Kabhi bhi koi bhara hua cell overwrite NAHI hota, formula cells chhod diye\n' jaate hain, aur RESOURCE_DB me NASA ke numbers/status ko ye code chhoota\n' hi nahi. Feedback sirf status bar pe ek line me aata hai.",
 "' v2.3 - AUTOMATIC LOCATION FILL (no Alt+F8, no popups)\n'\n' The watcher starts when the workbook opens (Auto_Open) and inspects\n' DRAWING_DATA.autoLWHTbl every AUTO_TICK_SECONDS. As soon as a new or\n' modified SITE row appears:\n'   1. the exact area-weighted centroid is computed offline\n'   2. the RESOURCE_DB row for that Project ID is located\n'   3. the exact centroid (identical to the Worker's) is written into the\n'      BLANK Latitude / Longitude cells\n'   4. a descriptive label (\"Exact Area, Taluka, District\") is written into\n'      the BLANK Location cell - VILLAGE_DB -> Google -> Nominatim ->\n'      BigDataCloud, using this workbook's existing rules\n'\n' Filled cells are NEVER overwritten, formula cells are skipped, and the\n' NASA numbers/status inside RESOURCE_DB are never touched. The only\n' feedback is a single status-bar line."),
("'Watcher chalu karo + turant ek sweep (already-drawn sites bhi bhar jaayen).",
 "'Start the watcher and run one immediate sweep (already-drawn sites are filled too)."),
("    Application.StatusBar = \"Solar EPC: drawn-site location AUTO-fill ON (har \" & _\n        CStr(AUTO_TICK_SECONDS) & \"s DRAWING_DATA dekhta hai).\"",
 "    Application.StatusBar = \"Solar EPC: drawn-site location AUTO-fill ON (scanning DRAWING_DATA every \" & _\n        CStr(AUTO_TICK_SECONDS) & \" s).\""),
("'Watcher band karo (OnTime cancel ke saath).",
 "'Stop the watcher (cancels the scheduled OnTime)."),
("'Bina wait kiye turant ek sweep - kisi button pe laga sakte ho, zaroori nahi.",
 "'Run one sweep immediately without waiting - may be attached to a button."),
("'Internal scheduler entry point (OnTime ko Public naam chahiye).",
 "'Internal scheduler entry point (OnTime requires a Public name)."),
("'Ek baar saari SITE rows dekho; nayi/badli hui drawing pe auto-fill karo.",
 "'Inspect all SITE rows once; auto-fill any new or re-drawn site."),
("                KeyText = ReferenceText & \"|\" & CoordinateText",
 "                KeyText = ReferenceText & \"|\" & CoordinateText"),
("'Key me coordinates bhi hain: site dobara draw hui to key badal\n                                'jaati hai aur fill dobara attempt hota hai.",
 "'The key includes the coordinates: re-drawing the site changes the key\n                                'and the fill is attempted again."),
("                                        \" ke lat/lon bhar gaye; Location label offline hai, retry \" & _",
 "                                        \" lat/lon filled; Location label offline, retry \" & _"),
("                                        \" (internet / VILLAGE_DB / geocoding key dekho)\"",
 "                                        \" (check internet / VILLAGE_DB / geocoding key)\""),
("                                    Application.StatusBar = \"Solar EPC: RESOURCE_DB me \" & ProjectID & _\n                                        \" ki row nahi mili (\" & CStr(Attempts + 1) & \") - \" & _\n                                        IIf(Attempts + 1 >= 2, \"ab row bana ke bhar raha hoon...\", \"retry...\")",
 "                                    Application.StatusBar = \"Solar EPC: no RESOURCE_DB row for \" & ProjectID & _\n                                        \" (\" & CStr(Attempts + 1) & \") - \" & _\n                                        IIf(Attempts + 1 >= 2, \"creating the row now...\", \"retrying...\")"),
("'\"NOROW:3\" / \"LOCPEND:1\" jaisi state me se attempt count nikaalo.",
 "'Extract the attempt count from states such as \"NOROW:3\" / \"LOCPEND:1\"."),
("'Hidden _CLOUD_CFG me last auto-fill ka record (diagnostic only; sheet na ho\n'to chupchaap skip).",
 "'Record of the last auto-fill in the hidden _CLOUD_CFG sheet\n'(diagnostic only; silently skipped when the sheet is absent)."),
("'RESOURCE_DB me us Project ID ki row dhundh ke BLANK lat/lon + BLANK Location\n'bharo. Return: kya bhara (\"lat lon location(Sonwadi Bk., Phaltan, Satara)\"),\n'ya SKIP-NO-BLANK / NOROW / LOCPEND...",
 "'Locate the RESOURCE_DB row for this Project ID and fill BLANK lat/lon +\n'BLANK Location. Returns what was filled, e.g.\n'\"lat lon location(Sonwadi Bk., Phaltan, Satara)\", or SKIP-NO-BLANK /\n'NOROW / LOCPEND."),
("'v2.5: row hai hi nahi to bana do - draw kiya hai to location RESOURCE_DB\n    'me dikhni chahiye. Baaki columns NASA import / manual bharenge.",
 "'v2.5: create the missing row - a drawn site must appear in RESOURCE_DB.\n    'Remaining columns are filled by the NASA import or manually."),
("'resource_db Excel Table (naam se), warna Nothing.",
 "'The resource_db Excel Table by name, or Nothing."),
("'Column index by header name, 0 agar column nahi hai.",
 "'Column index by header name; 0 when the column does not exist."),
("' v2.4 adapters - watcher ke naampari helpers ko complete module ke proven\n' helpers se jodte hain (koi logic duplicate nahi).",
 " v2.4 adapters - connect the watcher's named helpers to the proven\n' helpers of the complete module (no duplicated logic)."),
("' v2.5 - DIAGNOSTIC: chain kaha atki hai, ek click me batata hai.\n'   Alt+F8 -> SolarEPC_DrawnLocationDebug\n' Kuch bhi change NAHI karta (sirf padhta hai + report dikhata hai).",
 "' v2.5 - DIAGNOSTIC: reports where the chain is stuck, in one click.\n'   Alt+F8 -> SolarEPC_DrawnLocationDebug\n' Changes NOTHING (read-only report)."),
("        P1 = P1 & \"   >> Watcher OFF hai! SolarEPC_DrawnLocationAutoStart chalao\" & vbCrLf & _\n                  \"      ya workbook save karke DOBARA kholo (Auto_Open se chalu hota hai).\" & vbCrLf",
 "        P1 = P1 & \"   >> Watcher is OFF. Run SolarEPC_DrawnLocationAutoStart\" & vbCrLf & _\n                  \"      or save and REOPEN the workbook (Auto_Open starts it).\" & vbCrLf"),
("        P1 = P1 & \"2) DRAWING_DATA / '\" & SITE_TABLE & \"' table NAHIN MILA.\" & vbCrLf & _\n                  \"   >> Sheet/table ka naam check karo.\" & vbCrLf",
 "        P1 = P1 & \"2) DRAWING_DATA / '\" & SITE_TABLE & \"' table NOT FOUND.\" & vbCrLf & _\n                  \"   >> Check the sheet/table name.\" & vbCrLf"),
("            P1 = P1 & \"   >> Koi SITE-MAP-* row nahi mili - map me SAVE hua tha kya?\" & vbCrLf",
 "            P1 = P1 & \"   >> No SITE-MAP-* row found - was the drawing SAVED on the map?\" & vbCrLf"),
("                P1 = P1 & \"   >> Coordinates cell BLANK hai - map SAVE adhoora raha.\" & vbCrLf",
 "                P1 = P1 & \"   >> Coordinates cell is BLANK - the map SAVE was incomplete.\" & vbCrLf"),
]
for o, n in PAIRS:
    rep(o, n)

io.open(ROOT + 'modSolarEPCResource.bas', 'w', encoding='utf-8').write(s)
print('pass 1 done; misses:', len(miss))
for m in miss:
    print('MISS:', m)
