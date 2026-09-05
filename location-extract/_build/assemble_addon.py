import io, re

ROOT = '/home/user/Solar-Designed-Map/location-extract/'
complete = io.open(ROOT + 'modSolarEPCResource_v2.2_complete.bas', encoding='utf-8').read().split('\n')
watcher = io.open(ROOT + '_build/watcher_block.bas', encoding='utf-8').read()
adapters = io.open(ROOT + '_build/adapters_block.bas', encoding='utf-8').read()
debug = io.open(ROOT + '_build/debug_block.bas', encoding='utf-8').read()

def slice_proc(name):
    start = None
    for i, ln in enumerate(complete):
        if re.match(r'^(Private|Public) (Function|Sub) ' + re.escape(name) + r'\(', ln):
            start = i
            break
    if start is None:
        raise SystemExit('MISSING ' + name)
    for j in range(start + 1, len(complete)):
        if complete[j].strip() in ('End Function', 'End Sub'):
            return '\n'.join(complete[start:j + 1])
    raise SystemExit('UNTERMINATED ' + name)

SLICE = ["ResourceColumn", "ResourceSiteTable", "ResourceRowCellText",
         "ResourceStrictNumber", "ResourceDecimal", "ResourceParseSitePolygon",
         "ResourceCentroidLocalProjection", "ResourceHttpGetJson", "ResourceJSONValue",
         "ResourceLocationCompose", "ResourceSameName", "ResourceAdminClean",
         "ResourceLocalityName", "ResourceAdminNames", "ResourceDistanceKm",
         "ResourceVillageDbLabel", "ResourceCopyToClipboard", "ResourceLocationLabel"]

body = [slice_proc(n) for n in SLICE]
label = body[-1]
google_lines = """    '0b) OPTIONAL Google tier - only fires when a key is found and works.
    ResourceLocationLabel = ResourceGoogleLocationLabel(Lat, Lon)
    If Len(ResourceLocationLabel) > 0 Then GoTo StoreCache
"""
assert google_lines in label + '\n' or google_lines in '\n'.join(body[-1:])
body[-1] = label.replace(google_lines, '')

# watcher publics -> add-on names (Auto_Open/Auto_Close magic names stay)
for old, new in [("SolarEPC_DrawnLocationAutoStart", "SolarEPC_AutoLocationStart"),
                 ("SolarEPC_DrawnLocationAutoStop", "SolarEPC_AutoLocationStop"),
                 ("SolarEPC_DrawnLocationAutoNow", "SolarEPC_AutoLocationNow"),
                 ("SolarEPC_DrawnLocationTick", "SolarEPC_AutoLocationTick")]:
    watcher = watcher.replace(old, new)
debug = debug.replace("Public Sub SolarEPC_DrawnLocationDebug()",
                      "Public Sub SolarEPC_AutoLocationDebug()")

header = """Attribute VB_Name = "modSolarEPCAutoLocation"
Option Explicit

'==========================================================================
' SOLAR EPC - AUTOMATIC DRAWN-LOCATION ADD-ON  (v2.6 add-on)
'
' PURPOSE
'   Sit-ke saath chalne wala chhota add-on: map pe draw + SAVE hote hi
'   RESOURCE_DB ke BLANK Latitude/Longitude cells me exact centroid (8 dp)
'   aur BLANK Location cell me "Exact Area, Taluka, District" apne aap
'   bhar deta hai - koi Alt+F8, koi popup nahi. Row na ho to ~10s baad
'   row bhi apne aap bana deta hai.
'
' KISKE SAATH CHALTA HAI
'   - Legacy modSolarEPCResource v2.1 / v2.2 / v2.3 ke SAATH (replace nahi
'     karta). Saare helpers Private hain, isliye naam takrate nahi.
'   - v2.4+ complete module ke saath MAT lagana (wo watcher pehle se hai).
'
' LABEL TIERS (free/offline first)
'   0) VILLAGE_DB sheet (aapka apna gaon list, offline, exact revenue name)
'   1) OSM Nominatim (free, no key)
'   2) BigDataCloud (free, no key)
'   Google tier is add-on me nahi hai - legacy module wo khud chalata hai.
'
' PUBLIC MACROS
'   SolarEPC_AutoLocationStart / Stop / Now / Tick / Debug
'   Auto_Open / Auto_Close  (workbook khulte/band hote watcher ON/OFF)
'==========================================================================

Private Const CONFIG_SHEET As String = "_CLOUD_CFG"
Private Const DRAWING_SHEET As String = "DRAWING_DATA"
Private Const SITE_TABLE As String = "autoLWHTbl"
Private Const RESOURCE_DB_TABLE As String = "resource_db"
Private Const VILLAGE_SHEET As String = "VILLAGE_DB"
Private Const VILLAGE_MAX_KM As Double = 5#
Private Const SITE_MAX_VERTICES As Long = 4000
Private Const AUTO_TICK_SECONDS As Long = 5
Private Const AUTO_NOROW_RETRIES As Long = 12
Private Const AUTO_LABEL_RETRIES As Long = 6

Private mAutoActive As Boolean
Private mAutoNextRun As Date
Private mProcessed As Object
Private mLastLocKey As String
Private mLastLocLabel As String
"""

out = header + '\n' + watcher + '\n' + adapters + '\n' + '\n\n'.join(body) + '\n' + debug + '\n'
io.open(ROOT + 'modSolarEPCAutoLocation.bas', 'w', encoding='utf-8').write(out)
pubs = re.findall(r'^Public (?:Sub|Function) ([A-Za-z0-9_]+)', out, re.M)
print('lines:', out.count('\n') + 1, '| publics:', pubs)
