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

PAIRS = [
# leftover key comment
("""                'Key me coordinates bhi hain: site dobara draw hui to key badal
                                'jaati hai aur fill dobara attempt hota hai.""",
 """                'The key includes the coordinates: re-drawing the site changes
                                'the key and the fill is attempted again."""),
# debug part 2 strings
("""        P2 = "4) RESOURCE_DB ka 'resource_db' table NAHIN MILA." & vbCrLf & _
             "   >> Table ka naam exactly 'resource_db' hona chahiye." & vbCrLf""",
 """        P2 = "4) The 'resource_db' table was NOT FOUND in RESOURCE_DB." & vbCrLf & _
             "   >> The table name must be exactly 'resource_db'." & vbCrLf"""),
("""            P2 = P2 & "5) >> " & ProjectID & " ki koi BLANK-lat/lon row nahi mili." & vbCrLf & _
                          "   (v2.5 ab aisi halat me row APNE AAP bana deta hai - watcher ON ho to)" & vbCrLf""",
 """            P2 = P2 & "5) >> No BLANK-lat/lon row found for " & ProjectID & "." & vbCrLf & _
                          "   (v2.5+ creates such a row automatically while the watcher is ON)" & vbCrLf"""),
("""        LabelText = ResourceLocationLabel(Decimal8(CentroidLatitude), Decimal8(CentroidLongitude))
        P2 = P2 & "6) Label test: " & IIf(Len(LabelText) > 0, LabelText, _
             "KOI TIER NAHIN MILA (VILLAGE_DB sheet / internet / geocoding key)") & vbCrLf""",
 """        LabelText = ResourceLocationLabel(Decimal8(CentroidLatitude), Decimal8(CentroidLongitude))
        P2 = P2 & "6) Label test: " & IIf(Len(LabelText) > 0, LabelText, _
             "NO TIER RESOLVED (VILLAGE_DB sheet / internet / geocoding key)") & vbCrLf"""),
("""        P2 = P2 & "   Date/Time column: " & IIf(cDT > 0, _
             "mil gayi (" & Rdb.ListColumns(cDT).Name & ")", _
             "TABLE KE ANDAR NAHIN - header row ke andar add karo") & vbCrLf""",
 """        P2 = P2 & "   Date/Time column: " & IIf(cDT > 0, _
             "found (" & Rdb.ListColumns(cDT).Name & ")", _
             "NOT INSIDE THE TABLE - add it inside the header row") & vbCrLf"""),
("""    P2 = P2 & vbCrLf & "Ye report copy ho gayi hai clipboard pe." """,
 """    P2 = P2 & vbCrLf & "This report has been copied to the clipboard." """),
# v3 block comments
("""'v3.5: BLANK-ONLY local run stamp. Jis event ne row pe pehla data likha
'(watcher auto-fill ya NASA import), wahi din+samay stamp hota hai.
'Kabhi overwrite nahi karta - manual date entry sacred hai.""",
 """'v3.5: BLANK-ONLY local run stamp. Whichever event writes the first data
'into a row (watcher auto-fill or NASA import) stamps the local date and
'time. Never overwrites - a manual date entry is sacred."""),
("""'v3.10: NumberFormat kabhi-kabhi blocked hota hai (sheet protection / table
'column setting). Isliye pehle try karo, fail ho to text likh do - value
'hamesha pahunchti hai, format ki bheek nahi maangte.""",
 """'v3.10: NumberFormat can be blocked (sheet protection / table column
'setting). Try the format first; if Excel refuses, write clean text -
'the VALUE always lands, we never depend on the format."""),
("""' v3.8: THE BIG RED BUTTON. Ek click, sab kaam, usi waqt - koi session
' memory nahi, koi reopen nahi, koi wait nahi:
'   1. watcher ON na ho to ON karta hai
'   2. sweep TURANT chalata hai (blank lat/lon + Location + Date/Time stamp)
'   3. purani bhari hui rows jinka Date/Time khaali hai, unhe NASA ke
'      "Retrieval Date" (UTC) se tumhare LOCAL time me convert karke bhar
'      deta hai (sach waala time, jhootha Now nahi)
'   4. B21 + MsgBox me poori report deta hai""",
 """' v3.8: THE BIG RED BUTTON. One click performs everything immediately -
' no session memory, no reopen, no waiting:
'   1. starts the watcher if it is OFF
'   2. runs an immediate sweep (blank lat/lon + Location + Date/Time stamp)
'   3. back-fills Date/Time on previously filled rows by converting the
'      NASA "Retrieval Date" (UTC) into the user's LOCAL time (the true
'      run time, not a fake Now)
'   4. reports everything in _CLOUD_CFG B21 and a MsgBox"""),
("""' v3.2: PURANI bhari hui Location rows ko safely upgrade karo jab ek better
' tier (VILLAGE_DB / GeoNames) available ho jaye. Blank-only rule manual
' cells ke liye sacred hai, isliye ye macro sirf USER-ON-DEMAND chalta hai
' aur sirf tab likhta hai jab naya label purane label ka "village-plus"
' version ho (naya label ", " & purana label pe khatam hota ho).
'   "Phaltan, Satara" -> "Sonwadi Bk., Phaltan, Satara"   = upgrade OK
'   "My Farm, Phaltan" (manual)                          = skip, kabhi nahi chhuta""",
 """' v3.2: safely upgrade previously filled Location rows when a better tier
' (VILLAGE_DB) becomes available. The blank-only rule is sacred for manual
' cells, so this macro runs ONLY on user demand and writes only when the
' new label is exactly the "village-plus" version of the old one (the new
' label ends with ", " & old label).
'   "Phaltan, Satara" -> "Sonwadi Bk., Phaltan, Satara"   = upgrade OK
'   "My Farm, Phaltan" (manual)                          = skipped, never touched"""),
("""            Application.StatusBar = "Solar EPC: label refresh " & CStr(R) & _
                "/" & CStr(Total) & " (manual entries kabhi overwrite nahi hote)..." """,
 """            Application.StatusBar = "Solar EPC: label refresh " & CStr(R) & _
                "/" & CStr(Total) & " (manual entries are never overwritten)..." """),
("""        "Sirf wo rows badli hain jaha naya label purane label ka exact " & _
        "'village + ' version tha. Manual entries aur baaki sab untouched.", _""",
 """        "Only rows were changed where the new label is the exact " & _
        "'village + ' version of the old one. Manual entries and everything else untouched.", _"""),
("""        MsgBox "The RESOURCE_DB table was not found.", vbExclamation, "Solar EPC" """,
 """        MsgBox "The RESOURCE_DB table was not found.", vbExclamation, "Solar EPC" """),
("""    If Tbl Is Nothing Then
        MsgBox "resource_db table nahi mila.", vbExclamation, "Solar EPC"
        Exit Sub
    End If""",
 """    If Tbl Is Nothing Then
        MsgBox "The resource_db table was not found.", vbExclamation, "Solar EPC"
        Exit Sub
    End If"""),
("""    If cDT = 0 Or cLat = 0 Or cLon = 0 Then
        MsgBox "Date/Time ya Latitude/Longitude columns nahi mili.", _
            vbExclamation, "Solar EPC"
        Exit Sub
    End If""",
 """    If cDT = 0 Or cLat = 0 Or cLon = 0 Then
        MsgBox "The Date/Time or Latitude/Longitude columns were not found.", _
            vbExclamation, "Solar EPC"
        Exit Sub
    End If"""),
("""        "Blank chhodi (lat/lon ya Retrieval Date missing): " & CStr(LeftBlank) & vbCrLf & vbCrLf & _
        "Watcher ON hai - nayi draws ab apne aap bhar + stamp hongi.", _""",
 """        "Left blank (lat/lon or Retrieval Date missing): " & CStr(LeftBlank) & vbCrLf & vbCrLf & _
        "The watcher is ON - new drawings will fill and stamp automatically.", _"""),
("' Answers: \"jis jagah maine map par draw kiya, uski EXACT location kya hai?\"",
 "' Answers: \"what is the EXACT location of the place I drew on the map?\""),
]
for o, n in PAIRS:
    rep(o, n)

# ---------------- SPEED FIX: bulk-read filled map instead of per-key scans ----------------
old_fn_start = "'v3.7: watcher ki session-memory self-heal: agar koi site DONE mark hai par"
i = s.index(old_fn_start)
j = s.index('End Function', i) + len('End Function')
old_fn = s[i:j]
new_fn = """'v3.11: watcher session-memory self-heal, FAST edition. One bulk column read
'per sweep builds a ProjectID -> filled map; a DONE key is re-processed only
'when its row's lat/lon are blank again (user cleared them or an earlier
'session left them empty). No per-cell COM calls, no reopen needed.
Private Function ResourceBuildFilledMap() As Object
    Dim D As Object
    Dim Rdb As ListObject
    Dim cP As Long, cLat As Long, cLon As Long
    Dim ProjArr As Variant, LatArr As Variant, LonArr As Variant
    Dim n As Long, i As Long

    Set D = CreateObject("Scripting.Dictionary")
    D.CompareMode = 1                                  'vbTextCompare
    On Error GoTo Failed
    Set Rdb = ResourceDbTable()
    If Rdb Is Nothing Then GoTo DoneMap
    cP = TableColumn(Rdb, "Project ID")
    cLat = TableColumn(Rdb, "Latitude (" & ChrW(176) & ")")
    cLon = TableColumn(Rdb, "Longitude (" & ChrW(176) & ")")
    If cP = 0 Or cLat = 0 Or cLon = 0 Then GoTo DoneMap
    n = Rdb.ListRows.Count
    If n = 0 Then GoTo DoneMap
    ProjArr = Rdb.ListColumns(cP).Range.Value2         'one bulk read (incl. header)
    LatArr = Rdb.ListColumns(cLat).Range.Value2
    LonArr = Rdb.ListColumns(cLon).Range.Value2
    For i = 2 To n + 1
        If Len(Trim$(CStr(ProjArr(i, 1)))) > 0 Then
            If Len(Trim$(CStr(LatArr(i, 1)))) > 0 And _
               Len(Trim$(CStr(LonArr(i, 1)))) > 0 Then
                D(CStr(ProjArr(i, 1))) = True
            End If
        End If
    Next i
DoneMap:
    Set ResourceBuildFilledMap = D
    Exit Function
Failed:
    Set ResourceBuildFilledMap = D
End Function

Private Function ProjectFilledInCache(ByVal ProjectID As String) As Boolean
    If mFilledCache Is Nothing Then Exit Function
    ProjectFilledInCache = mFilledCache.Exists(ProjectID)
End Function"""
s = s.replace(old_fn, new_fn, 1)

rep("""    If mProcessed Is Nothing Then Set mProcessed = CreateObject("Scripting.Dictionary")

    For R = 1 To Tbl.ListRows.Count
        ReferenceText = RowCellText(Tbl, R, "Reference id")""",
"""    If mProcessed Is Nothing Then Set mProcessed = CreateObject("Scripting.Dictionary")
    Set mFilledCache = ResourceBuildFilledMap()

    For R = 1 To Tbl.ListRows.Count
        ReferenceText = RowCellText(Tbl, R, "Reference id")""")

rep("                   Not ResourceProjectRowFilled(ProjectID) Then",
    "                   Not ProjectFilledInCache(ProjectID) Then")

rep("Private mProcessed As Object          'key = ReferenceID|Coordinates -> state",
    "Private mProcessed As Object          'key = ReferenceID|Coordinates -> state\nPrivate mFilledCache As Object          'ProjectID -> row already carries lat/lon")

# ---------------- version -> 3.11 FINAL ----------------
rep("' Version 3.10", "' Version 3.11 (FINAL)")
rep("'   - v3.8: SolarEPC_FillAndStampNow - one click starts the watcher, runs an",
    "'   - v3.11 FINAL: every user-facing message, status-bar line and debug\n"
    "'     report is now professional English; the per-tick RESOURCE_DB scan\n"
    "'     was replaced by one bulk column read per sweep (faster ticks);\n"
    "'     K=12 range mode verified active (_CLOUD_CFG B12 = RANGE12).\n"
    "'   - v3.8: SolarEPC_FillAndStampNow - one click starts the watcher, runs an")

io.open(ROOT + 'modSolarEPCResource.bas', 'w', encoding='utf-8').write(s)
print('pass 2 done; misses:', len(miss))
for m in miss:
    print('MISS:', m)
