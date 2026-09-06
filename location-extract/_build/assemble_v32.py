import io

ROOT = '/home/user/Solar-Designed-Map/location-extract/'
s = io.open(ROOT + 'modSolarEPCResource.bas', encoding='utf-8').read()

refresh = '''
'--------------------------------------------------------------------------
' v3.2: PURANI bhari hui Location rows ko safely upgrade karo jab ek better
' tier (VILLAGE_DB / GeoNames) available ho jaye. Blank-only rule manual
' cells ke liye sacred hai, isliye ye macro sirf USER-ON-DEMAND chalta hai
' aur sirf tab likhta hai jab naya label purane label ka "village-plus"
' version ho (naya label ", " & purana label pe khatam hota ho).
'   "Phaltan, Satara" -> "Sonwadi Bk., Phaltan, Satara"   = upgrade OK
'   "My Farm, Phaltan" (manual)                          = skip, kabhi nahi chhuta
'--------------------------------------------------------------------------
Public Sub SolarEPC_ResourceRefreshLabels()
    Dim ws As Worksheet
    Dim Tbl As ListObject
    Dim R As Long
    Dim cLoc As Long, cLat As Long, cLon As Long
    Dim Cell As Range
    Dim Current As String, NewLabel As String
    Dim LatText As String, LonText As String
    Dim Checked As Long, Upgraded As Long
    Dim Total As Long

    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(RESOURCE_SHEET)
    Set Tbl = ResourceFindHeaderTable(ws)
    If Tbl Is Nothing Then
        MsgBox "The RESOURCE_DB table was not found.", vbExclamation, "Solar EPC Resource"
        Exit Sub
    End If
    cLoc = ResourceColumn(Tbl, "Location")
    cLat = ResourceColumn(Tbl, "Latitude (" & ChrW(176) & ")")
    cLon = ResourceColumn(Tbl, "Longitude (" & ChrW(176) & ")")
    If cLoc = 0 Or cLat = 0 Or cLon = 0 Then
        MsgBox "RESOURCE_DB requires the Location, Latitude and Longitude columns.", _
            vbExclamation, "Solar EPC Resource"
        Exit Sub
    End If

    Total = Tbl.ListRows.Count
    For R = 1 To Total
        Set Cell = Tbl.ListRows(R).Range.Cells(1, cLoc)
        Current = Trim$(CStr(Cell.Value2))
        LatText = Trim$(CStr(Tbl.ListRows(R).Range.Cells(1, cLat).Value2))
        LonText = Trim$(CStr(Tbl.ListRows(R).Range.Cells(1, cLon).Value2))
        If Len(Current) > 0 And Len(LatText) > 0 And Len(LonText) > 0 And _
           Not Cell.HasFormula Then
            Checked = Checked + 1
            Application.StatusBar = "Solar EPC: label refresh " & CStr(R) & _
                "/" & CStr(Total) & " (manual entries kabhi overwrite nahi hote)..."
            DoEvents
            mLastLocKey = vbNullString
            mLastLocLabel = vbNullString
            NewLabel = ResourceLocationLabel(LatText, LonText)
            If Len(NewLabel) > 0 Then
                If StrComp(NewLabel, Current, vbTextCompare) <> 0 And _
                   Len(NewLabel) > Len(Current) + 2 And _
                   Right$(NewLabel, Len(Current) + 2) = ", " & Current Then
                    Cell.Value2 = NewLabel
                    Upgraded = Upgraded + 1
                End If
            End If
            'OSM/GeoNames politeness: ~1 request/second.
            If R < Total Then Application.Wait Now + TimeSerial(0, 0, 1)
        End If
    Next R

    Application.StatusBar = False
    MsgBox "Label refresh complete." & vbCrLf & vbCrLf & _
        "Rows checked : " & CStr(Checked) & vbCrLf & _
        "Village-level upgrades : " & CStr(Upgraded) & vbCrLf & vbCrLf & _
        "Sirf wo rows badli hain jaha naya label purane label ka exact " & _
        "'village + ' version tha. Manual entries aur baaki sab untouched.", _
        vbInformation, "Solar EPC Resource"
    Exit Sub
Failed:
    Application.StatusBar = False
    MsgBox "Label refresh failed: " & Err.Description, vbExclamation, "Solar EPC Resource"
End Sub
'''
s = s.rstrip() + '\n' + refresh

assert "' Version 3.1" in s
s = s.replace("' Version 3.1", "' Version 3.2", 1)
hist = "'   - v3.1: GeoNames FREE tier"
assert hist in s
s = s.replace(hist, "'   - v3.2: SolarEPC_ResourceRefreshLabels - purani bhari hui Location\n"
    "'     rows ko village-level label pe safely upgrade karo (suffix-rule,\n"
    "'     manual entries kabhi nahi chhute). Blank-only auto-fill waisa hi.\n"
    + hist, 1)

io.open(ROOT + 'modSolarEPCResource.bas', 'w', encoding='utf-8').write(s)
print('v3.2 assembled:', s.count('\n') + 1, 'lines')
