import io

ROOT = '/home/user/Solar-Designed-Map/location-extract/'
s = io.open(ROOT + 'modSolarEPCResource.bas', encoding='utf-8').read()

# 1) stamp sub -> blank-only + candidate column finder
old_stamp = """Private Sub ResourceStampLocalRunTime(ByVal Tbl As ListObject, ByVal Target As Range)
    Dim C As Long
    Dim Col As ListColumn
    Dim Cell As Range

    On Error Resume Next
    C = ResourceColumn(Tbl, "Date/Time")              'user ka apna column
    If C = 0 Then C = ResourceColumn(Tbl, "Run Date-Time (Local)")
    If C = 0 Then
        Set Col = Tbl.ListColumns.Add(Tbl.ListColumns.Count + 1)
        If Not Col Is Nothing Then Col.Name = "Run Date-Time (Local)"
        C = ResourceColumn(Tbl, "Run Date-Time (Local)")
    End If
    On Error GoTo 0
    If C = 0 Then Exit Sub
    Set Cell = Target.Cells(1, C)
    If Cell Is Nothing Then Exit Sub
    If Cell.HasFormula Then Exit Sub
    Cell.NumberFormat = "dd-mm-yyyy hh:nn:ss"
    Cell.Value2 = Now                                  'real date value, sortable
End Sub
"""
new_stamp = """Private Function ResourceDateTimeColumn(ByVal Tbl As ListObject) As Long
    Dim Names As Variant
    Dim i As Long
    Names = Array("Date/Time", "Run Date-Time (Local)", "Date-Time", _
                  "DateTime", "Date / Time")
    For i = LBound(Names) To UBound(Names)
        ResourceDateTimeColumn = ResourceColumn(Tbl, CStr(Names(i)))
        If ResourceDateTimeColumn > 0 Then Exit Function
    Next i
End Function

'v3.5: BLANK-ONLY local run stamp. Jis event ne row pe pehla data likha
'(watcher auto-fill ya NASA import), wahi din+samay stamp hota hai.
'Kabhi overwrite nahi karta - manual date entry sacred hai.
Private Sub ResourceStampLocalRunTime(ByVal Tbl As ListObject, ByVal Target As Range)
    Dim C As Long
    Dim Col As ListColumn
    Dim Cell As Range

    On Error Resume Next
    C = ResourceDateTimeColumn(Tbl)
    If C = 0 Then
        Set Col = Tbl.ListColumns.Add(Tbl.ListColumns.Count + 1)
        If Not Col Is Nothing Then Col.Name = "Run Date-Time (Local)"
        C = ResourceDateTimeColumn(Tbl)
    End If
    On Error GoTo 0
    If C = 0 Then Exit Sub
    Set Cell = Target.Cells(1, C)
    If Cell Is Nothing Then Exit Sub
    If Cell.HasFormula Then Exit Sub
    If Len(Trim$(CStr(Cell.Value2))) > 0 Then Exit Sub     'blank-only, hamesha
    Cell.NumberFormat = "dd-mm-yyyy hh:nn:ss"
    Cell.Value2 = Now                                      'real date value
End Sub
"""
assert s.count(old_stamp) == 1
s = s.replace(old_stamp, new_stamp, 1)

# 2) watcher ke auto-fill pe bhi stamp (pehle sirf NASA import pe tha)
old_skip = '    If Len(DidText) = 0 Then DidText = "SKIP-NO-BLANK"\n'
assert s.count(old_skip) == 1
s = s.replace(old_skip,
    '    If Len(DidText) > 0 And DidText <> "SKIP-NO-BLANK" Then _\n        ResourceStampLocalRunTime Tbl, Target\n' + old_skip, 1)

# 3) debug macro me Date/Time column report
old_dims = '    Dim cProject As Long, cLat As Long, cLon As Long, cLoc As Long\n'
assert s.count(old_dims) == 1
s = s.replace(old_dims, old_dims + '    Dim cDT As Long\n', 1)
old_cols = '        cLoc = TableColumn(Rdb, "Location")\n'
assert s.count(old_cols) == 1
s = s.replace(old_cols, old_cols + '        cDT = ResourceDateTimeColumn(Rdb)\n', 1)
old_p2cols = '             IIf(cLon > 0, "Lon", "-") & IIf(cLoc > 0, "Loc", "-") & vbCrLf\n'
assert s.count(old_p2cols) == 1
s = s.replace(old_p2cols, old_p2cols +
    '        P2 = P2 & "   Date/Time column: " & IIf(cDT > 0, _\n             "mil gayi (" & Rdb.ListColumns(cDT).Name & ")", _\n             "TABLE KE ANDAR NAHIN - header row ke andar add karo") & vbCrLf\n', 1)

# 4) version 3.5
assert "' Version 3.4" in s
s = s.replace("' Version 3.4", "' Version 3.5", 1)
hist = "'   - v3.4 LEAN: GeoNames tier HATAYA"
assert hist in s
s = s.replace(hist, "'   - v3.5: Date/Time stamp ab BLANK-ONLY hai aur HAR pehle fill pe\n"
    "'     lagta hai (watcher auto-fill YA NASA import, jo pehle ho). Column\n"
    "'     candidates: Date/Time, Run Date-Time (Local), Date-Time, DateTime.\n"
    "'     Debug macro batata hai column table ke andar mili ya nahi.\n"
    + hist, 1)

io.open(ROOT + 'modSolarEPCResource.bas', 'w', encoding='utf-8').write(s)
print('v3.5 assembled:', s.count('\n') + 1, 'lines')
