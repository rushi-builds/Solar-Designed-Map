import io

ROOT = '/home/user/Solar-Designed-Map/location-extract/'
s = io.open(ROOT + 'modSolarEPCResource.bas', encoding='utf-8').read()

block = '''
'--------------------------------------------------------------------------
' v3.8: THE BIG RED BUTTON. Ek click, sab kaam, usi waqt - koi session
' memory nahi, koi reopen nahi, koi wait nahi:
'   1. watcher ON na ho to ON karta hai
'   2. sweep TURANT chalata hai (blank lat/lon + Location + Date/Time stamp)
'   3. purani bhari hui rows jinka Date/Time khaali hai, unhe NASA ke
'      "Retrieval Date" (UTC) se tumhare LOCAL time me convert karke bhar
'      deta hai (sach waala time, jhootha Now nahi)
'   4. B21 + MsgBox me poori report deta hai
'--------------------------------------------------------------------------
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

Private Function ResourceLocalUtcOffset() As Double
    Dim ST As SYSTEMTIME
    Dim UtcNow As Date
    On Error GoTo Fallback
    GetSystemTime ST
    UtcNow = DateSerial(ST.wYear, ST.wMonth, ST.wDay) + _
             TimeSerial(ST.wHour, ST.wMinute, ST.wSecond)
    ResourceLocalUtcOffset = Int((Now - UtcNow) * 48 + 0.5) / 48
    Exit Function
Fallback:
    ResourceLocalUtcOffset = 5.5 / 24      'IST fallback
End Function

Private Function ResourceTryUtcToLocal(ByVal TextValue As String, _
    ByRef OutTime As Date) As Boolean
    Dim Y As Long, M As Long, D As Long
    Dim h As Long, mi As Long, s As Long
    Dim UtcTime As Date
    On Error GoTo Failed
    If Len(TextValue) < 19 Then Exit Function
    If Mid$(TextValue, 5, 1) <> "-" Or Mid$(TextValue, 11, 1) <> "T" Then Exit Function
    Y = CLng(Left$(TextValue, 4)): M = CLng(Mid$(TextValue, 6, 2)): D = CLng(Mid$(TextValue, 9, 2))
    h = CLng(Mid$(TextValue, 12, 2)): mi = CLng(Mid$(TextValue, 15, 2)): s = CLng(Mid$(TextValue, 18, 2))
    UtcTime = DateSerial(Y, M, D) + TimeSerial(h, mi, s)
    OutTime = UtcTime + ResourceLocalUtcOffset()
    ResourceTryUtcToLocal = True
    Exit Function
Failed:
    ResourceTryUtcToLocal = False
End Function

Public Sub SolarEPC_FillAndStampNow()
    Dim Tbl As ListObject
    Dim R As Long
    Dim cDT As Long, cLat As Long, cLon As Long, cRet As Long
    Dim Col As ListColumn
    Dim Cell As Range
    Dim LatText As String, LonText As String, RetText As String
    Dim LocalTime As Date
    Dim FilledRows As Long, Stamped As Long, LeftBlank As Long

    On Error GoTo Failed
    If Not mAutoActive Then SolarEPC_DrawnLocationAutoStart
    DrawnLocationSweep

    Set Tbl = ResourceDbTable()
    If Tbl Is Nothing Then
        MsgBox "resource_db table nahi mila.", vbExclamation, "Solar EPC"
        Exit Sub
    End If
    cDT = ResourceDateTimeColumn(Tbl)
    If cDT = 0 Then
        On Error Resume Next
        Set Col = Tbl.ListColumns.Add(Tbl.ListColumns.Count + 1)
        If Not Col Is Nothing Then Col.Name = "Run Date-Time (Local)"
        On Error GoTo 0
        cDT = ResourceDateTimeColumn(Tbl)
    End If
    cLat = TableColumn(Tbl, "Latitude (" & ChrW(176) & ")")
    cLon = TableColumn(Tbl, "Longitude (" & ChrW(176) & ")")
    cRet = TableColumn(Tbl, "Retrieval Date")
    If cDT = 0 Or cLat = 0 Or cLon = 0 Then
        MsgBox "Date/Time ya Latitude/Longitude columns nahi mili.", _
            vbExclamation, "Solar EPC"
        Exit Sub
    End If

    For R = 1 To Tbl.ListRows.Count
        Set Cell = Tbl.ListRows(R).Range.Cells(1, cDT)
        If Len(Trim$(CStr(Cell.Value2))) = 0 And Not Cell.HasFormula Then
            LatText = Trim$(CStr(Tbl.ListRows(R).Range.Cells(1, cLat).Value2))
            LonText = Trim$(CStr(Tbl.ListRows(R).Range.Cells(1, cLon).Value2))
            If Len(LatText) > 0 And Len(LonText) > 0 Then
                If cRet > 0 Then
                    RetText = Trim$(CStr(Tbl.ListRows(R).Range.Cells(1, cRet).Value2))
                    If ResourceTryUtcToLocal(RetText, LocalTime) Then
                        Cell.NumberFormat = "dd-mm-yyyy hh:nn:ss"
                        Cell.Value2 = LocalTime
                        Stamped = Stamped + 1
                    Else
                        LeftBlank = LeftBlank + 1
                    End If
                Else
                    LeftBlank = LeftBlank + 1
                End If
            End If
        End If
    Next R

    ResourceSettingDiag "A21", "DATE/TIME STAMP", "B21", _
        "BUTTON " & Format$(Now, "dd-mm-yyyy hh:nn:ss") & ": backfilled " & _
        CStr(Stamped) & " row(s) from Retrieval Date (UTC->local), blank left " & _
        CStr(LeftBlank)
    MsgBox "Fill + Stamp NOW complete." & vbCrLf & vbCrLf & _
        "Date/Time backfilled (Retrieval Date UTC -> local): " & CStr(Stamped) & vbCrLf & _
        "Blank chhodi (lat/lon ya Retrieval Date missing): " & CStr(LeftBlank) & vbCrLf & vbCrLf & _
        "Watcher ON hai - nayi draws ab apne aap bhar + stamp hongi.", _
        vbInformation, "Solar EPC Resource"
    Exit Sub
Failed:
    MsgBox "FillAndStampNow failed: " & Err.Description, vbExclamation, "Solar EPC Resource"
End Sub
'''
s = s.rstrip() + '\n' + block

assert "' Version 3.7" in s
s = s.replace("' Version 3.7", "' Version 3.8", 1)
hist = "'   - v3.5: Date/Time stamp ab BLANK-ONLY"
assert hist in s
s = s.replace(hist, "'   - v3.8: SolarEPC_FillAndStampNow - THE BIG RED BUTTON: ek click me\n"
    "'     watcher ON + turant sweep + purani rows ka Date/Time Retrieval Date\n"
    "'     (UTC) se LOCAL time me backfill + B21 report. Koi reopen nahi.\n"
    + hist, 1)

io.open(ROOT + 'modSolarEPCResource.bas', 'w', encoding='utf-8').write(s)
print('v3.8 assembled:', s.count('\n') + 1, 'lines')
