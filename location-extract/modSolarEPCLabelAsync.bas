Option Explicit

'==========================================================================
' SOLAR EPC - ASYNC LOCATION LABEL SERVICE (companion to modSolarEPCResource)
' Version tag: 1.0.0 (companion; the resource module keeps its own version)
'
' WHY THIS MODULE EXISTS
'   The resource watcher must never wait on the internet. This companion
'   owns ALL reverse-geocoding HTTP for the Location column: a lightweight
'   4-second ticker picks rows whose Location is still blank but whose
'   Latitude/Longitude are filled, resolves the label with one BACKGROUND
'   request (OSM Nominatim, then BigDataCloud) and writes the BLANK Location
'   cell only. Manual entries, formula cells and filled cells are sacred.
'   Nothing here ever blocks Excel: every request is asynchronous and the
'   tick itself is a couple of bulk reads.
'==========================================================================

Private Const LABEL_DB_TABLE As String = "resource_db"
Private Const LABEL_TICK_SECONDS As Long = 4
Private Const LABEL_HTTP_TIMEOUT_SECONDS As Long = 20

Private mLabelRunning As Boolean
Private mLabelNext As Date
Private mLabelHttp As Object
Private mLabelKind As Long          '1 Nominatim, 2 BigDataCloud
Private mLabelLat As Double
Private mLabelLon As Double
Private mLabelStarted As Date
Private mLabelCacheKey As String
Private mLabelCacheText As String
Private mLabelDone As Object           'points whose label chain is exhausted

'Called by the resource watcher on every sweep; also safe to run manually.
Public Sub LabelAsyncEnsureRunning()
    If mLabelRunning Then Exit Sub
    mLabelRunning = True
    LabelSchedule 2
End Sub

Public Sub LabelAsyncStop()
    mLabelRunning = False
    On Error Resume Next
    If mLabelNext > 0 Then Application.OnTime EarliestTime:=mLabelNext, _
        Procedure:="LabelAsyncTick", Schedule:=False
    If Not mLabelHttp Is Nothing Then mLabelHttp.abort
    Set mLabelHttp = Nothing
    mLabelNext = 0
    On Error GoTo 0
End Sub

'OnTime entry point.
Public Sub LabelAsyncTick()
    mLabelNext = 0
    If Not mLabelRunning Then Exit Sub
    On Error Resume Next
    LabelCollectHttp
    If mLabelHttp Is Nothing Then LabelStartNextRequest
    LabelSchedule LABEL_TICK_SECONDS
    On Error GoTo 0
End Sub

Private Sub LabelSchedule(ByVal SecondsFromNow As Long)
    On Error Resume Next
    If mLabelNext > 0 Then Application.OnTime EarliestTime:=mLabelNext, _
        Procedure:="LabelAsyncTick", Schedule:=False
    mLabelNext = Now + TimeSerial(0, 0, SecondsFromNow)
    Application.OnTime EarliestTime:=mLabelNext, _
        Procedure:="LabelAsyncTick", Schedule:=True
    On Error GoTo 0
End Sub

'Offline-only label source for the watcher sweep: the in-memory cache of this
'service. Returns "" when nothing is cached - the sweep then leaves the cell
'blank and this service fills it in the background. NEVER touches the network.
Public Function ResourceLocationLabelOfflineSafe(ByVal Lat As Double, _
    ByVal Lon As Double) As String
    Dim KeyText As String
    KeyText = LabelDecimal(Lat, 8) & "," & LabelDecimal(Lon, 8)
    If KeyText = mLabelCacheKey Then ResourceLocationLabelOfflineSafe = mLabelCacheText
End Function

'Picks the first row that still needs a label and launches ONE background
'request for it. Bulk reads only; no per-cell COM calls.
Private Sub LabelStartNextRequest()
    Dim Tbl As ListObject
    Dim cLat As Long, cLon As Long, cLoc As Long
    Dim LatArr As Variant, LonArr As Variant, LocArr As Variant
    Dim n As Long, R As Long
    Dim Lat As Double, Lon As Double
    Dim UrlText As String
    Dim KeyText As String

    If mLabelDone Is Nothing Then Set mLabelDone = CreateObject("Scripting.Dictionary")
    Set Tbl = LabelDbTable()
    If Tbl Is Nothing Then Exit Sub
    cLat = LabelColumn(Tbl, "Latitude (" & ChrW(176) & ")")
    cLon = LabelColumn(Tbl, "Longitude (" & ChrW(176) & ")")
    cLoc = LabelColumn(Tbl, "Location")
    If cLat = 0 Or cLon = 0 Or cLoc = 0 Then Exit Sub
    n = Tbl.ListRows.Count
    If n = 0 Then Exit Sub
    LatArr = Tbl.ListColumns(cLat).Range.Value2
    LonArr = Tbl.ListColumns(cLon).Range.Value2
    LocArr = Tbl.ListColumns(cLoc).Range.Value2
    For R = 2 To n + 1
        If Len(LabelSafe(LocArr(R, 1))) = 0 And _
           IsNumeric(LatArr(R, 1)) And IsNumeric(LonArr(R, 1)) Then
            Lat = CDbl(LatArr(R, 1))
            Lon = CDbl(LonArr(R, 1))
            If Lat >= -90# And Lat <= 90# And Lon >= -180# And Lon <= 180# Then
                KeyText = LabelDecimal(Lat, 8) & "," & LabelDecimal(Lon, 8)
                If mLabelDone.Exists(KeyText) Then GoTo NextLabelRow
                If (Lat <> mLabelLat Or Lon <> mLabelLon Or mLabelKind = 0) Then mLabelKind = 1
                If mLabelKind = 1 Then
                    UrlText = "https://nominatim.openstreetmap.org/reverse?format=jsonv2&lat=" & _
                              LabelDecimal(Lat, 6) & "&lon=" & LabelDecimal(Lon, 6) & _
                              "&addressdetails=1&zoom=18&accept-language=en"
                Else
                    UrlText = "https://api.bigdatacloud.net/data/reverse-geocode-client?latitude=" & _
                              LabelDecimal(Lat, 6) & "&longitude=" & LabelDecimal(Lon, 6) & _
                              "&localityLanguage=en"
                End If
                Set mLabelHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
                mLabelHttp.Open "GET", UrlText, True
                mLabelHttp.setTimeouts 3000, 3000, 3000, LABEL_HTTP_TIMEOUT_SECONDS * 1000
                mLabelHttp.setRequestHeader "User-Agent", "Solar-EPC-Resource/1.0 (Excel workbook location label)"
                mLabelHttp.setRequestHeader "Accept", "application/json"
                mLabelHttp.send
                mLabelLat = Lat
                mLabelLon = Lon
                mLabelStarted = Now
                Exit Sub
            End If
        End If
NextLabelRow:
    Next R
End Sub

'Collects a finished background request and writes the label (blank-only).
Private Sub LabelCollectHttp()
    Dim BodyText As String
    Dim LabelText As String

    If mLabelHttp Is Nothing Then Exit Sub
    If mLabelHttp.readyState <> 4 Then
        If DateDiff("s", mLabelStarted, Now) < LABEL_HTTP_TIMEOUT_SECONDS Then Exit Sub
        mLabelHttp.abort
        Set mLabelHttp = Nothing
        BodyText = vbNullString
    Else
        If mLabelHttp.Status = 200 Then BodyText = mLabelHttp.ResponseText Else BodyText = vbNullString
        Set mLabelHttp = Nothing
    End If

    LabelText = LabelFromJson(BodyText, mLabelKind)
    If Len(LabelText) > 0 Then
        mLabelCacheKey = LabelDecimal(mLabelLat, 8) & "," & LabelDecimal(mLabelLon, 8)
        mLabelCacheText = LabelText
        LabelWriteToPoint mLabelLat, mLabelLon, LabelText
        mLabelKind = 0
    ElseIf mLabelKind = 1 Then
        mLabelKind = 2
    Else
        If mLabelKind = 2 Then
            If mLabelDone Is Nothing Then Set mLabelDone = CreateObject("Scripting.Dictionary")
            mLabelDone(LabelDecimal(mLabelLat, 8) & "," & LabelDecimal(mLabelLon, 8)) = True
        End If
        mLabelKind = 0
        mLabelLat = 0#
        mLabelLon = 0#
    End If
End Sub

Private Function LabelFromJson(ByVal JSONText As String, ByVal Kind As Long) As String
    Dim Primary As String, Secondary As String, DisplayText As String
    Dim Parts As Variant
    If Len(JSONText) = 0 Then Exit Function
    If Kind = 1 Then
        Primary = LabelFirstAddressPart(JSONText)
        Secondary = LabelAdminParts(JSONText, Primary)
        If Len(Primary) = 0 Then
            DisplayText = LabelJsonValue(JSONText, "display_name")
            If Len(DisplayText) > 0 Then
                Parts = Split(DisplayText, ", ")
                If UBound(Parts) >= 0 Then Primary = Trim$(CStr(Parts(0)))
                If UBound(Parts) >= 1 Then Secondary = Trim$(CStr(Parts(1)))
            End If
        End If
        If Len(Primary) > 0 Then LabelFromJson = LabelCompose(Primary, Secondary)
    Else
        Primary = LabelJsonValue(JSONText, "locality")
        If Len(Primary) = 0 Then Primary = LabelJsonValue(JSONText, "city")
        If Len(Primary) > 0 Then LabelFromJson = Primary
    End If
End Function

Private Sub LabelWriteToPoint(ByVal Lat As Double, ByVal Lon As Double, _
    ByVal LabelText As String)
    Dim Tbl As ListObject
    Dim cLat As Long, cLon As Long, cLoc As Long
    Dim LatArr As Variant, LonArr As Variant
    Dim n As Long, R As Long
    Dim Cell As Range
    Set Tbl = LabelDbTable()
    If Tbl Is Nothing Then Exit Sub
    cLat = LabelColumn(Tbl, "Latitude (" & ChrW(176) & ")")
    cLon = LabelColumn(Tbl, "Longitude (" & ChrW(176) & ")")
    cLoc = LabelColumn(Tbl, "Location")
    If cLat = 0 Or cLon = 0 Or cLoc = 0 Then Exit Sub
    n = Tbl.ListRows.Count
    If n = 0 Then Exit Sub
    LatArr = Tbl.ListColumns(cLat).Range.Value2
    LonArr = Tbl.ListColumns(cLon).Range.Value2
    For R = 2 To n + 1
        If IsNumeric(LatArr(R, 1)) And IsNumeric(LonArr(R, 1)) Then
            If Abs(CDbl(LatArr(R, 1)) - Lat) < 0.0000005 And _
               Abs(CDbl(LonArr(R, 1)) - Lon) < 0.0000005 Then
                Set Cell = Tbl.ListRows(R - 1).Range.Cells(1, cLoc)
                If Len(LabelSafe(Cell.Value2)) = 0 And Not Cell.HasFormula Then
                    Cell.Value2 = LabelText
                End If
            End If
        End If
    Next R
End Sub

Private Function LabelDbTable() As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        Set LabelDbTable = Nothing
        Set LabelDbTable = ws.ListObjects(LABEL_DB_TABLE)
        If Not LabelDbTable Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Private Function LabelColumn(ByVal Tbl As ListObject, ByVal HeaderText As String) As Long
    Dim Col As ListColumn
    On Error GoTo Failed
    For Each Col In Tbl.ListColumns
        If StrComp(Trim$(CStr(Col.Name)), HeaderText, vbTextCompare) = 0 Then
            LabelColumn = Col.Index
            Exit Function
        End If
    Next Col
Failed:
    LabelColumn = 0
End Function

Private Function LabelSafe(ByVal V As Variant) As String
    On Error GoTo Failed
    If IsError(V) Then Exit Function
    If IsNull(V) Then Exit Function
    LabelSafe = Trim$(CStr(V & ""))
Failed:
End Function

Private Function LabelDecimal(ByVal Value As Double, ByVal Places As Long) As String
    LabelDecimal = Replace$(Format$(Value, "0." & String$(Places, "0")), ",", ".")
End Function

Private Function LabelJsonValue(ByVal JSONText As String, ByVal KeyName As String) As String
    Dim Re As Object, Matches As Object
    On Error GoTo Failed
    Set Re = CreateObject("VBScript.RegExp")
    Re.Global = False
    Re.IgnoreCase = True
    Re.Pattern = """" & KeyName & """\s*:\s*(?:""([^""]*)""|([^,}\r\n]+))"
    If Re.Test(JSONText) Then
        Set Matches = Re.Execute(JSONText)
        If Len(Matches(0).SubMatches(0)) > 0 Then
            LabelJsonValue = Matches(0).SubMatches(0)
        Else
            LabelJsonValue = Trim$(Matches(0).SubMatches(1))
        End If
    End If
Failed:
End Function

Private Function LabelFirstAddressPart(ByVal JSONText As String) As String
    Dim Keys As Variant, i As Long, V As String
    Keys = Array("village", "hamlet", "suburb", "locality", _
                 "neighbourhood", "town", "city")
    For i = LBound(Keys) To UBound(Keys)
        V = Trim$(LabelJsonValue(JSONText, CStr(Keys(i))))
        If Len(V) > 0 Then
            LabelFirstAddressPart = V
            Exit Function
        End If
    Next i
End Function

Private Function LabelAdminClean(ByVal s As String) As String
    Dim Suffixes As Variant, i As Long, n As Long
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    If UCase$(Right$(s, 7)) = ", INDIA" Then s = Trim$(Left$(s, Len(s) - 7))
    Suffixes = Array(" taluka", " taluk", " tehsil", " tahsil", " subdivision", _
                     " mandal", " block", " district")
    For i = LBound(Suffixes) To UBound(Suffixes)
        n = Len(CStr(Suffixes(i)))
        If Len(s) > n Then
            If UCase$(Right$(s, n)) = UCase$(CStr(Suffixes(i))) Then
                s = Trim$(Left$(s, Len(s) - n))
                Exit For
            End If
        End If
    Next i
    LabelAdminClean = s
End Function

Private Function LabelAdminParts(ByVal JSONText As String, ByVal Locality As String) As String
    Dim Keys As Variant, i As Long, V As String
    Dim A1 As String, A2 As String
    Keys = Array("county", "district", "state_district")
    For i = LBound(Keys) To UBound(Keys)
        V = LabelAdminClean(LabelJsonValue(JSONText, CStr(Keys(i))))
        If Len(V) > 0 Then
            If StrComp(V, Locality, vbTextCompare) <> 0 Then
                If Len(A1) = 0 Then
                    A1 = V
                ElseIf StrComp(V, A1, vbTextCompare) <> 0 Then
                    If Len(A2) = 0 Then A2 = V Else Exit For
                End If
            End If
        End If
    Next i
    LabelAdminParts = LabelCompose(A1, A2)
End Function

Private Function LabelCompose(ByVal A As String, ByVal B As String) As String
    A = Trim$(A): B = Trim$(B)
    If Len(A) = 0 Then
        LabelCompose = B
    ElseIf Len(B) = 0 Then
        LabelCompose = A
    Else
        LabelCompose = A & ", " & B
    End If
End Function
