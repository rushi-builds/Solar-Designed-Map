'--------------------------------------------------------------------------
' v2.3 - AUTOMATIC LOCATION FILL (koi Alt+F8 nahi, koi popup nahi)
'
' Workbook khulte hi (Auto_Open) ye watcher chalu ho jata hai aur har
' AUTO_TICK_SECONDS me DRAWING_DATA.autoLWHTbl dekhta hai. Nayi ya badli hui
' SITE row milte hi:
'   1. exact area-weighted centroid offline calculate hota hai
'   2. RESOURCE_DB me us Project ID ki row dhundhi jaati hai
'   3. BLANK Latitude (deg) / Longitude (deg) cells me exact centroid likha
'      jaata hai (Worker wale centroid se identical)
'   4. BLANK Location cell me descriptive label likha jaata hai
'      ("Exact Area, Taluka, District") - VILLAGE_DB -> Google -> Nominatim
'      -> BigDataCloud, bilkul isi workbook ke existing rules se
'
' Kabhi bhi koi bhara hua cell overwrite NAHI hota, formula cells chhod diye
' jaate hain, aur RESOURCE_DB me NASA ke numbers/status ko ye code chhoota
' hi nahi. Feedback sirf status bar pe ek line me aata hai.
'--------------------------------------------------------------------------

Public Sub Auto_Open()
    On Error Resume Next
    SolarEPC_DrawnLocationAutoStart
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    SolarEPC_DrawnLocationAutoStop
End Sub

'Watcher chalu karo + turant ek sweep (already-drawn sites bhi bhar jaayen).
Public Sub SolarEPC_DrawnLocationAutoStart()
    On Error GoTo Failed
    If mAutoActive Then Exit Sub
    If mProcessed Is Nothing Then Set mProcessed = CreateObject("Scripting.Dictionary")
    mAutoActive = True
    DrawnLocationSweep
    DrawnLocationSchedule AUTO_TICK_SECONDS
    Application.StatusBar = "Solar EPC: drawn-site location AUTO-fill ON (har " & _
        CStr(AUTO_TICK_SECONDS) & "s DRAWING_DATA dekhta hai)."
    Exit Sub
Failed:
    mAutoActive = False
End Sub

'Watcher band karo (OnTime cancel ke saath).
Public Sub SolarEPC_DrawnLocationAutoStop()
    mAutoActive = False
    On Error Resume Next
    If mAutoNextRun > 0 Then
        Application.OnTime EarliestTime:=mAutoNextRun, _
            Procedure:="SolarEPC_DrawnLocationTick", Schedule:=False
    End If
    mAutoNextRun = 0
    Application.StatusBar = "Solar EPC: drawn-site location AUTO-fill OFF."
End Sub

'Bina wait kiye turant ek sweep - kisi button pe laga sakte ho, zaroori nahi.
Public Sub SolarEPC_DrawnLocationAutoNow()
    On Error GoTo Failed
    If mProcessed Is Nothing Then Set mProcessed = CreateObject("Scripting.Dictionary")
    DrawnLocationSweep
    Exit Sub
Failed:
End Sub

'Internal scheduler entry point (OnTime ko Public naam chahiye).
Public Sub SolarEPC_DrawnLocationTick()
    mAutoNextRun = 0
    If Not mAutoActive Then Exit Sub
    On Error Resume Next
    If mProcessed Is Nothing Then Set mProcessed = CreateObject("Scripting.Dictionary")
    DrawnLocationSweep
    DrawnLocationSchedule AUTO_TICK_SECONDS
End Sub

Private Sub DrawnLocationSchedule(ByVal SecondsFromNow As Long)
    On Error Resume Next
    If mAutoNextRun > 0 Then
        Application.OnTime EarliestTime:=mAutoNextRun, _
            Procedure:="SolarEPC_DrawnLocationTick", Schedule:=False
    End If
    mAutoNextRun = Now + TimeSerial(0, 0, SecondsFromNow)
    Application.OnTime EarliestTime:=mAutoNextRun, _
        Procedure:="SolarEPC_DrawnLocationTick", Schedule:=True
End Sub

'Ek baar saari SITE rows dekho; nayi/badli hui drawing pe auto-fill karo.
Private Sub DrawnLocationSweep()
    Dim Tbl As ListObject
    Dim R As Long
    Dim ReferenceText As String
    Dim ProjectID As String
    Dim CoordinateText As String
    Dim KeyText As String
    Dim StateText As String
    Dim Attempts As Long
    Dim Latitudes() As Double
    Dim Longitudes() As Double
    Dim VertexCount As Long
    Dim CentroidLatitude As Double
    Dim CentroidLongitude As Double
    Dim AreaM2 As Double
    Dim ErrorText As String
    Dim ResultText As String

    On Error GoTo Done
    Set Tbl = DrawnSiteTable()
    If Tbl Is Nothing Then Exit Sub
    If mProcessed Is Nothing Then Set mProcessed = CreateObject("Scripting.Dictionary")

    For R = 1 To Tbl.ListRows.Count
        ReferenceText = RowCellText(Tbl, R, "Reference id")
        If UCase$(Left$(ReferenceText, 9)) = "SITE-MAP-" Then
            ProjectID = RowCellText(Tbl, R, "Project ID")
            CoordinateText = RowCellText(Tbl, R, "Coordinates")
            If Len(ProjectID) > 0 And Len(CoordinateText) > 0 Then
                'Key me coordinates bhi hain: site dobara draw hui to key badal
                'jaati hai aur fill dobara attempt hota hai.
                KeyText = ReferenceText & "|" & CoordinateText
                If Not mProcessed.Exists(KeyText) Or _
                   Left$(CStr(mProcessed(KeyText)), 5) = "NOROW" Or _
                   Left$(CStr(mProcessed(KeyText)), 7) = "LOCPEND" Then

                    StateText = CStr(mProcessed(KeyText) & "")
                    Attempts = DrawnLocationAttempts(StateText)

                    If ParsePolygon(CoordinateText, Latitudes, Longitudes, VertexCount, ErrorText) Then
                        If CentroidLocalProjection(Latitudes, Longitudes, VertexCount, _
                               CentroidLatitude, CentroidLongitude, AreaM2, ErrorText) Then
                            If AreaM2 >= SITE_MIN_AREA_M2 Then
                                ResultText = FillResourceDbForSite(ProjectID, CentroidLatitude, _
                                    CentroidLongitude, Attempts >= 2)
                                If InStr(1, ResultText, "LOCPEND", vbBinaryCompare) > 0 Then
                                    'lat/lon bhar gaye, label offline tha - limited retries.
                                    Application.StatusBar = "Solar EPC: " & ProjectID & _
                                        " ke lat/lon bhar gaye; Location label offline hai, retry " & _
                                        CStr(Attempts + 1) & "/" & CStr(AUTO_LABEL_RETRIES) & _
                                        " (internet / VILLAGE_DB / geocoding key dekho)"
                                    If Attempts < AUTO_LABEL_RETRIES Then _
                                        mProcessed(KeyText) = "LOCPEND:" & CStr(Attempts + 1)
                                ElseIf ResultText = "NOROW" Then
                                    'RESOURCE_DB row abhi bani nahi - retry, phir auto-create.
                                    Application.StatusBar = "Solar EPC: RESOURCE_DB me " & ProjectID & _
                                        " ki row nahi mili (" & CStr(Attempts + 1) & ") - " & _
                                        IIf(Attempts + 1 >= 2, "ab row bana ke bhar raha hoon...", "retry...")
                                    If Attempts < AUTO_NOROW_RETRIES Then _
                                        mProcessed(KeyText) = "NOROW:" & CStr(Attempts + 1)
                                Else
                                    mProcessed(KeyText) = "DONE"
                                    If ResultText <> "SKIP-NO-BLANK" Then
                                        Application.StatusBar = "Solar EPC: exact drawn location " & _
                                            "auto-filled -> " & ProjectID & " | " & ResultText
                                        DrawnLocationDiag ProjectID, ResultText, _
                                            CentroidLatitude, CentroidLongitude
                                    End If
                                End If
                            Else
                                mProcessed(KeyText) = "DONE"
                            End If
                        Else
                            mProcessed(KeyText) = "DONE"
                        End If
                    Else
                        mProcessed(KeyText) = "DONE"
                    End If
                End If
            End If
        End If
    Next R
Done:
End Sub

'"NOROW:3" / "LOCPEND:1" jaisi state me se attempt count nikaalo.
Private Function DrawnLocationAttempts(ByVal StateText As String) As Long
    Dim P As Long
    P = InStr(1, StateText, ":", vbBinaryCompare)
    If P > 0 Then DrawnLocationAttempts = Val(Mid$(StateText, P + 1))
End Function

'Hidden _CLOUD_CFG me last auto-fill ka record (diagnostic only; sheet na ho
'to chupchaap skip).
Private Sub DrawnLocationDiag(ByVal ProjectID As String, ByVal ResultText As String, _
    ByVal CentroidLatitude As Double, ByVal CentroidLongitude As Double)

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    If Not ws Is Nothing Then
        ws.Range("A18").Value2 = "LAST AUTO-FILLED DRAWN LOCATION"
        ws.Range("B18").Value2 = ProjectID & " | " & _
            Decimal8(CentroidLatitude) & ", " & Decimal8(CentroidLongitude) & " | " & ResultText
    End If
    On Error GoTo 0
End Sub

'RESOURCE_DB me us Project ID ki row dhundh ke BLANK lat/lon + BLANK Location
'bharo. Return: kya bhara ("lat lon location(Sonwadi Bk., Phaltan, Satara)"),
'ya SKIP-NO-BLANK / NOROW / LOCPEND...
Private Function FillResourceDbForSite(ByVal ProjectID As String, _
    ByVal CentroidLatitude As Double, ByVal CentroidLongitude As Double, _
    Optional ByVal AllowCreate As Boolean = False) As String

    Dim Tbl As ListObject
    Dim R As Long
    Dim cProject As Long, cLat As Long, cLon As Long, cLoc As Long
    Dim RowRange As Range
    Dim Target As Range
    Dim RowProject As String
    Dim RowLat As String, RowLon As String
    Dim Cell As Range
    Dim LabelText As String
    Dim DidText As String
    Dim LatBlank As Boolean, LonBlank As Boolean, LocBlank As Boolean

    On Error GoTo Failed
    Set Tbl = ResourceDbTable()
    If Tbl Is Nothing Then
        FillResourceDbForSite = "NOROW"
        Exit Function
    End If
    cProject = TableColumn(Tbl, "Project ID")
    cLat = TableColumn(Tbl, "Latitude (" & ChrW(176) & ")")
    cLon = TableColumn(Tbl, "Longitude (" & ChrW(176) & ")")
    cLoc = TableColumn(Tbl, "Location")
    If cProject = 0 Or cLat = 0 Or cLon = 0 Then
        FillResourceDbForSite = "NOROW"
        Exit Function
    End If

    'Row match: same Project ID jiske lat/lon blank hain, ya same Project ID
    'jiske lat/lon already isi centroid ke barabar hain (Location blank ho sakta hai).
    For R = 1 To Tbl.ListRows.Count
        Set RowRange = Tbl.ListRows(R).Range
        RowProject = Trim$(CStr(RowRange.Cells(1, cProject).Value2))
        If StrComp(RowProject, ProjectID, vbTextCompare) = 0 Then
            RowLat = Trim$(CStr(RowRange.Cells(1, cLat).Value2))
            RowLon = Trim$(CStr(RowRange.Cells(1, cLon).Value2))
            If Len(RowLat) = 0 And Len(RowLon) = 0 Then
                Set Target = RowRange
                Exit For
            ElseIf IsNumeric(RowLat) And IsNumeric(RowLon) Then
                If Abs(CDbl(RowLat) - CentroidLatitude) < 0.0000005 And _
                   Abs(CDbl(RowLon) - CentroidLongitude) < 0.0000005 Then
                    Set Target = RowRange
                    Exit For
                End If
            End If
        End If
    Next R
    'v2.5: row hai hi nahi to bana do - draw kiya hai to location RESOURCE_DB
    'me dikhni chahiye. Baaki columns NASA import / manual bharenge.
    If Target Is Nothing Then
        If Not AllowCreate Then
            FillResourceDbForSite = "NOROW"
            Exit Function
        End If
        Set Target = Tbl.ListRows.Add.Range
        If cProject > 0 Then
            Set Cell = Target.Cells(1, cProject)
            If Len(Trim$(CStr(Cell.Value2))) = 0 And Not Cell.HasFormula Then Cell.Value2 = ProjectID
        End If
        DidText = "NEWROW "
    End If

    '1. exact centroid - BLANK lat/lon cells only.
    Set Cell = Target.Cells(1, cLat)
    LatBlank = (Len(Trim$(CStr(Cell.Value2))) = 0)
    If LatBlank And Not Cell.HasFormula Then
        Cell.Value2 = CentroidLatitude
        DidText = DidText & "lat "
    End If
    Set Cell = Target.Cells(1, cLon)
    LonBlank = (Len(Trim$(CStr(Cell.Value2))) = 0)
    If LonBlank And Not Cell.HasFormula Then
        Cell.Value2 = CentroidLongitude
        DidText = DidText & "lon "
    End If

    '2. descriptive Location label - BLANK cell only, existing tiers se.
    If cLoc > 0 Then
        Set Cell = Target.Cells(1, cLoc)
        LocBlank = (Len(Trim$(CStr(Cell.Value2))) = 0)
        If LocBlank And Not Cell.HasFormula Then
            LabelText = ResourceLocationLabel(Decimal8(CentroidLatitude), Decimal8(CentroidLongitude))
            If Len(LabelText) > 0 Then
                Cell.Value2 = LabelText
                DidText = DidText & "location(" & LabelText & ")"
            Else
                DidText = DidText & "LOCPEND"
            End If
        End If
    End If

    If Len(DidText) = 0 Then DidText = "SKIP-NO-BLANK"
    FillResourceDbForSite = DidText
    Exit Function
Failed:
    FillResourceDbForSite = "NOROW"
End Function

'resource_db Excel Table (naam se), warna Nothing.
Private Function ResourceDbTable() As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        Set ResourceDbTable = Nothing
        Set ResourceDbTable = ws.ListObjects(RESOURCE_DB_TABLE)
        If Not ResourceDbTable Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

'Column index by header name, 0 agar column nahi hai.
Private Function TableColumn(ByVal Tbl As ListObject, ByVal HeaderText As String) As Long
    Dim Col As ListColumn
    On Error GoTo Failed
    For Each Col In Tbl.ListColumns
        If StrComp(Trim$(CStr(Col.Name)), HeaderText, vbTextCompare) = 0 Then
            TableColumn = Col.Index
            Exit Function
        End If
    Next Col
    Exit Function
Failed:
    TableColumn = 0
End Function
