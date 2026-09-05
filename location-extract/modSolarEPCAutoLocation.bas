Attribute VB_Name = "modSolarEPCAutoLocation"
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
Private Const SITE_MAX_VERTICES As Long = 1000
Private Const AUTO_TICK_SECONDS As Long = 5
Private Const AUTO_NOROW_RETRIES As Long = 12
Private Const AUTO_LABEL_RETRIES As Long = 6

Private mAutoActive As Boolean
Private mAutoNextRun As Date
Private mProcessed As Object
Private mLastLocKey As String
Private mLastLocLabel As String

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
    SolarEPC_AutoLocationStart
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    SolarEPC_AutoLocationStop
End Sub

'Watcher chalu karo + turant ek sweep (already-drawn sites bhi bhar jaayen).
Public Sub SolarEPC_AutoLocationStart()
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
Public Sub SolarEPC_AutoLocationStop()
    mAutoActive = False
    On Error Resume Next
    If mAutoNextRun > 0 Then
        Application.OnTime EarliestTime:=mAutoNextRun, _
            Procedure:="SolarEPC_AutoLocationTick", Schedule:=False
    End If
    mAutoNextRun = 0
    Application.StatusBar = "Solar EPC: drawn-site location AUTO-fill OFF."
End Sub

'Bina wait kiye turant ek sweep - kisi button pe laga sakte ho, zaroori nahi.
Public Sub SolarEPC_AutoLocationNow()
    On Error GoTo Failed
    If mProcessed Is Nothing Then Set mProcessed = CreateObject("Scripting.Dictionary")
    DrawnLocationSweep
    Exit Sub
Failed:
End Sub

'Internal scheduler entry point (OnTime ko Public naam chahiye).
Public Sub SolarEPC_AutoLocationTick()
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
            Procedure:="SolarEPC_AutoLocationTick", Schedule:=False
    End If
    mAutoNextRun = Now + TimeSerial(0, 0, SecondsFromNow)
    Application.OnTime EarliestTime:=mAutoNextRun, _
        Procedure:="SolarEPC_AutoLocationTick", Schedule:=True
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

'--------------------------------------------------------------------------
' v2.4 adapters - watcher ke naampari helpers ko complete module ke proven
' helpers se jodte hain (koi logic duplicate nahi).
'--------------------------------------------------------------------------
Private Function DrawnSiteTable() As ListObject
    On Error GoTo Failed
    Set DrawnSiteTable = ResourceSiteTable()
    Exit Function
Failed:
    Set DrawnSiteTable = Nothing
End Function

Private Function RowCellText(ByVal Tbl As ListObject, ByVal RowIndex As Long, _
    ByVal HeaderText As String) As String
    On Error GoTo Failed
    RowCellText = ResourceRowCellText(Tbl.ListRows(RowIndex), HeaderText)
    Exit Function
Failed:
    RowCellText = vbNullString
End Function

Private Function Decimal8(ByVal Value As Double) As String
    Decimal8 = ResourceDecimal(Value, 8)
End Function

Private Function ParsePolygon(ByVal CoordinateText As String, _
    ByRef Latitudes() As Double, ByRef Longitudes() As Double, _
    ByRef VertexCount As Long, ByRef ErrorText As String) As Boolean
    ParsePolygon = ResourceParseSitePolygon(CoordinateText, Latitudes, _
        Longitudes, VertexCount, ErrorText)
End Function

Private Function CentroidLocalProjection(ByRef Latitudes() As Double, _
    ByRef Longitudes() As Double, ByVal VertexCount As Long, _
    ByRef CentroidLatitude As Double, ByRef CentroidLongitude As Double, _
    ByRef AreaM2 As Double, ByRef ErrorText As String) As Boolean
    CentroidLocalProjection = ResourceCentroidLocalProjection(Latitudes, Longitudes, _
        VertexCount, CentroidLatitude, CentroidLongitude, AreaM2, ErrorText)
End Function

Private Function ResourceColumn(ByVal Tbl As ListObject, ByVal HeaderText As String) As Long
    Dim C As ListColumn
    For Each C In Tbl.ListColumns
        If StrComp(Trim$(CStr(C.Name)), HeaderText, vbTextCompare) = 0 Then
            ResourceColumn = C.Index
            Exit Function
        End If
    Next C
End Function

Private Function ResourceSiteTable() As ListObject
    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(DRAWING_SHEET)
    Set ResourceSiteTable = ws.ListObjects(SITE_TABLE)
    Exit Function
Failed:
    Set ResourceSiteTable = Nothing
End Function

Private Function ResourceRowCellText(ByVal Row As ListRow, ByVal HeaderText As String) As String
    Dim C As Long
    On Error GoTo Failed
    C = ResourceColumn(Row.Parent, HeaderText)
    If C = 0 Then Exit Function
    ResourceRowCellText = Trim$(CStr(Row.Range.Cells(1, C).Value2))
    Exit Function
Failed:
    ResourceRowCellText = vbNullString
End Function

Private Function ResourceStrictNumber(ByVal TextValue As String) As Boolean
    Dim Re As Object
    Set Re = CreateObject("VBScript.RegExp")
    Re.Global = False
    Re.Pattern = "^-?[0-9]+(?:\.[0-9]+)?$"
    ResourceStrictNumber = Re.Test(Trim$(TextValue))
End Function

Private Function ResourceDecimal(ByVal Value As Double, ByVal Places As Long) As String
    ResourceDecimal = Replace$(Format$(Value, "0." & String$(Places, "0")), ",", ".")
End Function

Private Function ResourceParseSitePolygon(ByVal CoordinateText As String, _
    ByRef Latitudes() As Double, ByRef Longitudes() As Double, _
    ByRef VertexCount As Long, ByRef ErrorText As String) As Boolean

    Dim RawText As String
    Dim Pairs As Variant
    Dim Parts As Variant
    Dim SourceLat() As Double
    Dim SourceLon() As Double
    Dim RawCount As Long
    Dim KeepCount As Long
    Dim i As Long
    Dim LatitudeValue As Double
    Dim LongitudeValue As Double

    On Error GoTo Failed
    ErrorText = vbNullString
    VertexCount = 0
    RawText = Trim$(CoordinateText)
    If Len(RawText) = 0 Then
        ErrorText = "The SITE Coordinates cell is blank."
        Exit Function
    End If

    Pairs = Split(RawText, ";")
    ReDim SourceLat(0 To UBound(Pairs))
    ReDim SourceLon(0 To UBound(Pairs))
    RawCount = 0

    For i = LBound(Pairs) To UBound(Pairs)
        If Len(Trim$(CStr(Pairs(i)))) > 0 Then
            Parts = Split(Trim$(CStr(Pairs(i))), ",")
            If UBound(Parts) <> 1 Then
                ErrorText = "Each coordinate pair must be written as latitude,longitude."
                Exit Function
            End If
            If Not ResourceStrictNumber(CStr(Parts(0))) Or _
               Not ResourceStrictNumber(CStr(Parts(1))) Then
                ErrorText = "A coordinate is not a plain decimal number."
                Exit Function
            End If
            LatitudeValue = Val(Trim$(CStr(Parts(0))))
            LongitudeValue = Val(Trim$(CStr(Parts(1))))
            If LatitudeValue < -90# Or LatitudeValue > 90# Or _
               LongitudeValue < -180# Or LongitudeValue > 180# Then
                ErrorText = "A coordinate is outside the valid WGS84 range."
                Exit Function
            End If
            'The Worker rounds every vertex to 8 decimals before use.
            SourceLat(RawCount) = Val(ResourceDecimal(LatitudeValue, 8))
            SourceLon(RawCount) = Val(ResourceDecimal(LongitudeValue, 8))
            RawCount = RawCount + 1
        End If
    Next i

    If RawCount = 0 Then
        ErrorText = "No coordinate pair was found in the SITE Coordinates cell."
        Exit Function
    End If
    If RawCount > SITE_MAX_VERTICES Then
        ErrorText = "The SITE polygon exceeds " & CStr(SITE_MAX_VERTICES) & " vertices."
        Exit Function
    End If

    ReDim Latitudes(0 To RawCount - 1)
    ReDim Longitudes(0 To RawCount - 1)
    KeepCount = 0
    For i = 0 To RawCount - 1
        'Drop a vertex that repeats the previously kept one (8-decimal compare).
        If KeepCount > 0 Then
            If Latitudes(KeepCount - 1) = SourceLat(i) And _
               Longitudes(KeepCount - 1) = SourceLon(i) Then GoTo NextVertex
        End If
        Latitudes(KeepCount) = SourceLat(i)
        Longitudes(KeepCount) = SourceLon(i)
        KeepCount = KeepCount + 1
NextVertex:
    Next i

    'Drop an explicitly repeated closing vertex.
    If KeepCount > 1 Then
        If Latitudes(0) = Latitudes(KeepCount - 1) And _
           Longitudes(0) = Longitudes(KeepCount - 1) Then KeepCount = KeepCount - 1
    End If

    If KeepCount < 3 Then
        ErrorText = "The SITE polygon needs at least three distinct vertices."
        Exit Function
    End If
    If ResourceDistinctVertexCount(Latitudes, Longitudes, KeepCount) < 3 Then
        ErrorText = "The SITE polygon needs at least three distinct vertices."
        Exit Function
    End If

    VertexCount = KeepCount
    ResourceParseSitePolygon = True
    Exit Function
Failed:
    ErrorText = "Polygon parse error " & CStr(Err.Number) & ": " & Err.Description
    ResourceParseSitePolygon = False
End Function

Private Function ResourceCentroidLocalProjection(ByRef Latitudes() As Double, _
    ByRef Longitudes() As Double, ByVal VertexCount As Long, _
    ByRef CentroidLatitude As Double, ByRef CentroidLongitude As Double, _
    ByRef AreaM2 As Double, ByRef ErrorText As String) As Boolean

    Const PI As Double = 3.14159265358979
    Dim ProjX() As Double
    Dim ProjY() As Double
    Dim SumLatitude As Double
    Dim SumLongitude As Double
    Dim MeanLatitude As Double
    Dim MeanLongitude As Double
    Dim CosineLatitude As Double
    Dim TwiceArea As Double
    Dim CentroidX6A As Double
    Dim CentroidY6A As Double
    Dim CrossProduct As Double
    Dim CentroidX As Double
    Dim CentroidY As Double
    Dim i As Long
    Dim j As Long

    On Error GoTo Failed
    ErrorText = vbNullString
    If VertexCount < 3 Then
        ErrorText = "The SITE polygon needs at least three vertices."
        Exit Function
    End If

    SumLatitude = 0#
    SumLongitude = 0#
    For i = 0 To VertexCount - 1
        SumLatitude = SumLatitude + Latitudes(i)
        SumLongitude = SumLongitude + Longitudes(i)
    Next i
    MeanLatitude = SumLatitude / CDbl(VertexCount)
    MeanLongitude = SumLongitude / CDbl(VertexCount)

    CosineLatitude = Cos(MeanLatitude * PI / 180#)
    If Abs(CosineLatitude) < 0.000000000001 Then
        ErrorText = "The polygon is too close to a pole for this centroid projection."
        Exit Function
    End If

    ReDim ProjX(0 To VertexCount - 1)
    ReDim ProjY(0 To VertexCount - 1)
    For i = 0 To VertexCount - 1
        ProjX(i) = (Longitudes(i) - MeanLongitude) * PI / 180# * EARTH_RADIUS_M * CosineLatitude
        ProjY(i) = (Latitudes(i) - MeanLatitude) * PI / 180# * EARTH_RADIUS_M
    Next i

    TwiceArea = 0#
    CentroidX6A = 0#
    CentroidY6A = 0#
    For i = 0 To VertexCount - 1
        j = (i + 1) Mod VertexCount
        CrossProduct = ProjX(i) * ProjY(j) - ProjX(j) * ProjY(i)
        TwiceArea = TwiceArea + CrossProduct
        CentroidX6A = CentroidX6A + (ProjX(i) + ProjX(j)) * CrossProduct
        CentroidY6A = CentroidY6A + (ProjY(i) + ProjY(j)) * CrossProduct
    Next i

    If Abs(TwiceArea) < 0.000000001 Then
        ErrorText = "The SITE polygon has zero area (all vertices are collinear)."
        Exit Function
    End If

    CentroidX = CentroidX6A / (3# * TwiceArea)
    CentroidY = CentroidY6A / (3# * TwiceArea)

    CentroidLatitude = MeanLatitude + (CentroidY / EARTH_RADIUS_M) * 180# / PI
    CentroidLongitude = MeanLongitude + (CentroidX / (EARTH_RADIUS_M * CosineLatitude)) * 180# / PI
    AreaM2 = Abs(TwiceArea) / 2#

    ResourceCentroidLocalProjection = True
    Exit Function
Failed:
    ErrorText = "Centroid error " & CStr(Err.Number) & ": " & Err.Description
    ResourceCentroidLocalProjection = False
End Function

Private Function ResourceHttpGetJson(ByVal UrlText As String) As String
    Dim HTTP As Object
    On Error GoTo TryFallback
    Set HTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    HTTP.Open "GET", UrlText, False
    HTTP.setTimeouts 6000, 6000, 15000, 15000
    HTTP.setRequestHeader "User-Agent", "Solar-EPC-Resource/1.0 (Excel workbook location label)"
    HTTP.setRequestHeader "Accept", "application/json"
    HTTP.send
    If HTTP.Status = 200 Then
        ResourceHttpGetJson = HTTP.ResponseText
        Exit Function
    End If
TryFallback:
    On Error GoTo Failed
    Set HTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    HTTP.Open "GET", UrlText, False
    HTTP.setTimeouts 6000, 6000, 15000, 15000
    HTTP.setRequestHeader "User-Agent", "Solar-EPC-Resource/1.0 (Excel workbook location label)"
    HTTP.setRequestHeader "Accept", "application/json"
    HTTP.send
    If HTTP.Status = 200 Then ResourceHttpGetJson = HTTP.ResponseText
    Exit Function
Failed:
    ResourceHttpGetJson = vbNullString
End Function

Private Function ResourceJSONValue(ByVal JSONText As String, ByVal KeyName As String) As String
    Dim Re As Object
    Dim Matches As Object
    Set Re = CreateObject("VBScript.RegExp")
    Re.Global = False
    Re.IgnoreCase = True
    Re.Pattern = """" & KeyName & """\s*:\s*(?:""([^""]*)""|([^,}\r\n]+))"
    If Re.Test(JSONText) Then
        Set Matches = Re.Execute(JSONText)
        If Len(Matches(0).SubMatches(0)) > 0 Then
            ResourceJSONValue = Matches(0).SubMatches(0)
        Else
            ResourceJSONValue = Trim$(Matches(0).SubMatches(1))
        End If
    End If
End Function

Private Function ResourceLocationCompose(ByVal Locality As String, ByVal AdminText As String) As String
    Dim L As String, A As String
    L = Trim$(Locality)
    A = Trim$(AdminText)
    If Len(L) = 0 Then
        ResourceLocationCompose = A
    ElseIf Len(A) = 0 Then
        ResourceLocationCompose = L
    Else
        ResourceLocationCompose = L & ", " & A
    End If
End Function

Private Function ResourceSameName(ByVal A As String, ByVal B As String) As Boolean
    ResourceSameName = (StrComp(Trim$(A), Trim$(B), vbTextCompare) = 0)
End Function

Private Function ResourceAdminClean(ByVal TextValue As String) As String
    Dim S As String
    Dim Suffixes As Variant
    Dim i As Long, N As Long
    S = Trim$(TextValue)
    If Len(S) = 0 Then Exit Function
    If UCase$(Right$(S, 7)) = ", INDIA" Then S = Trim$(Left$(S, Len(S) - 7))
    Suffixes = Array(" taluka", " taluk", " tehsil", " tahsil", " subdivision", _
                     " mandal", " block", " district")
    For i = LBound(Suffixes) To UBound(Suffixes)
        N = Len(CStr(Suffixes(i)))
        If Len(S) > N Then
            If UCase$(Right$(S, N)) = UCase$(CStr(Suffixes(i))) Then
                S = Trim$(Left$(S, Len(S) - N))
                Exit For
            End If
        End If
    Next i
    ResourceAdminClean = S
End Function

Private Function ResourceLocalityName(ByVal JSONText As String) As String
    Dim Keys As Variant
    Dim i As Long, V As String
    Keys = Array("village", "hamlet", "suburb", "locality", _
                 "neighbourhood", "town", "city")
    For i = LBound(Keys) To UBound(Keys)
        V = Trim$(ResourceJSONValue(JSONText, CStr(Keys(i))))
        If Len(V) > 0 Then
            ResourceLocalityName = V
            Exit Function
        End If
    Next i
End Function

Private Function ResourceAdminNames(ByVal JSONText As String, ByVal Locality As String) As String
    Dim Keys As Variant
    Dim i As Long, V As String
    Dim A1 As String, A2 As String
    Keys = Array("county", "district", "state_district")   'NOTE: "state" excluded on purpose
    For i = LBound(Keys) To UBound(Keys)
        V = ResourceAdminClean(ResourceJSONValue(JSONText, CStr(Keys(i))))
        If Len(V) > 0 Then
            If Not ResourceSameName(V, Locality) Then
                If Len(A1) = 0 Then
                    A1 = V
                ElseIf Not ResourceSameName(V, A1) Then
                    If Len(A2) = 0 Then A2 = V Else Exit For
                End If
            End If
        End If
    Next i
    ResourceAdminNames = ResourceLocationCompose(A1, A2)
End Function

Private Function ResourceDistanceKm(ByVal Lat1 As Double, ByVal Lon1 As Double, _
    ByVal Lat2 As Double, ByVal Lon2 As Double) As Double

    Const PI As Double = 3.14159265358979
    Dim X As Double, Y As Double
    On Error GoTo Failed
    X = (Lon2 - Lon1) * Cos((Lat1 + Lat2) * PI / 360#)
    Y = Lat2 - Lat1
    ResourceDistanceKm = Sqr(X * X + Y * Y) * 111.32
    Exit Function
Failed:
    ResourceDistanceKm = 9999999#
End Function

Private Function ResourceVillageDbLabel(ByVal Lat As Double, ByVal Lon As Double) As String
    Dim ws As Worksheet
    Dim R As Long, LastRow As Long
    Dim Village As String, Taluka As String, District As String
    Dim BestKm As Double, Km As Double
    Dim BestVillage As String, BestAdmin As String

    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(VILLAGE_SHEET)
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If LastRow < 2 Then Exit Function

    BestKm = VILLAGE_MAX_KM
    For R = 2 To LastRow
        Village = Trim$(CStr(ws.Cells(R, "A").Value2))
        If Len(Village) > 0 Then
            If IsNumeric(ws.Cells(R, "D").Value2) And IsNumeric(ws.Cells(R, "E").Value2) Then
                Km = ResourceDistanceKm(Lat, Lon, _
                    CDbl(ws.Cells(R, "D").Value2), CDbl(ws.Cells(R, "E").Value2))
                If Km < BestKm Then
                    BestKm = Km
                    BestVillage = Village
                    Taluka = ResourceAdminClean(Trim$(CStr(ws.Cells(R, "B").Value2)))
                    District = ResourceAdminClean(Trim$(CStr(ws.Cells(R, "C").Value2)))
                    BestAdmin = ResourceLocationCompose(Taluka, District)
                End If
            End If
        End If
    Next R

    If Len(BestVillage) > 0 Then
        ResourceVillageDbLabel = ResourceLocationCompose(BestVillage, BestAdmin)
        On Error Resume Next
        With ThisWorkbook.Worksheets(CONFIG_SHEET)
            .Range("A16").Value2 = "VILLAGE DB MATCH"
            .Range("B16").Value2 = BestVillage & " (" & Format$(BestKm, "0.00") & " km)"
        End With
        On Error GoTo 0
        Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"), "VILLAGE DB:", _
            BestVillage, Format$(BestKm, "0.00") & " km"
    End If
    Exit Function
Failed:
    ResourceVillageDbLabel = vbNullString
End Function

Private Function ResourceCopyToClipboard(ByVal TextValue As String) As Boolean
    Dim DataObject As Object
    On Error GoTo Failed
    Set DataObject = CreateObject("MSForms.DataObject")
    DataObject.SetText TextValue
    DataObject.PutInClipboard
    ResourceCopyToClipboard = True
    Exit Function
Failed:
    ResourceCopyToClipboard = False
End Function

Private Function ResourceLocationLabel(ByVal LatitudeText As String, ByVal LongitudeText As String) As String
    Dim Lat As Double, Lon As Double
    Dim UrlText As String, CacheKey As String, JSONText As String
    Dim Primary As String, Secondary As String
    Dim Parts As Variant
    Dim DisplayText As String

    On Error GoTo Failed
    If Not IsNumeric(LatitudeText) Or Not IsNumeric(LongitudeText) Then Exit Function
    Lat = CDbl(LatitudeText)
    Lon = CDbl(LongitudeText)
    If Lat < -90 Or Lat > 90 Or Lon < -180 Or Lon > 180 Then Exit Function

    CacheKey = ResourceDecimal(Lat, 6) & "," & ResourceDecimal(Lon, 6)
    If CacheKey = mLastLocKey Then
        ResourceLocationLabel = mLastLocLabel
        Exit Function
    End If

    '0) v2.0: the user's own VILLAGE_DB - nearest known village within 5 km.
    '   Free, offline, and the only tier that carries real revenue-village
    '   names. Beats every online geocoder on purpose.
    ResourceLocationLabel = ResourceVillageDbLabel(Lat, Lon)
    If Len(ResourceLocationLabel) > 0 Then GoTo StoreCache


    '1) OpenStreetMap Nominatim (free, no API key)
    '   zoom=18 asks for the most detailed address block - village / hamlet /
    '   suburb / neighbourhood are only populated at the higher zoom levels.
    UrlText = "https://nominatim.openstreetmap.org/reverse?format=jsonv2&lat=" & _
              ResourceDecimal(Lat, 6) & "&lon=" & ResourceDecimal(Lon, 6) & _
              "&addressdetails=1&zoom=18&accept-language=en"
    JSONText = ResourceHttpGetJson(UrlText)
    If Len(JSONText) > 0 Then
        Primary = ResourceLocalityName(JSONText)
        Secondary = ResourceAdminNames(JSONText, Primary)
        If Len(Primary) = 0 Then
            'Last resort only. display_name is still geocoder output (never
            'invented) and its first element is the most specific name the
            'geocoder could resolve for this exact point.
            DisplayText = ResourceJSONValue(JSONText, "display_name")
            If Len(DisplayText) > 0 Then
                Parts = Split(DisplayText, ", ")
                If UBound(Parts) >= 0 Then Primary = Trim$(CStr(Parts(0)))
                If UBound(Parts) >= 1 Then Secondary = Trim$(CStr(Parts(1)))
            End If
        End If
        If Len(Primary) > 0 Then
            ResourceLocationLabel = ResourceLocationCompose(Primary, Secondary)
            GoTo StoreCache
        End If
    End If

    '2) BigDataCloud free client reverse-geocode (no key) fallback.
    '   This provider exposes no taluka at top level, so only the locality is
    '   used - state is deliberately not appended.
    UrlText = "https://api.bigdatacloud.net/data/reverse-geocode-client?latitude=" & _
              ResourceDecimal(Lat, 6) & "&longitude=" & ResourceDecimal(Lon, 6) & _
              "&localityLanguage=en"
    JSONText = ResourceHttpGetJson(UrlText)
    If Len(JSONText) > 0 Then
        Primary = ResourceJSONValue(JSONText, "locality")
        If Len(Primary) = 0 Then Primary = ResourceJSONValue(JSONText, "city")
        If Len(Primary) > 0 Then ResourceLocationLabel = Primary
    End If

StoreCache:
    mLastLocKey = CacheKey
    mLastLocLabel = ResourceLocationLabel
    Exit Function
Failed:
    ResourceLocationLabel = vbNullString
    mLastLocKey = vbNullString
    mLastLocLabel = vbNullString
End Function
'--------------------------------------------------------------------------
' v2.5 - DIAGNOSTIC: chain kaha atki hai, ek click me batata hai.
'   Alt+F8 -> SolarEPC_DrawnLocationDebug
' Kuch bhi change NAHI karta (sirf padhta hai + report dikhata hai).
'--------------------------------------------------------------------------
Public Sub SolarEPC_AutoLocationDebug()
    Dim P1 As String, P2 As String
    Dim Tbl As ListObject
    Dim Rdb As ListObject
    Dim R As Long
    Dim ReferenceText As String, ProjectID As String, CoordinateText As String
    Dim Latitudes() As Double, Longitudes() As Double
    Dim VertexCount As Long
    Dim CentroidLatitude As Double, CentroidLongitude As Double
    Dim AreaM2 As Double
    Dim ErrorText As String
    Dim cProject As Long, cLat As Long, cLon As Long, cLoc As Long
    Dim RowProject As String, RowLat As String, RowLon As String, RowLoc As String
    Dim MatchRow As Long
    Dim LabelText As String
    Dim ws As Worksheet
    Dim GoogleStatus As String, GoogleDetail As String, VillageMatch As String, LastFill As String

    On Error Resume Next

    '--- PART 1: watcher + drawing -------------------------------------
    P1 = "1) WATCHER: " & IIf(mAutoActive, "ON", "OFF") & _
         "   (next tick: " & IIf(mAutoNextRun > 0, Format$(mAutoNextRun, "hh:nn:ss"), "-") & ")" & vbCrLf
    If Not mAutoActive Then
        P1 = P1 & "   >> Watcher OFF hai! SolarEPC_DrawnLocationAutoStart chalao" & vbCrLf & _
                  "      ya workbook save karke DOBARA kholo (Auto_Open se chalu hota hai)." & vbCrLf
    End If

    Set Tbl = DrawnSiteTable()
    If Tbl Is Nothing Then
        P1 = P1 & "2) DRAWING_DATA / '" & SITE_TABLE & "' table NAHIN MILA." & vbCrLf & _
                  "   >> Sheet/table ka naam check karo." & vbCrLf
    Else
        P1 = P1 & "2) DRAWING_DATA." & SITE_TABLE & ": " & CStr(Tbl.ListRows.Count) & " row(s)" & vbCrLf
        For R = Tbl.ListRows.Count To 1 Step -1
            ReferenceText = RowCellText(Tbl, R, "Reference id")
            If UCase$(Left$(ReferenceText, 9)) = "SITE-MAP-" Then
                ProjectID = RowCellText(Tbl, R, "Project ID")
                CoordinateText = RowCellText(Tbl, R, "Coordinates")
                Exit For
            End If
        Next R
        If Len(ProjectID) = 0 Then
            P1 = P1 & "   >> Koi SITE-MAP-* row nahi mili - map me SAVE hua tha kya?" & vbCrLf
        Else
            P1 = P1 & "3) Latest SITE: " & ProjectID & "  (" & CStr(Len(CoordinateText)) & " char coordinates)" & vbCrLf
            If Len(CoordinateText) = 0 Then
                P1 = P1 & "   >> Coordinates cell BLANK hai - map SAVE adhoora raha." & vbCrLf
            ElseIf ParsePolygon(CoordinateText, Latitudes, Longitudes, VertexCount, ErrorText) Then
                If CentroidLocalProjection(Latitudes, Longitudes, VertexCount, _
                       CentroidLatitude, CentroidLongitude, AreaM2, ErrorText) Then
                    P1 = P1 & "   Centroid OK: " & Decimal8(CentroidLatitude) & ", " & _
                         Decimal8(CentroidLongitude) & "   (" & CStr(VertexCount) & _
                         " corners, " & Format$(AreaM2, "#,##0.00") & " m2)" & vbCrLf
                Else
                    P1 = P1 & "   >> Centroid fail: " & ErrorText & vbCrLf
                End If
            Else
                P1 = P1 & "   >> Parse fail: " & ErrorText & vbCrLf
            End If
        End If
    End If

    '--- PART 2: RESOURCE_DB + label -----------------------------------
    Set Rdb = ResourceDbTable()
    If Rdb Is Nothing Then
        P2 = "4) RESOURCE_DB ka 'resource_db' table NAHIN MILA." & vbCrLf & _
             "   >> Table ka naam exactly 'resource_db' hona chahiye." & vbCrLf
    Else
        cProject = TableColumn(Rdb, "Project ID")
        cLat = TableColumn(Rdb, "Latitude (" & ChrW(176) & ")")
        cLon = TableColumn(Rdb, "Longitude (" & ChrW(176) & ")")
        cLoc = TableColumn(Rdb, "Location")
        P2 = "4) resource_db table: " & CStr(Rdb.ListRows.Count) & " row(s); columns: " & _
             IIf(cProject > 0, "P", "-") & IIf(cLat > 0, "Lat", "-") & _
             IIf(cLon > 0, "Lon", "-") & IIf(cLoc > 0, "Loc", "-") & vbCrLf
        If cProject = 0 Or cLat = 0 Or cLon = 0 Then
            P2 = P2 & "   >> Required columns missing (Project ID / Latitude / Longitude)." & vbCrLf
        ElseIf Len(ProjectID) > 0 And CentroidLatitude <> 0 Then
            MatchRow = 0
            For R = 1 To Rdb.ListRows.Count
                RowProject = Trim$(CStr(Rdb.ListRows(R).Range.Cells(1, cProject).Value2))
                If StrComp(RowProject, ProjectID, vbTextCompare) = 0 Then
                    RowLat = Trim$(CStr(Rdb.ListRows(R).Range.Cells(1, cLat).Value2))
                    RowLon = Trim$(CStr(Rdb.ListRows(R).Range.Cells(1, cLon).Value2))
                    If (Len(RowLat) = 0 And Len(RowLon) = 0) Or _
                       (IsNumeric(RowLat) And IsNumeric(RowLon) And _
                        Abs(CDbl(RowLat) - CentroidLatitude) < 0.0000005 And _
                        Abs(CDbl(RowLon) - CentroidLongitude) < 0.0000005) Then
                        MatchRow = R
                        Exit For
                    End If
                End If
            Next R
            If MatchRow = 0 Then
                P2 = P2 & "5) >> " & ProjectID & " ki koi BLANK-lat/lon row nahi mili." & vbCrLf & _
                          "   (v2.5 ab aisi halat me row APNE AAP bana deta hai - watcher ON ho to)" & vbCrLf
            Else
                RowLat = Trim$(CStr(Rdb.ListRows(MatchRow).Range.Cells(1, cLat).Value2))
                RowLon = Trim$(CStr(Rdb.ListRows(MatchRow).Range.Cells(1, cLon).Value2))
                RowLoc = ""
                If cLoc > 0 Then RowLoc = Trim$(CStr(Rdb.ListRows(MatchRow).Range.Cells(1, cLoc).Value2))
                P2 = P2 & "5) Matched row #" & CStr(MatchRow) & ": lat=" & IIf(Len(RowLat) > 0, RowLat, "(blank)") & _
                     "  lon=" & IIf(Len(RowLon) > 0, RowLon, "(blank)") & _
                     "  Location=" & IIf(Len(RowLoc) > 0, RowLoc, "(blank)") & vbCrLf
            End If
        End If
    End If

    If CentroidLatitude <> 0 Then
        LabelText = ResourceLocationLabel(Decimal8(CentroidLatitude), Decimal8(CentroidLongitude))
        P2 = P2 & "6) Label test: " & IIf(Len(LabelText) > 0, LabelText, _
             "KOI TIER NAHIN MILA (VILLAGE_DB sheet / internet / geocoding key)") & vbCrLf
        Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
        If Not ws Is Nothing Then
            GoogleStatus = CStr(ws.Range("B14").Value2)
            GoogleDetail = CStr(ws.Range("B15").Value2)
            VillageMatch = CStr(ws.Range("B16").Value2)
            LastFill = CStr(ws.Range("B18").Value2)
            P2 = P2 & "   VILLAGE_DB match: " & IIf(Len(VillageMatch) > 0, VillageMatch, "-") & vbCrLf & _
                      "   Google tier: " & IIf(Len(GoogleStatus) > 0, GoogleStatus, "-") & _
                      IIf(Len(GoogleDetail) > 0, " (" & Left$(GoogleDetail, 120) & ")", "") & vbCrLf & _
                      "   Last auto-fill: " & IIf(Len(LastFill) > 0, LastFill, "-") & vbCrLf
        End If
    End If
    P2 = P2 & vbCrLf & "Ye report copy ho gayi hai clipboard pe."

    ResourceCopyToClipboard P1 & vbCrLf & P2
    MsgBox P1, vbInformation, "Solar EPC Debug 1/2"
    MsgBox P2, vbInformation, "Solar EPC Debug 2/2"
End Sub

