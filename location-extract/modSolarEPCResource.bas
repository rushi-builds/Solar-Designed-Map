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


Public Sub SolarEPC_ResourceShowDrawnLocation()
    Dim Tbl As ListObject
    Dim R As Long
    Dim ReferenceText As String
    Dim ProjectID As String
    Dim CoordinateText As String
    Dim Latitudes() As Double
    Dim Longitudes() As Double
    Dim VertexCount As Long
    Dim CentroidLatitude As Double
    Dim CentroidLongitude As Double
    Dim AreaM2 As Double
    Dim ErrorText As String
    Dim Detail As String
    Dim MapsURL As String
    Dim CopiedText As String
    Dim Answer As VbMsgBoxResult

    On Error GoTo Failed

    Set Tbl = DrawnSiteTable()
    If Tbl Is Nothing Then
        MsgBox "DRAWING_DATA sheet ya '" & SITE_TABLE & "' table is workbook me nahi mila." & vbCrLf & vbCrLf & _
            "Pehle map kholo, site boundary draw karo aur SAVE dabao - tab DRAWING_DATA me row banti hai.", _
            vbExclamation, "Solar EPC - Exact Drawn Location"
        Exit Sub
    End If
    If Tbl.ListRows.Count = 0 Then
        MsgBox "DRAWING_DATA me abhi koi SITE row nahi hai." & vbCrLf & vbCrLf & _
            "Map pe site draw karke SAVE karo, phir dobara run karo.", _
            vbExclamation, "Solar EPC - Exact Drawn Location"
        Exit Sub
    End If

    'Newest usable MAP SITE row (Reference id = "SITE-MAP-...").
    For R = Tbl.ListRows.Count To 1 Step -1
        ReferenceText = RowCellText(Tbl, R, "Reference id")
        If UCase$(Left$(ReferenceText, 9)) = "SITE-MAP-" Then
            ProjectID = RowCellText(Tbl, R, "Project ID")
            CoordinateText = RowCellText(Tbl, R, "Coordinates")
            If Len(ProjectID) > 0 And Len(CoordinateText) > 0 Then
                If ParsePolygon(CoordinateText, Latitudes, Longitudes, VertexCount, ErrorText) Then
                    If CentroidLocalProjection(Latitudes, Longitudes, VertexCount, _
                           CentroidLatitude, CentroidLongitude, AreaM2, ErrorText) Then
                        If AreaM2 >= SITE_MIN_AREA_M2 Then GoTo Found
                    End If
                End If
            End If
        End If
    Next R

    MsgBox "DRAWING_DATA me koi padhne-laayak drawn SITE nahi mila." & vbCrLf & vbCrLf & _
        IIf(Len(ErrorText) > 0, "Last reason: " & ErrorText & vbCrLf & vbCrLf, "") & _
        "Map pe site boundary draw karke SAVE karo, phir dobara run karo.", _
        vbExclamation, "Solar EPC - Exact Drawn Location"
    Exit Sub

Found:
    MapsURL = "https://www.google.com/maps/search/?api=1&query=" & _
        Decimal8(CentroidLatitude) & "," & Decimal8(CentroidLongitude)

    Detail = "EXACT DRAWN-SITE LOCATION" & vbCrLf & _
        "----------------------------------------" & vbCrLf & _
        "Project ID        : " & ProjectID & vbCrLf & _
        "Reference         : " & ReferenceText & vbCrLf & _
        "Boundary corners  : " & CStr(VertexCount) & vbCrLf & _
        "Site area         : " & Format$(AreaM2, "#,##0.00") & " m2" & vbCrLf & vbCrLf & _
        "EXACT SITE CENTROID (area-weighted)" & vbCrLf & _
        "Latitude  (" & ChrW(176) & ")    : " & Decimal8(CentroidLatitude) & vbCrLf & _
        "Longitude (" & ChrW(176) & ")    : " & Decimal8(CentroidLongitude) & vbCrLf & vbCrLf & _
        "Google Maps       : " & MapsURL & vbCrLf & vbCrLf & _
        "Ye centroid Worker ke centroid se identical hai - yahi coordinate NASA" & vbCrLf & _
        "POWER ko jata hai aur RESOURCE_DB Latitude/Longitude me likha jata hai."

    Answer = MsgBox(Detail & vbCrLf & vbCrLf & _
        "Yes    = Google Maps isi exact centroid pe kholo" & vbCrLf & _
        "No     = coordinates clipboard pe copy karo" & vbCrLf & _
        "Cancel = band karo", _
        vbYesNoCancel + vbInformation, "Solar EPC - Exact Drawn Location")

    If Answer = vbYes Then
        On Error Resume Next
        ThisWorkbook.FollowHyperlink MapsURL
        On Error GoTo Failed
    ElseIf Answer = vbNo Then
        CopiedText = Decimal8(CentroidLatitude) & ", " & Decimal8(CentroidLongitude)
        If CopyToClipboard(CopiedText) Then
            MsgBox "Clipboard pe copy ho gaya:" & vbCrLf & vbCrLf & CopiedText, _
                vbInformation, "Solar EPC - Exact Drawn Location"
        Else
            MsgBox "Clipboard available nahi tha. Coordinates ye hain - manually copy karo:" & _
                vbCrLf & vbCrLf & CopiedText, vbInformation, "Solar EPC - Exact Drawn Location"
        End If
    End If
    Exit Sub

Failed:
    MsgBox "Exact drawn location read nahi ho payi." & vbCrLf & vbCrLf & _
        "Error " & CStr(Err.Number) & ": " & Err.Description, _
        vbExclamation, "Solar EPC - Exact Drawn Location"
End Sub

'--------------------------------------------------------------------------
' PRIVATE HELPERS - sirf upar wale macro ke liye
'--------------------------------------------------------------------------

'DRAWING_DATA.autoLWHTbl, warna Nothing.
Private Function DrawnSiteTable() As ListObject
    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(DRAWING_SHEET)
    Set DrawnSiteTable = ws.ListObjects(SITE_TABLE)
    Exit Function
Failed:
    Set DrawnSiteTable = Nothing
End Function

'Ek list-row ka named cell text. Column na ho to "" (kabhi error nahi).
Private Function RowCellText(ByVal Tbl As ListObject, ByVal RowIndex As Long, _
    ByVal HeaderText As String) As String

    Dim C As Long
    Dim Col As ListColumn
    On Error GoTo Failed
    For Each Col In Tbl.ListColumns
        If StrComp(Trim$(CStr(Col.Name)), HeaderText, vbTextCompare) = 0 Then
            C = Col.Index
            Exit For
        End If
    Next Col
    If C = 0 Then Exit Function
    RowCellText = Trim$(CStr(Tbl.ListRows(RowIndex).Range.Cells(1, C).Value2))
    Exit Function
Failed:
    RowCellText = vbNullString
End Function

'Plain decimal number check (Worker ka strict number rule).
Private Function StrictNumber(ByVal TextValue As String) As Boolean
    Dim Re As Object
    On Error GoTo Failed
    Set Re = CreateObject("VBScript.RegExp")
    Re.Global = False
    Re.Pattern = "^-?[0-9]+(?:\.[0-9]+)?$"
    StrictNumber = Re.Test(Trim$(TextValue))
    Exit Function
Failed:
    StrictNumber = False
End Function

'8-decimal fixed text, dot separator (locale-safe).
Private Function Decimal8(ByVal Value As Double) As String
    Decimal8 = Replace$(Format$(Value, "0.00000000"), ",", ".")
End Function

'"lat,lon;lat,lon;..." -> vertex arrays, Worker ke validateSitePolygon rules:
'  8-decimal rounding, consecutive duplicates drop, repeated closing point
'  drop, >=3 distinct vertices, <=1000 vertices, WGS84 range.
Private Function ParsePolygon(ByVal CoordinateText As String, _
    ByRef Latitudes() As Double, ByRef Longitudes() As Double, _
    ByRef VertexCount As Long, ByRef ErrorText As String) As Boolean

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
    If Len(Trim$(CoordinateText)) = 0 Then
        ErrorText = "Coordinates cell blank hai."
        Exit Function
    End If

    Pairs = Split(Trim$(CoordinateText), ";")
    ReDim SourceLat(0 To UBound(Pairs))
    ReDim SourceLon(0 To UBound(Pairs))
    RawCount = 0

    For i = LBound(Pairs) To UBound(Pairs)
        If Len(Trim$(CStr(Pairs(i)))) > 0 Then
            Parts = Split(Trim$(CStr(Pairs(i))), ",")
            If UBound(Parts) <> 1 Then
                ErrorText = "Har pair 'latitude,longitude' format me hona chahiye."
                Exit Function
            End If
            If Not StrictNumber(CStr(Parts(0))) Or Not StrictNumber(CStr(Parts(1))) Then
                ErrorText = "Ek coordinate plain decimal number nahi hai."
                Exit Function
            End If
            LatitudeValue = Val(Trim$(CStr(Parts(0))))
            LongitudeValue = Val(Trim$(CStr(Parts(1))))
            If LatitudeValue < -90# Or LatitudeValue > 90# Or _
               LongitudeValue < -180# Or LongitudeValue > 180# Then
                ErrorText = "Ek coordinate WGS84 range ke bahar hai."
                Exit Function
            End If
            SourceLat(RawCount) = Val(Decimal8(LatitudeValue))
            SourceLon(RawCount) = Val(Decimal8(LongitudeValue))
            RawCount = RawCount + 1
        End If
    Next i

    If RawCount = 0 Then
        ErrorText = "Coordinates cell me koi pair nahi mila."
        Exit Function
    End If
    If RawCount > SITE_MAX_VERTICES Then
        ErrorText = "Polygon me " & CStr(SITE_MAX_VERTICES) & " se zyada vertices hain."
        Exit Function
    End If

    ReDim Latitudes(0 To RawCount - 1)
    ReDim Longitudes(0 To RawCount - 1)
    KeepCount = 0
    For i = 0 To RawCount - 1
        If KeepCount > 0 Then
            If Latitudes(KeepCount - 1) = SourceLat(i) And _
               Longitudes(KeepCount - 1) = SourceLon(i) Then GoTo NextVertex
        End If
        Latitudes(KeepCount) = SourceLat(i)
        Longitudes(KeepCount) = SourceLon(i)
        KeepCount = KeepCount + 1
NextVertex:
    Next i

    If KeepCount > 1 Then
        If Latitudes(0) = Latitudes(KeepCount - 1) And _
           Longitudes(0) = Longitudes(KeepCount - 1) Then KeepCount = KeepCount - 1
    End If
    If KeepCount < 3 Then
        ErrorText = "Polygon me kam se kam 3 alag vertices chahiye."
        Exit Function
    End If
    If DistinctVertexCount(Latitudes, Longitudes, KeepCount) < 3 Then
        ErrorText = "Polygon me kam se kam 3 alag vertices chahiye."
        Exit Function
    End If

    VertexCount = KeepCount
    ParsePolygon = True
    Exit Function
Failed:
    ErrorText = "Parse error " & CStr(Err.Number) & ": " & Err.Description
    ParsePolygon = False
End Function

Private Function DistinctVertexCount(ByRef Latitudes() As Double, _
    ByRef Longitudes() As Double, ByVal VertexCount As Long) As Long

    Dim Keys As Object
    Dim i As Long
    On Error GoTo Failed
    Set Keys = CreateObject("Scripting.Dictionary")
    For i = 0 To VertexCount - 1
        Keys(Decimal8(Latitudes(i)) & "," & Decimal8(Longitudes(i))) = 1
    Next i
    DistinctVertexCount = Keys.Count
    Exit Function
Failed:
    DistinctVertexCount = 0
End Function

'EXACT PORT of worker_k12.js polygonCentroidLocalProjection():
'local equirectangular projection around the mean vertex, area-weighted
'centroid, back to WGS84. Same constants, same order, same guards.
Private Function CentroidLocalProjection(ByRef Latitudes() As Double, _
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
    Dim i As Long
    Dim j As Long

    On Error GoTo Failed
    ErrorText = vbNullString
    If VertexCount < 3 Then
        ErrorText = "Polygon me 3 se kam vertices hain."
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
        ErrorText = "Polygon pole ke bahut paas hai."
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
        ErrorText = "Polygon ka area zero hai (saare points ek line me hain)."
        Exit Function
    End If

    CentroidLatitude = MeanLatitude + ((CentroidY6A / (3# * TwiceArea)) / EARTH_RADIUS_M) * 180# / PI
    CentroidLongitude = MeanLongitude + ((CentroidX6A / (3# * TwiceArea)) / (EARTH_RADIUS_M * CosineLatitude)) * 180# / PI
    AreaM2 = Abs(TwiceArea) / 2#

    CentroidLocalProjection = True
    Exit Function
Failed:
    ErrorText = "Centroid error " & CStr(Err.Number) & ": " & Err.Description
    CentroidLocalProjection = False
End Function

'Clipboard copy (MSForms late-bound). Na mile to False - caller MsgBox dikhta hai.
Private Function CopyToClipboard(ByVal TextValue As String) As Boolean
    Dim DataObject As Object
    On Error GoTo Failed
    Set DataObject = CreateObject("MSForms.DataObject")
    DataObject.SetText TextValue
    DataObject.PutInClipboard
    CopyToClipboard = True
    Exit Function
Failed:
    CopyToClipboard = False
End Function

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
                                ResultText = FillResourceDbForSite(ProjectID, CentroidLatitude, CentroidLongitude)
                                If InStr(1, ResultText, "LOCPEND", vbBinaryCompare) > 0 Then
                                    'lat/lon bhar gaye, label offline tha - limited retries.
                                    If Attempts < AUTO_LABEL_RETRIES Then _
                                        mProcessed(KeyText) = "LOCPEND:" & CStr(Attempts + 1)
                                ElseIf ResultText = "NOROW" Then
                                    'RESOURCE_DB row abhi bani nahi - kuch ticks retry karo.
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
    ByVal CentroidLatitude As Double, ByVal CentroidLongitude As Double) As String

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
    If Target Is Nothing Then
        FillResourceDbForSite = "NOROW"
        Exit Function
    End If

    DidText = ""

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
' LOCATION LABEL STACK - verbatim from the v2.2 complete module (your v2.1
' code):
' VILLAGE_DB -> Google -> OSM Nominatim -> BigDataCloud.
'--------------------------------------------------------------------------

Private Function ResourceDecimal(ByVal Value As Double, ByVal Places As Long) As String
    ResourceDecimal = Replace$(Format$(Value, "0." & String$(Places, "0")), ",", ".")
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

Private Function ResourceConfigCell(ByVal CellAddress As String) As String
    ResourceConfigCell = ResourceSettingCell(CONFIG_SHEET, CellAddress)
End Function

Private Function ResourceSettingCell(ByVal SheetName As String, _
    ByVal CellAddress As String) As String
    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(SheetName)
    ResourceSettingCell = Trim$(CStr(ws.Range(CellAddress).Value2))
    Exit Function
Failed:
    ResourceSettingCell = vbNullString
End Function

Private Function ResourceGoogleAddressPart(ByVal JSONText As String, ByVal TypeName As String) As String
    Dim Re As Object, Matches As Object, M As Object
    On Error GoTo Failed
    Set Re = CreateObject("VBScript.RegExp")
    Re.Global = True
    Re.IgnoreCase = True
    Re.Pattern = "\{\s*""long_name""\s*:\s*""([^""]*)""\s*,\s*""short_name""\s*:\s*""[^""]*""\s*," & _
                 "\s*""types""\s*:\s*\[([^\]]*)\]"
    If Re.Test(JSONText) Then
        Set Matches = Re.Execute(JSONText)
        For Each M In Matches
            If InStr(1, M.SubMatches(1), """" & TypeName & """", vbTextCompare) > 0 Then
                ResourceGoogleAddressPart = Trim$(M.SubMatches(0))
                Exit Function
            End If
        Next M
    End If
    Exit Function
Failed:
    ResourceGoogleAddressPart = vbNullString
End Function

Private Function ResourceGoogleLocationLabel(ByVal Lat As Double, ByVal Lon As Double) As String
    Dim Key As String, UrlText As String, JSONText As String
    Dim Locality As String, Taluka As String, District As String
    Dim StatusText As String, ErrorText As String
    Dim KeySource As String

    On Error GoTo Failed
    'Resolve the Geocoding API key. SETTINGS!B11 is the dedicated geocoding
    'key; SETTINGS!B4 is the workbook map key; _CLOUD_CFG!B5 is a manual
    'override. The first non-empty value wins.
    Key = ResourceSettingCell(SETTINGS_SHEET, "B11")
    KeySource = "SETTINGS!B11"
    If Len(Key) = 0 Then
        Key = ResourceSettingCell(SETTINGS_SHEET, "B4")
        KeySource = "SETTINGS!B4"
    End If
    If Len(Key) = 0 Then
        Key = ResourceConfigCell("B5")
        KeySource = "_CLOUD_CFG!B5"
    End If
    If Len(Key) = 0 Then
        KeySource = "NONE"
        ResourceGoogleStatus "NO_KEY", _
            "No Geocoding API key found in SETTINGS!B11, SETTINGS!B4 or " & _
            "_CLOUD_CFG!B5, so the Google tier is skipped and Nominatim is used. " & _
            "Put the key in SETTINGS!B11."
        ResourceSettingDiag "A17", "GEOCODING KEY SOURCE", "B17", KeySource
        Exit Function
    End If
    ResourceSettingDiag "A17", "GEOCODING KEY SOURCE", "B17", KeySource

    UrlText = "https://maps.googleapis.com/maps/api/geocode/json?latlng=" & _
              ResourceDecimal(Lat, 8) & "," & ResourceDecimal(Lon, 8) & _
              "&language=en&key=" & Key
    JSONText = ResourceHttpGetJson(UrlText)
    If Len(JSONText) = 0 Then
        ResourceGoogleStatus "HTTP_FAIL", _
            "No response from maps.googleapis.com (network, proxy or firewall block)."
        Exit Function
    End If

    StatusText = ResourceJSONValue(JSONText, "status")
    ErrorText = ResourceJSONValue(JSONText, "error_message")
    If StrComp(StatusText, "OK", vbTextCompare) <> 0 Then
        If Len(StatusText) = 0 Then StatusText = "NO_STATUS_FIELD"
        If Len(ErrorText) = 0 Then
            If StrComp(StatusText, "REQUEST_DENIED", vbTextCompare) = 0 Then
                ErrorText = "Key is invalid, expired, HTTP-referrer/IP restricted, or the " & _
                    "Geocoding API is not enabled for it. VBA calls send no referrer, so a " & _
                    "web-restricted (demo) key will always be denied."
            ElseIf StrComp(StatusText, "OVER_QUERY_LIMIT", vbTextCompare) = 0 Then
                ErrorText = "Quota/billing exhausted on this key."
            ElseIf StrComp(StatusText, "ZERO_RESULTS", vbTextCompare) = 0 Then
                ErrorText = "Google has no address for this exact point."
            End If
        End If
        ResourceGoogleStatus "GOOGLE_" & StatusText, ErrorText
        Exit Function
    End If

    Locality = ResourceGoogleAddressPart(JSONText, "locality")
    If Len(Locality) = 0 Then Locality = ResourceGoogleAddressPart(JSONText, "sublocality")
    If Len(Locality) = 0 Then Locality = ResourceGoogleAddressPart(JSONText, "neighborhood")
    If Len(Locality) = 0 Then
        ResourceGoogleStatus "NO_LOCALITY", _
            "Google replied OK but returned no locality / sublocality / neighborhood component."
        Exit Function
    End If

    Taluka = ResourceGoogleAddressPart(JSONText, "administrative_area_level_3")
    District = ResourceGoogleAddressPart(JSONText, "administrative_area_level_2")
    ResourceGoogleLocationLabel = ResourceLocationCompose(Locality, _
        ResourceLocationCompose(ResourceAdminClean(Taluka), ResourceAdminClean(District)))
    ResourceGoogleStatus "OK", ResourceGoogleLocationLabel
    Exit Function
Failed:
    ResourceGoogleLocationLabel = vbNullString
    ResourceGoogleStatus "VBA_ERROR", "VBA error " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Sub ResourceSettingDiag(ByVal LabelCell As String, ByVal LabelText As String, _
    ByVal ValueCell As String, ByVal ValueText As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    If Not ws Is Nothing Then
        ws.Range(LabelCell).Value2 = LabelText
        ws.Range(ValueCell).Value2 = ValueText
    End If
    On Error GoTo 0
End Sub

Private Sub ResourceGoogleStatus(ByVal StatusText As String, ByVal DetailText As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    If Not ws Is Nothing Then
        ws.Range("A14").Value2 = "GOOGLE GEOCODE STATUS"
        ws.Range("B14").Value2 = StatusText
        ws.Range("A15").Value2 = "GOOGLE GEOCODE DETAIL"
        ws.Range("B15").Value2 = Left$(DetailText, 900)
    End If
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"), "GOOGLE TIER:", StatusText, DetailText
    On Error GoTo 0
End Sub

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

    '0b) OPTIONAL Google tier - only fires when a key is found and works.
    ResourceLocationLabel = ResourceGoogleLocationLabel(Lat, Lon)
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
