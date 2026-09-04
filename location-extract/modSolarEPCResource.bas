Attribute VB_Name = "modSolarEPCDrawnLocation"
Option Explicit

'==========================================================================
' SOLAR EPC - EXACT DRAWN-SITE LOCATION  (MINIMAL STANDALONE BUILD)
' Version 2.2-min
'
' THIS MODULE DOES EXACTLY ONE THING:
'   SolarEPC_ResourceShowDrawnLocation
'       Shows the EXACT location of the site you drew on the map:
'         - the area-weighted CENTROID (latitude/longitude, 8 decimals)
'         - how many corners you drew and the exact site area (m2)
'         - a Google Maps link for that exact point (Yes)
'         - or the coordinates copied to the clipboard (No)
'
' NOTHING ELSE. No NASA call, no Worker call, no cloud configuration,
' no reverse geocoding, no RESOURCE_DB write, no scheduler, no internet.
' It only READS DRAWING_DATA.autoLWHTbl - nothing in the workbook is
' modified by this module.
'
' WHY A SEPARATE MODULE NAME (modSolarEPCDrawnLocation):
'   Your workbook already contains modSolarEPCResource (the NASA POWER
'   resource module). Importing this file ADDS a second module and cannot
'   collide with it - the module name and the only public macro name are
'   both unique. Do not remove your existing modSolarEPCResource.
'   (If you earlier imported the FULL v2.2 build of modSolarEPCResource
'   that also contained SolarEPC_ResourceShowDrawnLocation, remove that
'   copy first - two copies of the same public macro name cannot coexist.)
'
' THE MATH IS THE WORKER'S MATH:
'   The centroid below is a statement-by-statement port of
'   polygonCentroidLocalProjection() in worker_k12.js - same Earth radius
'   (6378137 m), same 8-decimal vertex rounding, same duplicate and
'   closing-point rules, same minimum area (0.01 m2), same projection.
'   So the number this macro shows is byte-for-byte the number the Worker
'   sends to NASA POWER and writes into RESOURCE_DB Latitude/Longitude.
'   Verified identical on 5000 generated polygons + 7 hand-built edge
'   cases (location-extract/verify/run_parity_check.sh).
'
' HOW TO USE
'   1. VBE (Alt+F11) -> right-click project -> Import File... -> this file
'   2. Map pe site draw karo aur SAVE dabao (DRAWING_DATA row banti hai)
'   3. Alt+F8 -> SolarEPC_ResourceShowDrawnLocation -> Run
'==========================================================================

Private Const DRAWING_SHEET As String = "DRAWING_DATA"
Private Const SITE_TABLE As String = "autoLWHTbl"
'Worker constants - do not change (see header).
Private Const EARTH_RADIUS_M As Double = 6378137#
Private Const SITE_MAX_VERTICES As Long = 1000
Private Const SITE_MIN_AREA_M2 As Double = 0.01

'--------------------------------------------------------------------------
' THE ONLY PUBLIC MACRO
'--------------------------------------------------------------------------
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
