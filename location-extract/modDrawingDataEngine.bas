Option Explicit

'===========================================================
' SOLAR EPC - DRAWING DATA ENGINE
'
' Sheet:
'   DRAWING_DATA
'
' Tables:
'   autoLWHTbl     = Site Dimensions
'   tblDrawingData = Obstacle Dimensions
'
' HISTORY MODE:
' Every SAVE uses a new blank row.
' Existing records are never automatically deleted.
'===========================================================

Private Const DRAWING_SHEET As String = "DRAWING_DATA"
Private Const SITE_TABLE As String = "autoLWHTbl"
Private Const OBSTACLE_TABLE As String = "tblDrawingData"


'===========================================================
' TEST ENGINE
'===========================================================

Public Sub TestDrawingDataEngine()

    Dim ws As Worksheet
    Dim SiteTbl As ListObject
    Dim ObstacleTbl As ListObject

    On Error GoTo Failed

    Set ws = ThisWorkbook.Worksheets(DRAWING_SHEET)
    Set SiteTbl = ws.ListObjects(SITE_TABLE)
    Set ObstacleTbl = ws.ListObjects(OBSTACLE_TABLE)

    MsgBox _
        "Drawing Data Engine is ready." & vbCrLf & vbCrLf & _
        "Sheet: " & ws.Name & vbCrLf & _
        "Site Table: " & SiteTbl.Name & vbCrLf & _
        "Obstacle Table: " & ObstacleTbl.Name, _
        vbInformation, _
        "Drawing Data Engine"

    Exit Sub

Failed:

    MsgBox _
        "Drawing Data Engine test failed." & _
        vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbCritical, _
        "Drawing Data Engine"

End Sub


'===========================================================
' GET PROJECT ID
'===========================================================

Public Function DrawingData_GetProjectID() As String

    On Error GoTo Failed

    DrawingData_GetProjectID = Trim$(CStr( _
        ThisWorkbook.Worksheets("INPUT"). _
        Range("C5").Value2))

    Exit Function

Failed:

    DrawingData_GetProjectID = vbNullString

End Function


'===========================================================
' NORMALIZE HEADER
'===========================================================

Private Function NormalizeDrawingHeader( _
    ByVal HeaderText As String) As String

    Dim s As String

    s = LCase$(Trim$(HeaderText))

    'Convert m  to m2.
    s = Replace$(s, ChrW(178), "2")

    'Handle damaged encoding.
    s = Replace$(s, "?", "2")

    s = Replace$(s, vbCr, " ")
    s = Replace$(s, vbLf, " ")
    s = Replace$(s, vbTab, " ")

    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace$(s, "  ", " ")
    Loop

    'Allow Reference and Reference id.
    If s = "reference id" Then
        s = "reference"
    End If

    'Allow Remark and Remarks.
    If s = "remarks" Then
        s = "remark"
    End If

    NormalizeDrawingHeader = Trim$(s)

End Function


'===========================================================
' FIND TABLE COLUMN
'===========================================================

Private Function GetTableColumnNumber( _
    ByVal Tbl As ListObject, _
    ByVal HeaderName As String) As Long

    Dim ColumnItem As ListColumn
    Dim WantedHeader As String
    Dim CurrentHeader As String

    WantedHeader = NormalizeDrawingHeader(HeaderName)

    For Each ColumnItem In Tbl.ListColumns

        CurrentHeader = NormalizeDrawingHeader( _
                            CStr(ColumnItem.Name))

        If StrComp( _
                CurrentHeader, _
                WantedHeader, _
                vbTextCompare) = 0 Then

            GetTableColumnNumber = ColumnItem.Index
            Exit Function

        End If

    Next ColumnItem

    GetTableColumnNumber = 0

End Function


'===========================================================
' VALIDATE SITE TABLE
'===========================================================

Private Function ValidateSiteTable( _
    ByVal Tbl As ListObject) As Boolean

    Dim Headers As Variant
    Dim i As Long

    Headers = Array( _
        "Project ID", _
        "Project Location", _
        "Length (m)", _
        "Width (m)", _
        "Height (m)", _
        "Site Area (m2)", _
        "Perimeter (m)", _
        "Point Count", _
        "Source", _
        "Reference id", _
        "Coordinates", _
        "Remarks" _
    )

    For i = LBound(Headers) To UBound(Headers)

        If GetTableColumnNumber( _
                Tbl, CStr(Headers(i))) = 0 Then

            MsgBox _
                "Required site column not found:" & _
                vbCrLf & CStr(Headers(i)), _
                vbCritical, _
                "DRAWING_DATA Structure Error"

            Exit Function

        End If

    Next i

    ValidateSiteTable = True

End Function


'===========================================================
' VALIDATE OBSTACLE TABLE
'===========================================================

Private Function ValidateObstacleTable( _
    ByVal Tbl As ListObject) As Boolean

    Dim Headers As Variant
    Dim i As Long

    Headers = Array( _
        "Project ID", _
        "Obstacle No", _
        "Obstacle Type", _
        "Length (m)", _
        "Width (m)", _
        "Height (m)", _
        "Clearance (m)", _
        "Area (m2)", _
        "Perimeter (m)", _
        "Source", _
        "Point Count", _
        "Reference id", _
        "Coordinates", _
        "Remarks" _
    )

    For i = LBound(Headers) To UBound(Headers)

        If GetTableColumnNumber( _
                Tbl, CStr(Headers(i))) = 0 Then

            MsgBox _
                "Required obstacle column not found:" & _
                vbCrLf & CStr(Headers(i)), _
                vbCritical, _
                "DRAWING_DATA Structure Error"

            Exit Function

        End If

    Next i

    ValidateObstacleTable = True

End Function


'===========================================================
' GET FIRST BLANK TABLE ROW
'
' Existing records are never overwritten.
'===========================================================

Private Function GetBlankOrNewTableRow( _
    ByVal Tbl As ListObject, _
    ByVal ProjectColumn As Long) As ListRow

    Dim ProjectValues As Variant
    Dim RowNumber As Long

    'Read the complete Project ID column once. This is much faster than
    'making a separate Excel object-model call for every existing row.
    If ProjectColumn > 0 And Tbl.ListRows.Count > 0 Then
        ProjectValues = Tbl.ListColumns(ProjectColumn).DataBodyRange.Value2

        If Tbl.ListRows.Count = 1 Then
            If Len(Trim$(CStr(ProjectValues))) = 0 Then
                Set GetBlankOrNewTableRow = Tbl.ListRows(1)
                Exit Function
            End If
        Else
            For RowNumber = 1 To UBound(ProjectValues, 1)
                If Len(Trim$(CStr(ProjectValues(RowNumber, 1)))) = 0 Then
                    Set GetBlankOrNewTableRow = Tbl.ListRows(RowNumber)
                    Exit Function
                End If
            Next RowNumber
        End If
    End If

    Set GetBlankOrNewTableRow = Tbl.ListRows.Add

End Function


'===========================================================
' SAVE SITE
'
' Parameters after ReferenceID:
'   Remarks
'   Coordinates
'
' This order keeps the existing manual-form calls working.
'===========================================================

Public Function DrawingData_SaveSite( _
    ByVal ProjectID As String, _
    ByVal LengthM As Double, _
    ByVal WidthM As Double, _
    ByVal HeightM As Double, _
    ByVal SiteAreaM2 As Double, _
    ByVal PerimeterM As Double, _
    ByVal PointCount As Variant, _
    ByVal SourceName As String, _
    ByVal ReferenceID As String, _
    Optional ByVal Remarks As String = "", _
    Optional ByVal Coordinates As String = "" _
) As Boolean

    Dim ws As Worksheet
    Dim Tbl As ListObject
    Dim TargetRow As ListRow

    Dim cProject As Long
    Dim cLocation As Long
    Dim cLength As Long
    Dim cWidth As Long
    Dim cHeight As Long
    Dim cArea As Long
    Dim cPerimeter As Long
    Dim cPoints As Long
    Dim cSource As Long
    Dim cReference As Long
    Dim cCoordinates As Long
    Dim cRemarks As Long

    On Error GoTo Failed

    ProjectID = Trim$(ProjectID)

    If Len(ProjectID) = 0 Then

        Err.Raise _
            vbObjectError + 1001, _
            "DrawingData_SaveSite", _
            "Project ID cannot be blank."

    End If

    Set ws = ThisWorkbook.Worksheets(DRAWING_SHEET)
    Set Tbl = ws.ListObjects(SITE_TABLE)

    If Not ValidateSiteTable(Tbl) Then
        Exit Function
    End If

    cProject = GetTableColumnNumber(Tbl, "Project ID")
    cLocation = GetTableColumnNumber(Tbl, "Project Location")
    cLength = GetTableColumnNumber(Tbl, "Length (m)")
    cWidth = GetTableColumnNumber(Tbl, "Width (m)")
    cHeight = GetTableColumnNumber(Tbl, "Height (m)")
    cArea = GetTableColumnNumber(Tbl, "Site Area (m2)")
    cPerimeter = GetTableColumnNumber(Tbl, "Perimeter (m)")
    cPoints = GetTableColumnNumber(Tbl, "Point Count")
    cSource = GetTableColumnNumber(Tbl, "Source")
    cReference = GetTableColumnNumber(Tbl, "Reference id")
    cCoordinates = GetTableColumnNumber(Tbl, "Coordinates")
    cRemarks = GetTableColumnNumber(Tbl, "Remarks")

    Set TargetRow = GetBlankOrNewTableRow( _
                        Tbl, cProject)

    If TargetRow Is Nothing Then

        Err.Raise _
            vbObjectError + 1002, _
            "DrawingData_SaveSite", _
            "Could not create a site table row."

    End If

    With TargetRow.Range

        .Cells(1, cProject).Value2 = ProjectID

        '=====================================================
        ' PROJECT LOCATION SNAPSHOT
        ' Capture INPUT!C7 only for THIS new site row.
        ' Existing DRAWING_DATA rows are never updated later.
        '=====================================================
        .Cells(1, cLocation).Value2 = _
            ThisWorkbook.Worksheets("INPUT").Range("C7").Value2

        .Cells(1, cLength).Value2 = LengthM
        .Cells(1, cWidth).Value2 = WidthM
        .Cells(1, cHeight).Value2 = HeightM
        .Cells(1, cArea).Value2 = SiteAreaM2
        .Cells(1, cPerimeter).Value2 = PerimeterM

        If IsEmpty(PointCount) Or _
           Len(Trim$(CStr(PointCount))) = 0 Then

            .Cells(1, cPoints).ClearContents

        Else

            .Cells(1, cPoints).Value2 = PointCount

        End If

        .Cells(1, cSource).Value2 = _
            UCase$(Trim$(SourceName))

        .Cells(1, cReference).Value2 = ReferenceID
        .Cells(1, cCoordinates).Value2 = Coordinates
        .Cells(1, cRemarks).Value2 = Remarks

        .Cells(1, cLength).NumberFormat = "0.00"
        .Cells(1, cWidth).NumberFormat = "0.00"
        .Cells(1, cHeight).NumberFormat = "0.00"
        .Cells(1, cArea).NumberFormat = "0.00"
        .Cells(1, cPerimeter).NumberFormat = "0.00"

    End With

    DrawingData_SaveSite = True
    Exit Function

Failed:

    DrawingData_SaveSite = False

    MsgBox _
        "Site data could not be saved." & _
        vbCrLf & vbCrLf & _
        "Project: " & ProjectID & vbCrLf & _
        "Error " & Err.Number & ": " & _
        Err.Description, _
        vbCritical, _
        "DRAWING_DATA Error"

End Function


'===========================================================
' SAVE OBSTACLE
'
' Parameters after ReferenceID:
'   Remarks
'   Coordinates
'===========================================================

Public Function DrawingData_SaveObstacle( _
    ByVal ProjectID As String, _
    ByVal ObstacleNo As String, _
    ByVal SourceName As String, _
    ByVal ObstacleType As String, _
    ByVal LengthM As Double, _
    ByVal WidthM As Double, _
    ByVal HeightM As Double, _
    ByVal ClearanceM As Double, _
    ByVal AreaM2 As Double, _
    ByVal PerimeterM As Double, _
    ByVal PointCount As Variant, _
    ByVal ReferenceID As String, _
    Optional ByVal Remarks As String = "", _
    Optional ByVal Coordinates As String = "" _
) As Boolean

    Dim ws As Worksheet
    Dim Tbl As ListObject
    Dim TargetRow As ListRow

    Dim cProject As Long
    Dim cObstacleNo As Long
    Dim cType As Long
    Dim cLength As Long
    Dim cWidth As Long
    Dim cHeight As Long
    Dim cClearance As Long
    Dim cArea As Long
    Dim cPerimeter As Long
    Dim cSource As Long
    Dim cPoints As Long
    Dim cReference As Long
    Dim cCoordinates As Long
    Dim cRemarks As Long

    On Error GoTo Failed

    ProjectID = Trim$(ProjectID)

    If Len(ProjectID) = 0 Then

        Err.Raise _
            vbObjectError + 1101, _
            "DrawingData_SaveObstacle", _
            "Project ID cannot be blank."

    End If

    Set ws = ThisWorkbook.Worksheets(DRAWING_SHEET)
    Set Tbl = ws.ListObjects(OBSTACLE_TABLE)

    If Not ValidateObstacleTable(Tbl) Then
        Exit Function
    End If

    cProject = GetTableColumnNumber(Tbl, "Project ID")
    cObstacleNo = GetTableColumnNumber(Tbl, "Obstacle No")
    cType = GetTableColumnNumber(Tbl, "Obstacle Type")
    cLength = GetTableColumnNumber(Tbl, "Length (m)")
    cWidth = GetTableColumnNumber(Tbl, "Width (m)")
    cHeight = GetTableColumnNumber(Tbl, "Height (m)")
    cClearance = GetTableColumnNumber(Tbl, "Clearance (m)")
    cArea = GetTableColumnNumber(Tbl, "Area (m2)")
    cPerimeter = GetTableColumnNumber(Tbl, "Perimeter (m)")
    cSource = GetTableColumnNumber(Tbl, "Source")
    cPoints = GetTableColumnNumber(Tbl, "Point Count")
    cReference = GetTableColumnNumber(Tbl, "Reference id")
    cCoordinates = GetTableColumnNumber(Tbl, "Coordinates")
    cRemarks = GetTableColumnNumber(Tbl, "Remarks")

    Set TargetRow = GetBlankOrNewTableRow( _
                        Tbl, cProject)

    If TargetRow Is Nothing Then

        Err.Raise _
            vbObjectError + 1102, _
            "DrawingData_SaveObstacle", _
            "Could not create an obstacle table row."

    End If

    With TargetRow.Range

        .Cells(1, cProject).Value2 = ProjectID
        .Cells(1, cObstacleNo).Value2 = ObstacleNo
        .Cells(1, cType).Value2 = ObstacleType
        .Cells(1, cLength).Value2 = LengthM
        .Cells(1, cWidth).Value2 = WidthM
        .Cells(1, cHeight).Value2 = HeightM
        .Cells(1, cClearance).Value2 = ClearanceM
        .Cells(1, cArea).Value2 = AreaM2
        .Cells(1, cPerimeter).Value2 = PerimeterM

        .Cells(1, cSource).Value2 = _
            UCase$(Trim$(SourceName))

        If IsEmpty(PointCount) Or _
           Len(Trim$(CStr(PointCount))) = 0 Then

            .Cells(1, cPoints).ClearContents

        Else

            .Cells(1, cPoints).Value2 = PointCount

        End If

        .Cells(1, cReference).Value2 = ReferenceID
        'Coordinates = 12th pipe-field (1-based) = ObstacleRecord(11).
        'Blank-only: never clear or overwrite a cell that already has text.
        If Len(Trim$(Coordinates)) > 0 Then
            If Len(Trim$(CStr(.Cells(1, cCoordinates).Value2 & ""))) = 0 Then
                If Not .Cells(1, cCoordinates).HasFormula Then
                    .Cells(1, cCoordinates).Value2 = Trim$(Coordinates)
                End If
            End If
        End If
        .Cells(1, cRemarks).Value2 = Remarks

        .Cells(1, cLength).NumberFormat = "0.00"
        .Cells(1, cWidth).NumberFormat = "0.00"
        .Cells(1, cHeight).NumberFormat = "0.00"
        .Cells(1, cClearance).NumberFormat = "0.00"
        .Cells(1, cArea).NumberFormat = "0.00"
        .Cells(1, cPerimeter).NumberFormat = "0.00"

    End With

    DrawingData_SaveObstacle = True
    Exit Function

Failed:

    DrawingData_SaveObstacle = False

    MsgBox _
        "Obstacle data could not be saved." & _
        vbCrLf & vbCrLf & _
        "Project: " & ProjectID & vbCrLf & _
        "Obstacle: " & ObstacleNo & vbCrLf & _
        "Error " & Err.Number & ": " & _
        Err.Description, _
        vbCritical, _
        "DRAWING_DATA Error"

End Function


'===========================================================
' IMPORT MAP PAYLOAD
'
' New SITE format:
'
' 0 = SITE
' 1 = Length
' 2 = Width
' 3 = Height
' 4 = Area
' 5 = Perimeter
' 6 = Point Count
' 7 = Reference ID
' 8 = Coordinates
' 9 = Remarks
'
' New OBSTACLE format:
'
' 0  = OBSTACLE
' 1  = Obstacle No
' 2  = Obstacle Type
' 3  = Length
' 4  = Width
' 5  = Height
' 6  = Clearance
' 7  = Area
' 8  = Perimeter
' 9  = Point Count
' 10 = Reference ID
' 11 = Coordinates   (12th pipe-field, 1-based)
' 12 = Remarks
'===========================================================

Public Function DrawingData_ImportFromMapPayload( _
    ByVal PayloadText As String) As Boolean

    Dim Lines As Variant
    Dim Parts As Variant
    Dim SiteRecord As Variant
    Dim ObstacleRecords As Collection
    Dim ObstacleRecord As Variant

    Dim i As Long
    Dim HasSiteRecord As Boolean

    Dim PayloadProjectID As String
    Dim ProjectID As String

    Dim SiteCoordinates As String
    Dim SiteRemarks As String

    Dim ObstacleCoordinates As String
    Dim ObstacleRemarks As String

    Dim ws As Worksheet
    Dim SiteTbl As ListObject
    Dim ObstacleTbl As ListObject

    On Error GoTo Failed

    PayloadText = Replace$( _
                        PayloadText, _
                        vbCrLf, _
                        vbLf)

    PayloadText = Replace$( _
                        PayloadText, _
                        vbCr, _
                        vbLf)

    PayloadText = Trim$(PayloadText)

    If Left$( _
            PayloadText, _
            Len("SOLAR_EPC_MAP_SAVE_V1")) <> _
            "SOLAR_EPC_MAP_SAVE_V1" Then

        Err.Raise _
            vbObjectError + 1201, _
            "DrawingData_ImportFromMapPayload", _
            "Invalid Solar EPC map payload."

    End If

    Lines = Split(PayloadText, vbLf)

    Set ObstacleRecords = New Collection

    For i = 1 To UBound(Lines)

        If Len(Trim$(CStr(Lines(i)))) > 0 Then

            Parts = Split(CStr(Lines(i)), "|")

            Select Case UCase$( _
                Trim$(CStr(Parts(0))))

                Case "PROJECT"

                    If UBound(Parts) >= 1 Then
                        PayloadProjectID = _
                            Trim$(CStr(Parts(1)))
                    End If

                Case "SITE"

                    'Old payload needs indexes 0-8.
                    'New payload uses indexes 0-9.
                    If UBound(Parts) < 8 Then

                        Err.Raise _
                            vbObjectError + 1202, _
                            "DrawingData_ImportFromMapPayload", _
                            "SITE record is incomplete."

                    End If

                    SiteRecord = Parts
                    HasSiteRecord = True

                Case "OBSTACLE"

                    'Old payload needs indexes 0-11.
                    'New payload uses indexes 0-12.
                    If UBound(Parts) < 11 Then

                        Err.Raise _
                            vbObjectError + 1203, _
                            "DrawingData_ImportFromMapPayload", _
                            "OBSTACLE record is incomplete."

                    End If

                    ObstacleRecords.Add Parts

            End Select

        End If

    Next i

    If Not HasSiteRecord Then

        Err.Raise _
            vbObjectError + 1204, _
            "DrawingData_ImportFromMapPayload", _
            "SITE record was not found."

    End If

    ProjectID = DrawingData_GetProjectID()

    If Len(ProjectID) = 0 Then
        ProjectID = PayloadProjectID
    End If

    If Len(ProjectID) = 0 Then

        Err.Raise _
            vbObjectError + 1205, _
            "DrawingData_ImportFromMapPayload", _
            "Project ID is blank. " & _
            "Enter a project name in INPUT!C5."

    End If

    Set ws = ThisWorkbook.Worksheets(DRAWING_SHEET)
    Set SiteTbl = ws.ListObjects(SITE_TABLE)
    Set ObstacleTbl = ws.ListObjects(OBSTACLE_TABLE)

    If Not ValidateSiteTable(SiteTbl) Then
        Exit Function
    End If

    If Not ValidateObstacleTable(ObstacleTbl) Then
        Exit Function
    End If

    '=======================================================
    ' READ SITE COORDINATES AND REMARKS
    '=======================================================

    If UBound(SiteRecord) >= 9 Then

        SiteCoordinates = CStr(SiteRecord(8))
        SiteRemarks = CStr(SiteRecord(9))

    Else

        'Old payload did not contain Coordinates.
        SiteCoordinates = vbNullString
        SiteRemarks = CStr(SiteRecord(8))

    End If

    '=======================================================
    ' SAVE NEW SITE ROW
    '=======================================================

    If Not DrawingData_SaveSite( _
            ProjectID, _
            MapPayloadNumber(SiteRecord(1)), _
            MapPayloadNumber(SiteRecord(2)), _
            MapPayloadNumber(SiteRecord(3)), _
            MapPayloadNumber(SiteRecord(4)), _
            MapPayloadNumber(SiteRecord(5)), _
            CLng(MapPayloadNumber(SiteRecord(6))), _
            "MAP", _
            CStr(SiteRecord(7)), _
            SiteRemarks, _
            SiteCoordinates) Then

        Exit Function

    End If

    '=======================================================
    ' SAVE NEW OBSTACLE ROWS
    ' 12th pipe-field (1-based) = Coordinates = index 11
    '=======================================================

    For Each ObstacleRecord In ObstacleRecords

        If UBound(ObstacleRecord) >= 12 Then

            ObstacleCoordinates = _
                Trim$(CStr(ObstacleRecord(11)))

            ObstacleRemarks = _
                CStr(ObstacleRecord(12))

        ElseIf UBound(ObstacleRecord) >= 11 Then

            'New coords field present, remarks may be missing.
            ObstacleCoordinates = _
                Trim$(CStr(ObstacleRecord(11)))

            ObstacleRemarks = vbNullString

        Else

            'Old payload did not contain Coordinates.
            ObstacleCoordinates = vbNullString
            ObstacleRemarks = CStr(ObstacleRecord(11))

        End If

        If Not DrawingData_SaveObstacle( _
                ProjectID, _
                CStr(ObstacleRecord(1)), _
                "MAP", _
                CStr(ObstacleRecord(2)), _
                MapPayloadNumber(ObstacleRecord(3)), _
                MapPayloadNumber(ObstacleRecord(4)), _
                MapPayloadNumber(ObstacleRecord(5)), _
                MapPayloadNumber(ObstacleRecord(6)), _
                MapPayloadNumber(ObstacleRecord(7)), _
                MapPayloadNumber(ObstacleRecord(8)), _
                CLng(MapPayloadNumber( _
                    ObstacleRecord(9))), _
                CStr(ObstacleRecord(10)), _
                ObstacleRemarks, _
                ObstacleCoordinates) Then

            Exit Function

        End If

    Next ObstacleRecord

    DrawingData_ImportFromMapPayload = True
    Exit Function

Failed:

    DrawingData_ImportFromMapPayload = False

    MsgBox _
        "Map data could not be imported into DRAWING_DATA." & _
        vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & _
        Err.Description, _
        vbCritical, _
        "Solar EPC Map Import"

End Function


'===========================================================
' PARSE JAVASCRIPT NUMBER
'===========================================================

Private Function MapPayloadNumber( _
    ByVal ValueText As Variant) As Double

    Dim s As String

    s = Trim$(CStr(ValueText))

    If Len(s) = 0 Then
        MapPayloadNumber = 0#
    Else
        MapPayloadNumber = Val(s)
    End If

End Function


'===========================================================
' CLEAR SELECTED DRAWING DATA
'===========================================================
Public Sub DrawingData_ClearSelected( _
    ByVal ClearSite As Boolean, _
    ByVal ClearObstacle As Boolean)

    Dim ws As Worksheet
    Dim Tbl As ListObject

    On Error GoTo Failed

    Set ws = ThisWorkbook.Worksheets(DRAWING_SHEET)

    If ClearSite Then
        Set Tbl = ws.ListObjects(SITE_TABLE)
        If Not Tbl.DataBodyRange Is Nothing Then
            Tbl.DataBodyRange.ClearContents
        End If
    End If

    If ClearObstacle Then
        Set Tbl = ws.ListObjects(OBSTACLE_TABLE)
        If Not Tbl.DataBodyRange Is Nothing Then
            Tbl.DataBodyRange.ClearContents
        End If
    End If

    Exit Sub

Failed:
    Err.Raise _
        vbObjectError + 1301, _
        "DrawingData_ClearSelected", _
        "Could not clear the selected Drawing Data." & _
        vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description

End Sub
