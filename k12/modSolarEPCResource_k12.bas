Option Explicit

'==========================================================================
' SOLAR EPC - NASA POWER HOURLY RESOURCE MODULE v1.9.2 (RANGE12 + LOCATION AUTO-FILL)
'
' - Reads only the SITE row in DRAWING_DATA.autoLWHTbl.
' - Never reads tblDrawingData (obstacles) for the resource centroid.
' - Sends the complete SITE polygon to the authenticated Worker.
' - The Worker validates the polygon, calculates its area-weighted centroid,
'   downloads the explicitly configured resource period in hourly chunks,
'   validates and stores every hourly record, and returns a project summary.
' - T2M minimum/maximum are derived from validated hourly T2M because NASA's
'   Hourly API rejects T2M_MIN and T2M_MAX.
'
' v1.9.2 CHANGE (Location auto-fill):
' - RESOURCE_DB ka descriptive "Location" column ab automatically fill hota hai
'   exact centroid coordinates se: free reverse-geocoding (OSM Nominatim,
'   fallback BigDataCloud, dono bina API key) -> "Town, District" format
'   (e.g. "Mahabaleshwar, Satara", "Phaltan, Satara", "Baramati, Pune").
' - Sirf BLANK cells fill hote hain; manually typed Location kabhi overwrite
'   nahi hota. Coordinates/weather values/status kuch change nahi hote.
' - Network/service fail ho toh Location blank rehta hai (koi error nahi,
'   data import block nahi hota).
' - New macro: SolarEPC_ResourceFillLocations -> existing rows ki blank
'   Location cells ek saath fill kar deta hai.
'
' v1.9.1 CHANGE (RESOURCE_DB Location column):
' - RESOURCE_DB may now include an additional descriptive "Location" column.
'   Required columns are matched BY NAME (Project ID, Latitude (°), Longitude (°),
'   ...). Extra descriptive columns such as Location are ALLOWED and are never
'   overwritten. Column order and total count are no longer enforced.
' - If a required column is missing, a visible error is written to _CLOUD_CFG
'   B10/B11 instead of failing silently.
'
' v1.8 CHANGE (K=6/K=12 ranges):
' - /v1/resource/plan now returns "ranges" (e.g. 202101-202106). VBA prefers
'   ranges and POSTs each range to /v1/resource/process-range. One range =
'   one NASA request covering up to RESOURCE_RANGE_MONTHS (Worker default 12)
'   months => 60 months = 5 ranges = only 1 wave at FAST_PARALLEL_REQUESTS=6
'   (was 10 waves). Mode + measured CPU ms are written to _CLOUD_CFG B12/B13.
' - If ANY range request fails (network/CPU/NASA 422/timeout) VBA switches to
'   the original month-by-month path (/v1/resource/process-month) for the
'   remaining months and never leaves them missing. Worst case = old speed,
'   never wrong data.
' - Old month behavior is retained byte-for-byte: same validation, same
'   monthly resource_chunks/summaries, same finalize + Excel import.
'==========================================================================

Private Const CONFIG_SHEET As String = "_CLOUD_CFG"
Private Const DRAWING_SHEET As String = "DRAWING_DATA"
Private Const SITE_TABLE As String = "autoLWHTbl"
Private Const RESOURCE_SHEET As String = "RESOURCE_DB"
Private Const RESOURCE_START_CONFIG_CELL As String = "B6"
Private Const RESOURCE_END_CONFIG_CELL As String = "B7"
Private Const RESOURCE_MODE_CONFIG_CELL As String = "B12"
Private Const RESOURCE_CPU_CONFIG_CELL As String = "B13"
Private Const PARAMETER_MAP_TABLE As String = "parameter_map"
Private Const RESOURCE_DB_TABLE As String = "resource_db"
Private Const MAX_RESOURCE_FAILURES As Long = 3
Private Const FAST_PARALLEL_REQUESTS As Long = 6
Private Const FAST_REQUEST_TIMEOUT_SECONDS As Long = 90

Private mRelayURL As String
Private mWorkbookID As String
Private mWorkbookKey As String
Private mQueue As Collection
Private mBusy As Boolean
Private mNextRun As Date
Private mFailureCount As Long
Private mStoppedAfterFailure As Boolean
Private mFastHTTP(1 To 6) As Object
Private mFastMonth(1 To 6) As String
Private mFastStarted(1 To 6) As Date
Private mFastActive As Boolean
Private mFastHadFailure As Boolean
Private mFastRequestID As String
Private mLastResourceError As String
' v1.8: range mode state. mFastRangeMode = currently launching ranges;
' mFastRangeFallback = a range wave failed, use month-by-month from now on.
Private mFastRangeMode As Boolean
Private mFastRangeFallback As Boolean
Private mFastResponse(1 To 6) As String
Private mFastRangeMonthsUsed As Long
' v1.9.2: last reverse-geocoded coordinate cache (same coords -> same label).
Private mLastLocKey As String
Private mLastLocLabel As String

'One-time SYSTEM configuration. Dates are not added to customer INPUT fields.
'Nothing is saved unless both dates are explicitly entered and validated.
Public Sub SolarEPC_ResourceConfigurePeriod()
    Dim ws As Worksheet
    Dim StartText As String
    Dim EndText As String
    Dim CurrentStart As String
    Dim CurrentEnd As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    If Not ws Is Nothing Then
        CurrentStart = ResourceConfiguredDateText(ws.Range(RESOURCE_START_CONFIG_CELL))
        CurrentEnd = ResourceConfiguredDateText(ws.Range(RESOURCE_END_CONFIG_CELL))
    End If
    On Error GoTo Failed
    If ws Is Nothing Then
        MsgBox "Cloud configuration is required before setting the resource period.", _
            vbExclamation, "Solar EPC Resource Configuration"
        Exit Sub
    End If

    StartText = Trim$(InputBox( _
        "Enter the NASA POWER resource start date in YYYY-MM-DD format." & vbCrLf & _
        "Example test configuration: 2021-01-01", _
        "Solar EPC Resource Period", CurrentStart))
    If Len(StartText) = 0 Then Exit Sub

    EndText = Trim$(InputBox( _
        "Enter the NASA POWER resource end date in YYYY-MM-DD format." & vbCrLf & _
        "Example test configuration: 2025-12-31", _
        "Solar EPC Resource Period", CurrentEnd))
    If Len(EndText) = 0 Then Exit Sub

    If Not ResourcePeriodValid(StartText, EndText) Then
        MsgBox "The resource period is invalid." & vbCrLf & vbCrLf & _
            "Use valid YYYY-MM-DD dates, start must not be after end, " & _
            "NASA POWER Hourly starts at 2001-01-01, and end cannot be in the future.", _
            vbExclamation, "Solar EPC Resource Configuration"
        Exit Sub
    End If

    'Store only system configuration values in the existing hidden cloud
    'configuration sheet. No customer INPUT field or Excel Table is changed.
    ws.Range("A6").Value2 = "NASA RESOURCE START DATE"
    With ws.Range(RESOURCE_START_CONFIG_CELL)
        .NumberFormat = "@"
        .Value2 = StartText
    End With
    ws.Range("A7").Value2 = "NASA RESOURCE END DATE"
    With ws.Range(RESOURCE_END_CONFIG_CELL)
        .NumberFormat = "@"
        .Value2 = EndText
    End With
    ws.Range("A8").Value2 = "NASA PERIOD CONFIGURED ON"
    ws.Range("B8").Value2 = Now
    ResourceSetLastError vbNullString

    MsgBox "NASA POWER resource period saved:" & vbCrLf & _
        StartText & " to " & EndText, vbInformation, _
        "Solar EPC Resource Configuration"
    Exit Sub
Failed:
    MsgBox "The NASA POWER resource period could not be saved." & vbCrLf & _
        Err.Description, vbExclamation, "Solar EPC Resource Configuration"
End Sub

Private Function ResourceLoadConfiguredPeriod(ByRef StartText As String, _
    ByRef EndText As String) As Boolean

    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    StartText = ResourceConfiguredDateText(ws.Range(RESOURCE_START_CONFIG_CELL))
    EndText = ResourceConfiguredDateText(ws.Range(RESOURCE_END_CONFIG_CELL))
    ResourceLoadConfiguredPeriod = ResourcePeriodValid(StartText, EndText)

    'Self-heal older configuration cells that Excel converted to date serials.
    If ResourceLoadConfiguredPeriod Then
        With ws.Range(RESOURCE_START_CONFIG_CELL)
            .NumberFormat = "@"
            .Value2 = StartText
        End With
        With ws.Range(RESOURCE_END_CONFIG_CELL)
            .NumberFormat = "@"
            .Value2 = EndText
        End With
    End If
    Exit Function
Failed:
    StartText = vbNullString
    EndText = vbNullString
    ResourceLoadConfiguredPeriod = False
End Function

Private Function ResourceConfiguredDateText(ByVal Target As Range) As String
    Dim RawValue As Variant
    Dim SerialValue As Double
    Dim Parsed As Date
    Dim TextValue As String

    On Error GoTo Failed
    RawValue = Target.Value2

    'New/correct format: exact ISO text.
    If VarType(RawValue) = vbString Then
        TextValue = Trim$(CStr(RawValue))
        If ResourceISODateValid(TextValue) Then
            ResourceConfiguredDateText = TextValue
        End If
        Exit Function
    End If

    'Backward-compatible recovery: Excel may already have converted the
    'previously validated ISO text into its numeric date serial.
    If IsNumeric(RawValue) Then
        SerialValue = CDbl(RawValue)
        If SerialValue > 0# And SerialValue < 2958466# Then
            Parsed = DateSerial(1899, 12, 30) + Fix(SerialValue)
            ResourceConfiguredDateText = Format$(Parsed, "yyyy-mm-dd")
        End If
    End If
    Exit Function
Failed:
    ResourceConfiguredDateText = vbNullString
End Function

Private Function ResourcePeriodValid(ByVal StartText As String, _
    ByVal EndText As String) As Boolean

    If Not ResourceISODateValid(StartText) Then Exit Function
    If Not ResourceISODateValid(EndText) Then Exit Function
    If StartText > EndText Then Exit Function
    If StartText < "2001-01-01" Then Exit Function
    If EndText > Format$(Date, "yyyy-mm-dd") Then Exit Function
    ResourcePeriodValid = True
End Function

Private Function ResourceISODateValid(ByVal TextValue As String) As Boolean
    Dim Re As Object
    Dim Y As Long, M As Long, D As Long
    Dim Parsed As Date

    Set Re = CreateObject("VBScript.RegExp")
    Re.Global = False
    Re.Pattern = "^[0-9]{4}-[0-9]{2}-[0-9]{2}$"
    If Not Re.Test(TextValue) Then Exit Function

    On Error GoTo InvalidDate
    Y = CLng(Left$(TextValue, 4))
    M = CLng(Mid$(TextValue, 6, 2))
    D = CLng(Right$(TextValue, 2))
    Parsed = DateSerial(Y, M, D)
    ResourceISODateValid = (Year(Parsed) = Y And Month(Parsed) = M And Day(Parsed) = D)
    Exit Function
InvalidDate:
    ResourceISODateValid = False
End Function

'Called additively by modSolarEPCCloudRelay only after the existing Excel
'import, message-ID log and verified cloud ACK have all succeeded.
Public Sub SolarEPC_ResourceQueueImportedSite(ByVal MessageID As String)
    Dim ws As Worksheet
    Dim Tbl As ListObject
    Dim ParameterMapTbl As ListObject
    Dim SiteRow As ListRow
    Dim ProjectID As String
    Dim CoordinateText As String
    Dim PolygonJSON As String
    Dim ResourceStartDate As String
    Dim ResourceEndDate As String
    Dim Body As String
    Dim ResponseText As String
    Dim RequestID As String
    Dim StatusCode As Long
    Dim Attempt As Long
    Dim cProject As Long
    Dim cCoordinates As Long

    On Error GoTo Failed
    ResourceSetLastError vbNullString
    MessageID = Trim$(MessageID)
    If Len(MessageID) = 0 Then
        ResourceSetLastError "RESOURCE START: message ID is blank."
        Exit Sub
    End If

    If Not ResourceLoadConfiguration() Then
        ResourceSetLastError "RESOURCE START: cloud configuration could not be loaded."
        Exit Sub
    End If

    Set ParameterMapTbl = ResourceFindWorkbookTable(PARAMETER_MAP_TABLE)
    If ParameterMapTbl Is Nothing Then
        ResourceSetLastError "RESOURCE START: existing Excel Table 'parameter_map' was not found."
        Exit Sub
    End If

    If Not ResourceLoadConfiguredPeriod(ResourceStartDate, ResourceEndDate) Then
        ResourceSetLastError "RESOURCE START: valid configured start/end dates were not found. Run SolarEPC_ResourceConfigurePeriod."
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets(DRAWING_SHEET)
    Set Tbl = ws.ListObjects(SITE_TABLE)
    Set SiteRow = ResourceFindImportedSiteRow(Tbl, "SITE-MAP-" & MessageID)
    If SiteRow Is Nothing Then
        ResourceSetLastError "RESOURCE START: imported SITE row was not found for message " & MessageID & "."
        Exit Sub
    End If

    cProject = ResourceColumn(Tbl, "Project ID")
    cCoordinates = ResourceColumn(Tbl, "Coordinates")
    If cProject = 0 Or cCoordinates = 0 Then
        ResourceSetLastError "RESOURCE START: required SITE table columns were not found."
        Exit Sub
    End If

    ProjectID = Trim$(CStr(SiteRow.Range.Cells(1, cProject).Value2))
    CoordinateText = Trim$(CStr(SiteRow.Range.Cells(1, cCoordinates).Value2))
    If Len(ProjectID) = 0 Then
        ResourceSetLastError "RESOURCE START: Project ID is blank in the imported SITE row."
        Exit Sub
    End If
    If Len(CoordinateText) = 0 Then
        ResourceSetLastError "RESOURCE START: SITE Coordinates are blank for project " & ProjectID & "."
        Exit Sub
    End If

    PolygonJSON = ResourcePolygonJSON(CoordinateText)
    If Len(PolygonJSON) = 0 Then
        ResourceSetLastError "RESOURCE START: SITE polygon coordinates are invalid for project " & ProjectID & "."
        Exit Sub
    End If

    Body = "{""projectId"":""" & ResourceJSONEscape(ProjectID) & _
           """,""messageId"":""" & ResourceJSONEscape(MessageID) & _
           """,""startDate"":""" & ResourceJSONEscape(ResourceStartDate) & _
           """,""endDate"":""" & ResourceJSONEscape(ResourceEndDate) & _
           """,""sitePolygon"":" & PolygonJSON & "}"

    For Attempt = 1 To 3
        ResponseText = ResourceRequest("POST", "/v1/resource/start", Body, StatusCode)
        If StatusCode = 200 Or StatusCode = 201 Then Exit For
        If StatusCode >= 400 And StatusCode < 500 And StatusCode <> 429 Then Exit For
        If Attempt < 3 Then
            DoEvents
            Application.Wait Now + TimeSerial(0, 0, 1)
        End If
    Next Attempt

    If StatusCode <> 200 And StatusCode <> 201 Then
        ResourceSetLastError "RESOURCE START failed for project " & ProjectID & _
            " (HTTP " & CStr(StatusCode) & "): " & Left$(ResponseText, 1000)
        Exit Sub
    End If

    RequestID = ResourceJSONValue(ResponseText, "requestId")
    If Len(RequestID) < 20 Then
        ResourceSetLastError "RESOURCE START: Worker returned no valid request ID for project " & ProjectID & "."
        Exit Sub
    End If

    'A completed cache hit already has a durable validated summary. Import it
    'immediately instead of waiting for an unnecessary scheduler round trip.
    If InStr(1, ResponseText, """complete"":true", vbTextCompare) > 0 Then
        ResourceImportSummary RequestID
        mStoppedAfterFailure = False
        ResourceSetLastError vbNullString
        Application.StatusBar = "Solar EPC: cached NASA resource imported for project " & ProjectID & "."
        Exit Sub
    End If

    ResourceEnqueue RequestID
    mStoppedAfterFailure = False
    ResourceSetLastError vbNullString
    Application.StatusBar = "Solar EPC: NASA resource queued for project " & ProjectID & "."
    If Not mBusy And mNextRun = 0 Then SolarEPC_ResourceProcessNext
    Exit Sub

Failed:
    ResourceSetLastError "RESOURCE START VBA error " & CStr(Err.Number) & ": " & Err.Description
End Sub

'Runs only on explicit user request and retries the most recently imported MAP
'SITE row. It reads autoLWHTbl only and never uses obstacle coordinates.
Public Sub SolarEPC_ResourceRetryLatestSite()
    Dim ws As Worksheet
    Dim Tbl As ListObject
    Dim R As Long
    Dim cReference As Long
    Dim ReferenceID As String
    Dim MessageID As String

    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(DRAWING_SHEET)
    Set Tbl = ws.ListObjects(SITE_TABLE)
    cReference = ResourceColumn(Tbl, "Reference id")
    If cReference = 0 Then Err.Raise vbObjectError + 8601, , "Reference id column was not found."

    For R = Tbl.ListRows.Count To 1 Step -1
        ReferenceID = Trim$(CStr(Tbl.ListRows(R).Range.Cells(1, cReference).Value2))
        If UCase$(Left$(ReferenceID, 9)) = "SITE-MAP-" Then
            MessageID = Mid$(ReferenceID, 10)
            Exit For
        End If
    Next R

    If Len(MessageID) = 0 Then
        MsgBox "No imported MAP SITE row was found in DRAWING_DATA.", vbExclamation, "Solar EPC Resource"
        Exit Sub
    End If

    SolarEPC_ResourceQueueImportedSite MessageID
    If Len(mLastResourceError) > 0 Then
        MsgBox mLastResourceError, vbExclamation, "Solar EPC Resource Start"
    Else
        MsgBox "The latest imported SITE resource request was queued successfully.", vbInformation, "Solar EPC Resource"
    End If
    Exit Sub
Failed:
    MsgBox "Could not retry the latest SITE resource request." & vbCrLf & _
        Err.Description, vbExclamation, "Solar EPC Resource"
End Sub

Public Sub SolarEPC_ResourceShowLastError()
    If Len(mLastResourceError) = 0 Then
        On Error Resume Next
        mLastResourceError = Trim$(CStr(ThisWorkbook.Worksheets(CONFIG_SHEET).Range("B11").Value2))
        On Error GoTo 0
    End If
    If Len(mLastResourceError) = 0 Then
        MsgBox "No resource-start error is recorded.", vbInformation, "Solar EPC Resource"
    Else
        MsgBox mLastResourceError, vbExclamation, "Solar EPC Resource Last Error"
    End If
End Sub

'v1.9.2: fills the descriptive Location column for existing RESOURCE_DB rows
'whose Location is still blank (nearest town/district via free reverse geocoding).
'No NASA/Worker call, no coordinates/values/status change, manual edits kept.
Public Sub SolarEPC_ResourceFillLocations()
    Dim ws As Worksheet
    Dim Tbl As ListObject
    Dim R As Long
    Dim cLoc As Long, cLat As Long, cLon As Long
    Dim RowRange As Range
    Dim Filled As Long
    Dim LatitudeText As String, LongitudeText As String

    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(RESOURCE_SHEET)
    Set Tbl = ResourceFindHeaderTable(ws)
    If Tbl Is Nothing Then
        MsgBox "RESOURCE_DB table nahi mila ya required column missing hai.", _
               vbExclamation, "Solar EPC Resource"
        Exit Sub
    End If
    cLoc = ResourceColumn(Tbl, "Location")
    cLat = ResourceColumn(Tbl, "Latitude (" & ChrW(176) & ")")
    cLon = ResourceColumn(Tbl, "Longitude (" & ChrW(176) & ")")
    If cLoc = 0 Or cLat = 0 Or cLon = 0 Then
        MsgBox "RESOURCE_DB me Location, Latitude aur Longitude columns required hain.", _
               vbExclamation, "Solar EPC Resource"
        Exit Sub
    End If

    For R = 1 To Tbl.ListRows.Count
        Set RowRange = Tbl.ListRows(R).Range
        If Len(Trim$(CStr(RowRange.Cells(1, cLoc).Value2))) = 0 Then
            LatitudeText = Trim$(CStr(RowRange.Cells(1, cLat).Value2))
            LongitudeText = Trim$(CStr(RowRange.Cells(1, cLon).Value2))
            If Len(LatitudeText) > 0 And Len(LongitudeText) > 0 Then
                ResourceFillLocationCell Tbl, RowRange, LatitudeText, LongitudeText
                If Len(Trim$(CStr(RowRange.Cells(1, cLoc).Value2))) > 0 Then
                    Filled = Filled + 1
                    'OSM policy: max ~1 request/second between lookups.
                    If R < Tbl.ListRows.Count Then Application.Wait Now + TimeSerial(0, 0, 1)
                End If
            End If
        End If
    Next R

    MsgBox "Location fill complete: " & CStr(Filled) & " row(s) filled. " & _
           "(Sirf blank cells bhare gaye; manually typed location change nahi hota.)", _
           vbInformation, "Solar EPC Resource"
    Exit Sub
Failed:
    MsgBox "Location fill failed: " & Err.Description, vbExclamation, "Solar EPC Resource"
End Sub

Private Sub ResourceSetLastError(ByVal ErrorText As String)
    Dim ws As Worksheet
    mLastResourceError = ErrorText
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    If Not ws Is Nothing Then
        ws.Range("A10").Value2 = "LAST NASA RESOURCE START ERROR"
        ws.Range("B10").Value2 = IIf(Len(ErrorText) > 0, Now, vbNullString)
        ws.Range("A11").Value2 = "ERROR DETAIL"
        ws.Range("B11").Value2 = ErrorText
    End If
    If Len(ErrorText) > 0 Then
        Application.StatusBar = ErrorText
        Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"), ErrorText
    End If
    On Error GoTo 0
End Sub

'Automatic finite resume: called at workbook open. It does nothing when no
'incomplete backend request exists and never creates permanent idle polling.
Public Sub SolarEPC_ResourceResumePending()
    Dim ResponseText As String
    Dim RequestID As String
    Dim StatusCode As Long

    On Error GoTo Failed
    If mBusy Or mNextRun > 0 Then Exit Sub
    If Not ResourceLoadConfiguration() Then Exit Sub

    ResponseText = ResourceRequest("GET", "/v1/resource/pending", vbNullString, StatusCode)
    If StatusCode = 204 Then Exit Sub
    If StatusCode <> 200 Then Exit Sub

    RequestID = ResourceJSONValue(ResponseText, "requestId")
    If Len(RequestID) < 20 Then Exit Sub
    ResourceEnqueue RequestID
    mStoppedAfterFailure = False
    SolarEPC_ResourceProcessNext
    Exit Sub
Failed:
End Sub

'Public OnTime entry point. Each invocation launches one parallel wave: up to
'six range requests (K=6) on the fast path, or six single-month requests on
'the proven fallback path. Excel stays responsive between waves.
Public Sub SolarEPC_ResourceProcessNext()
    Dim RequestID As String
    Dim ResponseText As String
    Dim StatusCode As Long
    Dim MonthsText As String
    Dim Months As Variant
    Dim RangesText As String
    Dim Ranges As Variant
    Dim UseRanges As Boolean
    Dim i As Long
    Dim Slot As Long
    Dim IsComplete As Boolean
    Dim ReadyToFinalize As Boolean

    On Error GoTo Failed
    mNextRun = 0
    If mQueue Is Nothing Then Exit Sub
    If mQueue.Count = 0 Then Exit Sub
    If Not ResourceLoadConfiguration() Then Exit Sub

    If mFastActive Then
        ResourceCheckFastRequests
        Exit Sub
    End If

    mBusy = True
    RequestID = CStr(mQueue(1))
    Application.StatusBar = "Solar EPC: checking fast NASA POWER resource progress..."

    ResponseText = ResourceRequest("GET", "/v1/resource/plan/" & RequestID, _
        vbNullString, StatusCode)
    If StatusCode <> 200 Then GoTo RetryFailed

    IsComplete = (InStr(1, ResponseText, """complete"":true", vbTextCompare) > 0)
    ReadyToFinalize = (InStr(1, ResponseText, """readyToFinalize"":true", vbTextCompare) > 0)

    If IsComplete Then
        ResourceFinishCurrent RequestID
        Exit Sub
    End If

    If ReadyToFinalize Then
        ResponseText = ResourceRequest("POST", "/v1/resource/finalize", _
            "{""requestId"":""" & ResourceJSONEscape(RequestID) & """}", StatusCode)
        If StatusCode = 200 And _
           InStr(1, ResponseText, """complete"":true", vbTextCompare) > 0 Then
            ResourceFinishCurrent RequestID
            Exit Sub
        End If
        GoTo RetryFailed
    End If

    MonthsText = ResourceJSONValue(ResponseText, "months")
    If Len(MonthsText) = 0 Then GoTo RetryFailed
    Months = Split(MonthsText, ",")

    'v1.8: prefer range mode (K=6) unless a previous range wave failed, in
    'which case the proven month-by-month path takes over for this request.
    UseRanges = False
    mFastRangeMonthsUsed = 0
    If Not mFastRangeFallback Then
        RangesText = ResourceJSONValue(ResponseText, "ranges")
        If Len(RangesText) > 0 Then
            Ranges = Split(RangesText, ",")
            If UBound(Ranges) >= LBound(Ranges) Then UseRanges = True
            mFastRangeMonthsUsed = Val(ResourceJSONValue(ResponseText, "rangeMonths"))
            If mFastRangeMonthsUsed <= 0 Then mFastRangeMonthsUsed = 6
        End If
    End If

    ResourceClearFastRequests False
    mFastRequestID = RequestID
    mFastHadFailure = False
    Slot = 0

    If UseRanges Then
        mFastRangeMode = True
        For i = LBound(Ranges) To UBound(Ranges)
            If Slot >= FAST_PARALLEL_REQUESTS Then Exit For
            Slot = Slot + 1
            If Not ResourceStartFastRange(Slot, RequestID, Trim$(CStr(Ranges(i)))) Then
                mFastHadFailure = True
            End If
        Next i
        Application.StatusBar = "Solar EPC: NASA " & CStr(mFastRangeMonthsUsed) & _
            "-month ranges processing in parallel; Excel remains available..."
    Else
        mFastRangeMode = False
        For i = LBound(Months) To UBound(Months)
            If Slot >= FAST_PARALLEL_REQUESTS Then Exit For
            Slot = Slot + 1
            If Not ResourceStartFastMonth(Slot, RequestID, Trim$(CStr(Months(i)))) Then
                mFastHadFailure = True
            End If
        Next i
        Application.StatusBar = "Solar EPC: NASA hourly months processing in parallel; Excel remains available..."
    End If

    If Slot = 0 Then GoTo RetryFailed
    mFastActive = True
    mBusy = False
    ResourceSchedule 1
    Exit Sub

RetryFailed:
    mBusy = False
    mFailureCount = mFailureCount + 1
    If mFailureCount < MAX_RESOURCE_FAILURES Then
        ResourceSchedule 5
    Else
        mStoppedAfterFailure = True
        Application.StatusBar = "Solar EPC: fast NASA processing stopped safely; run SolarEPC_ResourceResume to retry."
    End If
    Exit Sub

Failed:
    mBusy = False
    mFailureCount = mFailureCount + 1
    ResourceClearFastRequests True
    If mFailureCount < MAX_RESOURCE_FAILURES Then
        ResourceSchedule 5
    Else
        mStoppedAfterFailure = True
        Application.StatusBar = "Solar EPC: NASA resource processing stopped safely; map data remains safe."
    End If
End Sub

Private Function ResourceStartFastMonth(ByVal Slot As Long, _
    ByVal RequestID As String, ByVal MonthText As String) As Boolean

    Dim HTTP As Object
    Dim Body As String
    On Error GoTo Failed
    If Slot < 1 Or Slot > FAST_PARALLEL_REQUESTS Then Exit Function
    If Len(MonthText) <> 6 Or Not IsNumeric(MonthText) Then Exit Function

    Set HTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    HTTP.Open "POST", mRelayURL & "/v1/resource/process-month", True
    HTTP.SetTimeouts 5000, 5000, 5000, FAST_REQUEST_TIMEOUT_SECONDS * 1000
    HTTP.SetRequestHeader "Content-Type", "application/json"
    HTTP.SetRequestHeader "Cache-Control", "no-store"
    HTTP.SetRequestHeader "X-Solar-EPC-Workbook", mWorkbookID
    HTTP.SetRequestHeader "X-Solar-EPC-Key", mWorkbookKey
    Body = "{""requestId"":""" & ResourceJSONEscape(RequestID) & _
           """,""month"":""" & ResourceJSONEscape(MonthText) & """}"
    HTTP.Send Body

    Set mFastHTTP(Slot) = HTTP
    mFastMonth(Slot) = MonthText
    mFastStarted(Slot) = Now
    ResourceStartFastMonth = True
    Exit Function
Failed:
    Set mFastHTTP(Slot) = Nothing
    mFastMonth(Slot) = vbNullString
    mFastStarted(Slot) = 0
End Function

'v1.8: launches one 6-month range request to /v1/resource/process-range.
'RangeText is expected in Worker format "YYYYMM-YYYYMM" (e.g. 202101-202106).
Private Function ResourceStartFastRange(ByVal Slot As Long, _
    ByVal RequestID As String, ByVal RangeText As String) As Boolean

    Dim HTTP As Object
    Dim Body As String
    Dim StartMonth As String
    Dim EndMonth As String
    On Error GoTo Failed
    If Slot < 1 Or Slot > FAST_PARALLEL_REQUESTS Then Exit Function
    If Len(RangeText) <> 13 Then Exit Function
    If Mid$(RangeText, 7, 1) <> "-" Then Exit Function
    StartMonth = Left$(RangeText, 6)
    EndMonth = Right$(RangeText, 6)
    If Not IsNumeric(StartMonth) Or Not IsNumeric(EndMonth) Then Exit Function
    If StartMonth > EndMonth Then Exit Function

    Set HTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    HTTP.Open "POST", mRelayURL & "/v1/resource/process-range", True
    HTTP.SetTimeouts 5000, 5000, 5000, FAST_REQUEST_TIMEOUT_SECONDS * 1000
    HTTP.SetRequestHeader "Content-Type", "application/json"
    HTTP.SetRequestHeader "Cache-Control", "no-store"
    HTTP.SetRequestHeader "X-Solar-EPC-Workbook", mWorkbookID
    HTTP.SetRequestHeader "X-Solar-EPC-Key", mWorkbookKey
    Body = "{""requestId"":""" & ResourceJSONEscape(RequestID) & _
           """,""rangeStart"":""" & ResourceJSONEscape(StartMonth) & _
           """,""rangeEnd"":""" & ResourceJSONEscape(EndMonth) & """}"
    HTTP.Send Body

    Set mFastHTTP(Slot) = HTTP
    mFastMonth(Slot) = RangeText
    mFastStarted(Slot) = Now
    ResourceStartFastRange = True
    Exit Function
Failed:
    Set mFastHTTP(Slot) = Nothing
    mFastMonth(Slot) = vbNullString
    mFastStarted(Slot) = 0
End Function

Private Sub ResourceCheckFastRequests()
    Dim i As Long
    Dim StillRunning As Boolean
    Dim FinishedCount As Long
    Dim FailedCount As Long
    Dim StatusCode As Long

    On Error GoTo Failed
    mNextRun = 0

    For i = 1 To FAST_PARALLEL_REQUESTS
        If Not mFastHTTP(i) Is Nothing Then
            If mFastHTTP(i).readyState = 4 Then
                StatusCode = 0
                On Error Resume Next
                StatusCode = CLng(mFastHTTP(i).Status)
                On Error GoTo Failed
                If StatusCode <> 200 Then FailedCount = FailedCount + 1
                mFastResponse(i) = vbNullString
                If StatusCode = 200 Then
                    On Error Resume Next
                    mFastResponse(i) = CStr(mFastHTTP(i).responseText)
                    On Error GoTo Failed
                End If
                Set mFastHTTP(i) = Nothing
                mFastMonth(i) = vbNullString
                mFastStarted(i) = 0
                FinishedCount = FinishedCount + 1
            ElseIf DateDiff("s", mFastStarted(i), Now) >= FAST_REQUEST_TIMEOUT_SECONDS Then
                On Error Resume Next
                mFastHTTP(i).abort
                On Error GoTo Failed
                Set mFastHTTP(i) = Nothing
                mFastMonth(i) = vbNullString
                mFastStarted(i) = 0
                mFastResponse(i) = vbNullString
                FailedCount = FailedCount + 1
                FinishedCount = FinishedCount + 1
            Else
                StillRunning = True
            End If
        End If
    Next i

    If StillRunning Then
        ResourceSchedule 1
        Exit Sub
    End If

    mFastActive = False

    'v1.8: record which path ran (RANGE12 / RANGE6 / MONTH) and the fastest
    'range CPU ms seen, so the Free 10ms CPU budget can be verified.
    If mFastRangeMode And FailedCount = 0 Then
        Dim Wire As Long
        Dim CpuMs As Long
        For Wire = 1 To FAST_PARALLEL_REQUESTS
            If Len(mFastResponse(Wire)) > 0 Then
                CpuMs = Val(ResourceJSONValue(mFastResponse(Wire), "processMs"))
                If CpuMs > 0 Then
                    ResourceWriteDiagnostic "RANGE" & CStr(mFastRangeMonthsUsed), CpuMs
                    Exit For
                End If
            End If
        Next Wire
        If CpuMs = 0 Then ResourceWriteDiagnostic "RANGE" & CStr(mFastRangeMonthsUsed), 0
    ElseIf mFastRangeMode And FailedCount > 0 Then
        ResourceWriteDiagnostic "RANGE" & CStr(mFastRangeMonthsUsed) & "_FAIL", 0
    ElseIf Not mFastRangeMode Then
        ResourceWriteDiagnostic "MONTH", 0
    End If
    If FailedCount > 0 Or mFastHadFailure Then
        'v1.8: any failure in range mode permanently switches this request to
        'the proven month-by-month path (old behavior), so nothing stays missing.
        If mFastRangeMode Then mFastRangeFallback = True
        mFailureCount = mFailureCount + 1
        If mFailureCount >= MAX_RESOURCE_FAILURES Then
            mStoppedAfterFailure = True
            Application.StatusBar = "Solar EPC: some NASA months failed safely; run SolarEPC_ResourceResume to retry."
            Exit Sub
        End If
        Application.StatusBar = "Solar EPC: retrying only incomplete NASA months..."
        ResourceSchedule 5
    Else
        mFailureCount = 0
        Application.StatusBar = "Solar EPC: validated NASA data stored; continuing fast parallel acquisition..."
        SolarEPC_ResourceProcessNext
    End If
    Exit Sub
Failed:
    mFastActive = False
    ResourceClearFastRequests True
    mFailureCount = mFailureCount + 1
    If mFailureCount < MAX_RESOURCE_FAILURES Then ResourceSchedule 5
End Sub

Private Sub ResourceFinishCurrent(ByVal RequestID As String)
    ResourceImportSummary RequestID
    If Not mQueue Is Nothing Then
        If mQueue.Count > 0 Then mQueue.Remove 1
    End If
    mFailureCount = 0
    mStoppedAfterFailure = False
    mFastActive = False
    mBusy = False
    mFastRangeMode = False
    mFastRangeFallback = False
    ResourceClearFastRequests False
    Application.StatusBar = "Solar EPC: NASA POWER resource validated and RESOURCE_DB updated."
    If Not mQueue Is Nothing Then
        If mQueue.Count > 0 Then
            ResourceSchedule 1
        Else
            SolarEPC_ResourceResumePending
        End If
    End If
End Sub

Private Sub ResourceClearFastRequests(ByVal AbortRequests As Boolean)
    Dim i As Long
    For i = 1 To FAST_PARALLEL_REQUESTS
        If Not mFastHTTP(i) Is Nothing Then
            If AbortRequests Then
                On Error Resume Next
                mFastHTTP(i).abort
                On Error GoTo 0
            End If
        End If
        Set mFastHTTP(i) = Nothing
        mFastMonth(i) = vbNullString
        mFastStarted(i) = 0
        mFastResponse(i) = vbNullString
    Next i
    mFastActive = False
    mFastRequestID = vbNullString
    mFastRangeMode = False
End Sub

'v1.8: writes the verified processing mode and measured range CPU ms into the
'existing hidden _CLOUD_CFG sheet (B12/B13). Diagnostic only; no input field
'and no Excel Table is touched.
Private Sub ResourceWriteDiagnostic(ByVal ModeText As String, ByVal CpuMs As Long)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    If Not ws Is Nothing Then
        ws.Range("A12").Value2 = "RESOURCE PROCESSING MODE"
        ws.Range(RESOURCE_MODE_CONFIG_CELL).Value2 = ModeText
        ws.Range("A13").Value2 = "RANGE PROCESS MS (LAST)"
        ws.Range(RESOURCE_CPU_CONFIG_CELL).Value2 = CpuMs
    End If
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"), "RESOURCE MODE:", ModeText, "CPU(ms):", CpuMs
    On Error GoTo 0
End Sub

Public Sub SolarEPC_ResourceStop()
    On Error Resume Next
    If mNextRun > 0 Then
        Application.OnTime EarliestTime:=mNextRun, _
            Procedure:="SolarEPC_ResourceProcessNext", Schedule:=False
    End If
    mNextRun = 0
    ResourceClearFastRequests True
    mFastRangeFallback = False
    mBusy = False
    On Error GoTo 0
End Sub

'Optional button/macro for a user-initiated retry after an outage.
Public Sub SolarEPC_ResourceResume()
    mFailureCount = 0
    mStoppedAfterFailure = False
    mFastRangeFallback = False
    If Not mQueue Is Nothing Then
        If mQueue.Count > 0 Then
            If mNextRun = 0 Then ResourceSchedule 1
            Exit Sub
        End If
    End If
    SolarEPC_ResourceResumePending
End Sub

Private Sub ResourceEnqueue(ByVal RequestID As String)
    Dim Item As Variant
    If mQueue Is Nothing Then Set mQueue = New Collection
    For Each Item In mQueue
        If StrComp(CStr(Item), RequestID, vbBinaryCompare) = 0 Then Exit Sub
    Next Item
    mQueue.Add RequestID
End Sub

Private Sub ResourceSchedule(ByVal SecondsFromNow As Long)
    If mBusy Then Exit Sub
    If mQueue Is Nothing Then Exit Sub
    If mQueue.Count = 0 Then Exit Sub
    On Error Resume Next
    If mNextRun > 0 Then
        Application.OnTime EarliestTime:=mNextRun, _
            Procedure:="SolarEPC_ResourceProcessNext", Schedule:=False
    End If
    On Error GoTo 0
    mNextRun = Now + TimeSerial(0, 0, SecondsFromNow)
    Application.OnTime EarliestTime:=mNextRun, _
        Procedure:="SolarEPC_ResourceProcessNext", Schedule:=True
End Sub

Private Function ResourceFindImportedSiteRow(ByVal Tbl As ListObject, _
    ByVal ReferenceID As String) As ListRow

    Dim R As Long
    Dim cRef As Long
    cRef = ResourceColumn(Tbl, "Reference id")
    If cRef = 0 Then Exit Function
    For R = Tbl.ListRows.Count To 1 Step -1
        If StrComp(Trim$(CStr(Tbl.ListRows(R).Range.Cells(1, cRef).Value2)), _
                   ReferenceID, vbBinaryCompare) = 0 Then
            Set ResourceFindImportedSiteRow = Tbl.ListRows(R)
            Exit Function
        End If
    Next R
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

Private Function ResourcePolygonJSON(ByVal CoordinateText As String) As String
    Dim Points As Variant
    Dim Pair As Variant
    Dim i As Long
    Dim LatText As String
    Dim LonText As String
    Dim LatValue As Double
    Dim LonValue As Double
    Dim Result As String
    Dim Count As Long

    Points = Split(CoordinateText, ";")
    Result = "["
    For i = LBound(Points) To UBound(Points)
        Pair = Split(Trim$(CStr(Points(i))), ",")
        If UBound(Pair) <> 1 Then Exit Function
        LatText = Trim$(CStr(Pair(0)))
        LonText = Trim$(CStr(Pair(1)))
        If Not ResourceStrictNumber(LatText) Or Not ResourceStrictNumber(LonText) Then Exit Function
        LatValue = Val(LatText)
        LonValue = Val(LonText)
        If LatValue < -90# Or LatValue > 90# Or LonValue < -180# Or LonValue > 180# Then Exit Function
        If Count > 0 Then Result = Result & ","
        Result = Result & "[" & ResourceDecimal(LatValue, 8) & "," & _
                 ResourceDecimal(LonValue, 8) & "]"
        Count = Count + 1
    Next i
    If Count < 3 Then Exit Function
    ResourcePolygonJSON = Result & "]"
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

Private Function ResourceRequest(ByVal MethodName As String, ByVal Path As String, _
    ByVal Body As String, ByRef StatusCode As Long) As String

    ResourceRequest = ResourceRequestClient("WinHttp.WinHttpRequest.5.1", _
        MethodName, Path, Body, StatusCode)
    If StatusCode = 0 Then
        ResourceRequest = ResourceRequestClient("MSXML2.ServerXMLHTTP.6.0", _
            MethodName, Path, Body, StatusCode)
    End If
End Function

Private Function ResourceRequestClient(ByVal ProgID As String, _
    ByVal MethodName As String, ByVal Path As String, ByVal Body As String, _
    ByRef StatusCode As Long) As String

    Dim HTTP As Object
    On Error GoTo Failed
    Set HTTP = CreateObject(ProgID)
    HTTP.Open MethodName, mRelayURL & Path, False
    HTTP.SetTimeouts 5000, 5000, 5000, 60000
    HTTP.SetRequestHeader "Content-Type", "application/json"
    HTTP.SetRequestHeader "Cache-Control", "no-store"
    HTTP.SetRequestHeader "X-Solar-EPC-Workbook", mWorkbookID
    HTTP.SetRequestHeader "X-Solar-EPC-Key", mWorkbookKey
    If UCase$(MethodName) = "GET" Then HTTP.Send Else HTTP.Send Body
    StatusCode = HTTP.Status
    If StatusCode <> 204 Then ResourceRequestClient = HTTP.ResponseText
    Exit Function
Failed:
    StatusCode = 0
    ResourceRequestClient = vbNullString
End Function

'v1.9.2: generic read-only JSON GET for LOCATION LABEL ONLY (reverse geocoding).
'Never touches resource data. Tries WinHttp, then MSXML2.
Private Function ResourceHttpGetJson(ByVal URLText As String) As String
    Dim HTTP As Object
    On Error GoTo TryFallback
    Set HTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    HTTP.Open "GET", URLText, False
    HTTP.SetTimeouts 6000, 6000, 15000, 15000
    HTTP.SetRequestHeader "User-Agent", "Solar-EPC-Resource/1.0 (Excel workbook location label)"
    HTTP.SetRequestHeader "Accept", "application/json"
    HTTP.Send
    If HTTP.Status = 200 Then
        ResourceHttpGetJson = HTTP.ResponseText
        Exit Function
    End If
TryFallback:
    On Error GoTo Failed
    Set HTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    HTTP.Open "GET", URLText, False
    HTTP.setTimeouts 6000, 6000, 15000, 15000
    HTTP.setRequestHeader "User-Agent", "Solar-EPC-Resource/1.0 (Excel workbook location label)"
    HTTP.setRequestHeader "Accept", "application/json"
    HTTP.send
    If HTTP.Status = 200 Then ResourceHttpGetJson = HTTP.ResponseText
    Exit Function
Failed:
    ResourceHttpGetJson = vbNullString
End Function

'v1.9.2: compose "Town/Village, District" from reverse-geocoded components.
Private Function ResourceLocationCompose(ByVal Primary As String, ByVal Secondary As String) As String
    Dim P As String, S2 As String
    P = Trim$(Primary)
    S2 = Trim$(Secondary)
    If Len(P) = 0 Then Exit Function
    If UCase$(Right$(P, 7)) = ", INDIA" Then P = Left$(P, Len(P) - 7)
    If UCase$(Right$(S2, 7)) = ", INDIA" Then S2 = Left$(S2, Len(S2) - 7)
    If Len(S2) > 0 Then
        ResourceLocationCompose = P & ", " & S2
    Else
        ResourceLocationCompose = P
    End If
End Function

'v1.9.2: nearest readable location from exact centroid (e.g. "Phaltan, Satara").
'Primary: OSM Nominatim (free, no key); fallback: BigDataCloud (free, no key).
'Result is descriptive ONLY - identity stays centroid_lat/centroid_lon.
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

    '1) OpenStreetMap Nominatim (free, no API key)
    UrlText = "https://nominatim.openstreetmap.org/reverse?format=jsonv2&lat=" & _
              ResourceDecimal(Lat, 6) & "&lon=" & ResourceDecimal(Lon, 6) & _
              "&addressdetails=1&zoom=16&accept-language=en"
    JSONText = ResourceHttpGetJson(UrlText)
    If Len(JSONText) > 0 Then
        Primary = ResourceJSONValue(JSONText, "town")
        If Len(Primary) = 0 Then Primary = ResourceJSONValue(JSONText, "city")
        If Len(Primary) = 0 Then Primary = ResourceJSONValue(JSONText, "village")
        If Len(Primary) = 0 Then Primary = ResourceJSONValue(JSONText, "hamlet")
        Secondary = ResourceJSONValue(JSONText, "state_district")
        If Len(Secondary) = 0 Then Secondary = ResourceJSONValue(JSONText, "county")
        If UCase$(Right$(Secondary, 9)) = " DISTRICT" Then Secondary = Left$(Secondary, Len(Secondary) - 9)
        If Len(Primary) = 0 Then
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

    '2) BigDataCloud free client reverse-geocode (no key) fallback
    UrlText = "https://api.bigdatacloud.net/data/reverse-geocode-client?latitude=" & _
              ResourceDecimal(Lat, 6) & "&longitude=" & ResourceDecimal(Lon, 6) & _
              "&localityLanguage=en"
    JSONText = ResourceHttpGetJson(UrlText)
    If Len(JSONText) > 0 Then
        Primary = ResourceJSONValue(JSONText, "locality")
        If Len(Primary) = 0 Then Primary = ResourceJSONValue(JSONText, "city")
        Secondary = ResourceJSONValue(JSONText, "principalSubdivision")
        If Len(Primary) > 0 Then
            ResourceLocationLabel = ResourceLocationCompose(Primary, Secondary)
        End If
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

'v1.9.2: fill blank Location cell for one row (never overwrites manual entries).
Private Sub ResourceFillLocationCell(ByVal Tbl As ListObject, ByVal Target As Range, _
    ByVal LatitudeText As String, ByVal LongitudeText As String)
    Dim cLoc As Long
    Dim Cell As Range
    Dim LabelText As String
    On Error GoTo Failed
    cLoc = ResourceColumn(Tbl, "Location")
    If cLoc = 0 Then Exit Sub
    Set Cell = Target.Cells(1, cLoc)
    If Len(Trim$(CStr(Cell.Value2))) > 0 Then Exit Sub
    If Len(Trim$(LatitudeText)) = 0 Or Len(Trim$(LongitudeText)) = 0 Then Exit Sub
    LabelText = ResourceLocationLabel(LatitudeText, LongitudeText)
    If Len(LabelText) > 0 Then Cell.Value2 = LabelText
    Exit Sub
Failed:
End Sub

Private Function ResourceLoadConfiguration() As Boolean
    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    mRelayURL = Trim$(CStr(ws.Range("B2").Value2))
    mWorkbookID = Trim$(CStr(ws.Range("B3").Value2))
    mWorkbookKey = Trim$(CStr(ws.Range("B4").Value2))
    Do While Right$(mRelayURL, 1) = "/"
        mRelayURL = Left$(mRelayURL, Len(mRelayURL) - 1)
    Loop
    ResourceLoadConfiguration = (LCase$(Left$(mRelayURL, 8)) = "https://" And _
        Len(mWorkbookID) >= 8 And Len(mWorkbookKey) >= 24)
    Exit Function
Failed:
    ResourceLoadConfiguration = False
End Function

Private Function ResourceJSONEscape(ByVal TextValue As String) As String
    Dim S As String
    S = TextValue
    S = Replace$(S, "\", "\\")
    S = Replace$(S, Chr$(34), "\" & Chr$(34))
    S = Replace$(S, vbCr, " ")
    S = Replace$(S, vbLf, " ")
    S = Replace$(S, vbTab, " ")
    ResourceJSONEscape = S
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

Private Sub ResourceImportSummary(ByVal RequestID As String)
    Dim StatusCode As Long
    Dim Text As String
    Dim Values As Object

    On Error GoTo Failed
    Text = ResourceRequest("GET", "/v1/resource/summary/" & RequestID, _
                           vbNullString, StatusCode)
    If StatusCode <> 200 Then Exit Sub
    Set Values = ResourceParseSummary(Text)
    If Values Is Nothing Then Exit Sub
    ResourceWriteToDatabase Values
    Exit Sub
Failed:
End Sub

Private Function ResourceParseSummary(ByVal Text As String) As Object
    Dim D As Object
    Dim Lines As Variant
    Dim P As Long
    Dim i As Long
    Dim K As String
    Dim V As String

    Set D = CreateObject("Scripting.Dictionary")
    D.CompareMode = vbTextCompare
    Lines = Split(Replace$(Replace$(Text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For i = LBound(Lines) To UBound(Lines)
        P = InStr(1, CStr(Lines(i)), "|", vbBinaryCompare)
        If P > 1 Then
            K = Left$(CStr(Lines(i)), P - 1)
            V = Mid$(CStr(Lines(i)), P + 1)
            D(K) = V
        End If
    Next i
    If Not D.Exists("schema") Then Exit Function
    If D("schema") <> "SOLAR_EPC_RESOURCE_SUMMARY_V1" Then Exit Function
    Set ResourceParseSummary = D
End Function

Private Sub ResourceWriteToDatabase(ByVal D As Object)
    Dim ws As Worksheet
    Dim Tbl As ListObject
    Dim Target As Range
    Dim ProjectID As String

    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(RESOURCE_SHEET)
    ProjectID = ResourceValue(D, "project_id")
    If Len(ProjectID) = 0 Then Exit Sub

    Set Tbl = ResourceFindHeaderTable(ws)
    If Tbl Is Nothing Then Exit Sub
    Set Target = ResourceProjectTableRow(Tbl, ProjectID, _
        ResourceValue(D, "latitude"), ResourceValue(D, "longitude"))
    If Target Is Nothing Then Exit Sub
    ResourceWriteTableField Tbl, Target, "Sr.No", CStr(ResourceSerialForTable(Tbl, Target))
    ResourceWriteTableField Tbl, Target, "Project ID", ProjectID
    'v1.9.2: descriptive Location auto-fill (blank cells only, never overwrite).
    ResourceFillLocationCell Tbl, Target, ResourceValue(D, "latitude"), ResourceValue(D, "longitude")
    ResourceWriteTableField Tbl, Target, "Latitude (" & ChrW(176) & ")", ResourceValue(D, "latitude"), True
    ResourceWriteTableField Tbl, Target, "Longitude (" & ChrW(176) & ")", ResourceValue(D, "longitude"), True
    ResourceWriteTableField Tbl, Target, "Source", ResourceValue(D, "source")
    ResourceWriteTableField Tbl, Target, "Dataset / Product", ResourceValue(D, "dataset_product")
    ResourceWriteTableField Tbl, Target, "Retrieval Date", ResourceValue(D, "retrieval_date")
    ResourceWriteTableField Tbl, Target, "Data Period Start", ResourceValue(D, "data_period_start")
    ResourceWriteTableField Tbl, Target, "Data Period End", ResourceValue(D, "data_period_end")
    ResourceWriteTableField Tbl, Target, "Resolution", ResourceValue(D, "resolution")
    ResourceWriteTableField Tbl, Target, "GHI", ResourceValue(D, "ghi"), True
    ResourceWriteTableField Tbl, Target, "DNI", ResourceValue(D, "dni"), True
    ResourceWriteTableField Tbl, Target, "DHI", ResourceValue(D, "dhi"), True
    ResourceWriteTableField Tbl, Target, "GTI / POA Irradiance", "DERIVED"
    ResourceWriteTableField Tbl, Target, "Ambient Temperature", ResourceValue(D, "ambient_temperature"), True
    ResourceWriteTableField Tbl, Target, "Min Temperature", ResourceValue(D, "min_temperature"), True
    ResourceWriteTableField Tbl, Target, "Max Temperature", ResourceValue(D, "max_temperature"), True
    ResourceWriteTableField Tbl, Target, "Wind Speed", ResourceValue(D, "wind_speed"), True
    ResourceWriteTableField Tbl, Target, "Wind Direction", ResourceValue(D, "wind_direction"), True
    ResourceWriteTableField Tbl, Target, "Relative Humidity", ResourceValue(D, "relative_humidity"), True
    ResourceWriteTableField Tbl, Target, "Surface Pressure", ResourceValue(D, "surface_pressure"), True
    ResourceWriteTableField Tbl, Target, "Precipitation", ResourceValue(D, "precipitation"), True
    ResourceWriteTableField Tbl, Target, "Albedo", ResourceValue(D, "albedo"), True
    ResourceWriteTableField Tbl, Target, "Data Status", ResourceValue(D, "data_status")
    ResourceWriteTableField Tbl, Target, "Source Reference", ResourceValue(D, "source_reference")
    ResourceWriteTableField Tbl, Target, "Remarks", ResourceValue(D, "remarks")
    Exit Sub
Failed:
End Sub

Private Function ResourceHeaders() As Variant
    ResourceHeaders = Array("Sr.No", "Project ID", "Latitude (" & ChrW(176) & ")", "Longitude (" & ChrW(176) & ")", _
        "Source", "Dataset / Product", "Retrieval Date", "Data Period Start", _
        "Data Period End", "Resolution", "GHI", "DNI", "DHI", _
        "GTI / POA Irradiance", "Ambient Temperature", "Min Temperature", _
        "Max Temperature", "Wind Speed", "Wind Direction", "Relative Humidity", _
        "Surface Pressure", "Precipitation", "Albedo", "Data Status", _
        "Source Reference", "Remarks")
End Function

Private Function ResourceFindWorkbookTable(ByVal TableName As String) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        Set ResourceFindWorkbookTable = Nothing
        Set ResourceFindWorkbookTable = ws.ListObjects(TableName)
        If Not ResourceFindWorkbookTable Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Private Function ResourceFindHeaderTable(ByVal ws As Worksheet) As ListObject
    Dim Tbl As ListObject
    Dim Required As Variant
    Dim i As Long
    Dim Missing As String

    'Use only the existing RESOURCE_DB Excel Table by its confirmed name.
    'Never create, convert, rename or substitute another table/range.
    Set Tbl = ResourceFindWorkbookTable(RESOURCE_DB_TABLE)
    If Tbl Is Nothing Then
        ResourceSetLastError "RESOURCE IMPORT: Excel Table '" & RESOURCE_DB_TABLE & "' was not found in this workbook."
        Exit Function
    End If
    If Not Tbl.Parent Is ws Then Exit Function

    'MATCH BY NAME ONLY. RESOURCE_DB may include extra descriptive columns
    '(e.g. Location). Column count/order are NOT enforced; required columns are
    'located by header name, written by name, and Location is never overwritten.
    Required = ResourceHeaders()
    For i = LBound(Required) To UBound(Required)
        If ResourceColumn(Tbl, CStr(Required(i))) = 0 Then
            If Len(Missing) = 0 Then
                Missing = CStr(Required(i))
            Else
                Missing = Missing & ", " & CStr(Required(i))
            End If
        End If
    Next i
    If Len(Missing) > 0 Then
        ResourceSetLastError "RESOURCE IMPORT: RESOURCE_DB is missing required column(s): " & _
            Missing & ". Required columns are matched by name; extra descriptive columns (e.g. Location) are allowed."
        Exit Function
    End If

    Set ResourceFindHeaderTable = Tbl
End Function

Private Function ResourceProjectTableRow(ByVal Tbl As ListObject, _
    ByVal ProjectID As String, ByVal LatitudeText As String, _
    ByVal LongitudeText As String) As Range

    Dim R As Long
    Dim cProject As Long
    Dim cLatitude As Long
    Dim cLongitude As Long
    Dim RowProject As String
    Dim RowLatitude As String
    Dim RowLongitude As String
    Dim TargetLatitude As Double
    Dim TargetLongitude As Double
    Dim EmptyProjectMatch As Range

    cProject = Tbl.ListColumns("Project ID").Index
    cLatitude = Tbl.ListColumns("Latitude (" & ChrW(176) & ")").Index
    cLongitude = Tbl.ListColumns("Longitude (" & ChrW(176) & ")").Index
    TargetLatitude = Val(LatitudeText)
    TargetLongitude = Val(LongitudeText)

    'A project name is not a unique SITE identity. Reuse a row only when the
    'same project and centroid already exist (idempotent retry), or when that
    'project has a prepared row whose resource coordinates are still blank.
    For R = 1 To Tbl.ListRows.Count
        RowProject = Trim$(CStr(Tbl.ListRows(R).Range.Cells(1, cProject).Value2))
        If StrComp(RowProject, ProjectID, vbTextCompare) = 0 Then
            RowLatitude = Trim$(CStr(Tbl.ListRows(R).Range.Cells(1, cLatitude).Value2))
            RowLongitude = Trim$(CStr(Tbl.ListRows(R).Range.Cells(1, cLongitude).Value2))
            If Len(RowLatitude) = 0 And Len(RowLongitude) = 0 Then
                If EmptyProjectMatch Is Nothing Then _
                    Set EmptyProjectMatch = Tbl.ListRows(R).Range
            ElseIf IsNumeric(RowLatitude) And IsNumeric(RowLongitude) Then
                If Abs(CDbl(RowLatitude) - TargetLatitude) < 0.0000005 And _
                   Abs(CDbl(RowLongitude) - TargetLongitude) < 0.0000005 Then
                    Set ResourceProjectTableRow = Tbl.ListRows(R).Range
                    Exit Function
                End If
            End If
        End If
    Next R

    If Not EmptyProjectMatch Is Nothing Then
        Set ResourceProjectTableRow = EmptyProjectMatch
        Exit Function
    End If

    'v3.22.11: never steal the first blank template row at the top of the
    'table. Insert directly below the last row that already carries content
    'so fills stack one-below-the-other (ek ke niche ek), same as auto.
    Set ResourceProjectTableRow = ResourceInsertAfterLastUsed(Tbl, cProject, cLatitude, cLongitude)
End Function

Private Function ResourceInsertAfterLastUsed(ByVal Tbl As ListObject, _
    ByVal cProject As Long, ByVal cLat As Long, ByVal cLon As Long) As Range

    Dim R As Long
    Dim LastUsed As Long
    Dim RowRange As Range
    Dim NewRow As ListRow

    LastUsed = 0
    For R = 1 To Tbl.ListRows.Count
        Set RowRange = Tbl.ListRows(R).Range
        If Len(Trim$(CStr(RowRange.Cells(1, cProject).Value2 & ""))) > 0 Then LastUsed = R
        If Len(Trim$(CStr(RowRange.Cells(1, cLat).Value2 & ""))) > 0 Then LastUsed = R
        If Len(Trim$(CStr(RowRange.Cells(1, cLon).Value2 & ""))) > 0 Then LastUsed = R
    Next R
    If LastUsed = 0 Or LastUsed >= Tbl.ListRows.Count Then
        Set NewRow = Tbl.ListRows.Add
    Else
        Set NewRow = Tbl.ListRows.Add(LastUsed + 1)
    End If
    Set ResourceInsertAfterLastUsed = NewRow.Range
End Function

Private Function ResourceSerialForTable(ByVal Tbl As ListObject, ByVal Target As Range) As Long
    ResourceSerialForTable = Target.Row - Tbl.HeaderRowRange.Row
End Function

Private Sub ResourceWriteTableField(ByVal Tbl As ListObject, ByVal Target As Range, _
    ByVal HeaderText As String, ByVal ValueText As String, Optional ByVal Numeric As Boolean = False)

    Dim Cell As Range
    Set Cell = Target.Cells(1, Tbl.ListColumns(HeaderText).Index)
    If Cell.HasFormula Then Exit Sub
    If Numeric Then
        If Len(Trim$(ValueText)) = 0 Then
            Cell.ClearContents
        Else
            Cell.Value2 = Val(ValueText)
        End If
    Else
        Cell.Value2 = ValueText
    End If
End Sub

Private Function ResourceValue(ByVal D As Object, ByVal KeyName As String) As String
    If D.Exists(KeyName) Then ResourceValue = CStr(D(KeyName))
End Function
