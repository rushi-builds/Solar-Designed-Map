Attribute VB_Name = "modSolarEPCResource"
Option Explicit

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

'==========================================================================
' SOLAR EPC - NASA POWER HOURLY RESOURCE MODULE + AUTOMATIC LOCATION FILL
' Version 3.11 (FINAL)
'
' THIS FILE IS A COMPLETE REPLACEMENT FOR THE LEGACY modSolarEPCResource.
'   - Every legacy capability keeps working unchanged: NASA POWER hourly
'     import, range/month scheduler, RESOURCE_DB writes, Location auto-fill,
'     the VILLAGE_DB / Google / Nominatim / BigDataCloud label tiers and all
'     public macros (ConfigurePeriod, QueueImportedSite, RetryLatestSite,
'     FillLocations, MakeVillageDb, AddVillage, Resume, Stop, ShowLastError,
'     ResumePending, ProcessNext) - so ThisWorkbook, modSolarEPCCloudRelay
'     and every sheet button continue to compile.
'   - v2.5: when RESOURCE_DB has no row for a project, the watcher creates
'     it automatically (blank-only rules still apply); every step reports
'     its reason on the status bar; SolarEPC_DrawnLocationDebug produces a
'     read-only report of the whole chain.
'   - v3.11 FINAL: every user-facing message, status-bar line and debug
'     report is now professional English; the per-tick RESOURCE_DB scan
'     was replaced by one bulk column read per sweep (faster ticks);
'     K=12 range mode verified active (_CLOUD_CFG B12 = RANGE12).
'   - v3.8: SolarEPC_FillAndStampNow - one click starts the watcher, runs an
'     immediate sweep and back-fills Date/Time on existing rows from the
'     NASA Retrieval Date (UTC converted to local time). No reopen needed.
'   - v3.5: the Date/Time stamp is BLANK-ONLY and is applied on the FIRST
'     fill of a row (watcher auto-fill or NASA import, whichever happens
'     first). Accepted column names: Date/Time, Run Date-Time (Local),
'     Date-Time, DateTime. The debug macro reports whether the column was
'     found inside the table.
'   - v3.4: the GeoNames tier was REMOVED (its database provably contains no
'     Indian revenue villages). The user's own "Date/Time" column is filled
'     with the local run date/time; if absent, "Run Date-Time (Local)" is
'     created automatically. Every NASA fetch uses that row's OWN centroid -
'     two sites in the same town still receive independently fetched values.
'   - v3.2: SolarEPC_ResourceRefreshLabels safely upgrades previously filled
'     Location cells to village-level labels (suffix rule; manual entries
'     are never touched). Blank-only auto-fill is unchanged.
'   - v3.0: SINGLE MODULE - NASA pipeline + watcher + auto row creation +
'     diagnostics in one file.
'   - PLUS the v2.3 AUTOMATIC watcher: starts when the workbook opens and
'     inspects DRAWING_DATA.autoLWHTbl every 5 s; for every new or re-drawn
'     SITE row it writes the exact centroid (Worker-identical, computed
'     offline) into BLANK Latitude/Longitude cells and a descriptive
'     "Exact Area, Taluka, District" label into the BLANK Location cell.
'     No Alt+F8, no popups - a single status-bar line is the only feedback.
'
' IMPORT PROCEDURE (required - otherwise compile errors):
'   1. In the VBE, REMOVE any module named "modSolarEPCDrawnLocation".
'   2. REMOVE the old "modSolarEPCResource" module (export a backup first
'      if desired).
'   3. Right-click the project -> Import File... -> this .bas file.
'   4. Save and reopen the workbook.
'
' SAFETY RULES OF THE AUTO-FILL:
'   - only BLANK cells are written; manual entries, formulas and NASA
'     values are never overwritten
'   - NASA numbers, Data Status, cache_key and polygon_hash are untouched
'   - if no label can be resolved offline/online the lat/lon are still
'     filled; label retries are limited and no incorrect value is ever
'     written
'   - stop anytime: SolarEPC_DrawnLocationAutoStop ; immediate sweep:
'     SolarEPC_DrawnLocationAutoNow ; manual inspection:
'     SolarEPC_ResourceShowDrawnLocation
'
' CENTROID MATH = WORKER MATH (port of polygonCentroidLocalProjection):
' verified byte-for-byte on 5000 generated polygons + 7 edge cases
' (location-extract/verify/run_parity_check.sh).
'==========================================================================

'Watcher tuning (all other constants already exist inside the module).
Private Const AUTO_TICK_SECONDS As Long = 5
Private Const AUTO_NOROW_RETRIES As Long = 12    'retry ~1 min when the RESOURCE_DB row is missing
Private Const AUTO_LABEL_RETRIES As Long = 6     'limited retries while the label tier is offline


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
Private Const SETTINGS_SHEET As String = "SETTINGS"
'v2.0: user-owned village list. Beats every online source because the user
'knows the real revenue-village name, which no free geocoder carries.
Private Const VILLAGE_SHEET As String = "VILLAGE_DB"
Private Const VILLAGE_MAX_KM As Double = 5#
'v2.2: exact drawn-site extraction limits. These three values MUST stay equal
'to the Worker's validateSitePolygon() limits so that Excel and the Worker
'always agree on whether a drawing is valid and on its exact centroid.
Private Const SITE_MAX_VERTICES As Long = 1000
Private Const SITE_MIN_AREA_M2 As Double = 0.01
Private Const EARTH_RADIUS_M As Double = 6378137#

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
'v2.3 watcher state.
Private mAutoActive As Boolean
Private mAutoNextRun As Date
Private mProcessed As Object          'key = ReferenceID|Coordinates -> state
Private mFilledCache As Object          'ProjectID -> row already carries lat/lon

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

Private Function ResourcePeriodValid(ByVal StartText As String, ByVal EndText As String) As Boolean

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
        MsgBox "The RESOURCE_DB table was not found, or a required column is missing.", _
               vbExclamation, "Solar EPC Resource"
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
           "(Only blank cells were filled. Manually entered locations are never overwritten.)", _
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
    HTTP.setTimeouts 5000, 5000, 5000, FAST_REQUEST_TIMEOUT_SECONDS * 1000
    HTTP.setRequestHeader "Content-Type", "application/json"
    HTTP.setRequestHeader "Cache-Control", "no-store"
    HTTP.setRequestHeader "X-Solar-EPC-Workbook", mWorkbookID
    HTTP.setRequestHeader "X-Solar-EPC-Key", mWorkbookKey
    Body = "{""requestId"":""" & ResourceJSONEscape(RequestID) & _
           """,""month"":""" & ResourceJSONEscape(MonthText) & """}"
    HTTP.send Body

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
    HTTP.setTimeouts 5000, 5000, 5000, FAST_REQUEST_TIMEOUT_SECONDS * 1000
    HTTP.setRequestHeader "Content-Type", "application/json"
    HTTP.setRequestHeader "Cache-Control", "no-store"
    HTTP.setRequestHeader "X-Solar-EPC-Workbook", mWorkbookID
    HTTP.setRequestHeader "X-Solar-EPC-Key", mWorkbookKey
    Body = "{""requestId"":""" & ResourceJSONEscape(RequestID) & _
           """,""rangeStart"":""" & ResourceJSONEscape(StartMonth) & _
           """,""rangeEnd"":""" & ResourceJSONEscape(EndMonth) & """}"
    HTTP.send Body

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
                    mFastResponse(i) = CStr(mFastHTTP(i).ResponseText)
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
    HTTP.setTimeouts 5000, 5000, 5000, 60000
    HTTP.setRequestHeader "Content-Type", "application/json"
    HTTP.setRequestHeader "Cache-Control", "no-store"
    HTTP.setRequestHeader "X-Solar-EPC-Workbook", mWorkbookID
    HTTP.setRequestHeader "X-Solar-EPC-Key", mWorkbookKey
    If UCase$(MethodName) = "GET" Then HTTP.send Else HTTP.send Body
    StatusCode = HTTP.Status
    If StatusCode <> 204 Then ResourceRequestClient = HTTP.ResponseText
    Exit Function
Failed:
    StatusCode = 0
    ResourceRequestClient = vbNullString
End Function

'v1.9.2: generic read-only JSON GET for LOCATION LABEL ONLY (reverse geocoding).
'Never touches resource data. Tries WinHttp, then MSXML2.
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

'v1.9.3: join "Locality, Taluka, District" using ONLY the parts that exist.
'Both arguments may already be multi-part ("Phaltan, Satara"); empties are
'dropped so the label never shows a stray comma.
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

'v1.9.3: case-insensitive name comparison (used to collapse duplicates).
Private Function ResourceSameName(ByVal A As String, ByVal B As String) As Boolean
    ResourceSameName = (StrComp(Trim$(A), Trim$(B), vbTextCompare) = 0)
End Function

'v1.9.3: normalise one administrative name.
'Drops a trailing ", India" and generic Indian administrative suffixes so that
'"Phaltan taluka" -> "Phaltan" and "Satara district" -> "Satara".
'Only administrative names are cleaned - the locality is always left verbatim.
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

'v1.9.3: most specific LOCALITY actually returned by the reverse geocoder.
'Priority: village -> hamlet -> suburb -> locality -> neighbourhood -> town -> city.
'Returns "" when the geocoder exposes none of these, so the caller falls back to
'the next available level. A locality is NEVER invented.
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

'v1.9.5: TALUKA + DISTRICT (two slots). STATE is deliberately NOT included -
'the label is "Exact Area, Taluka, District" and nothing more.
'In India Nominatim returns the taluka as "county" ("Phaltan",
'"Paithan taluka") and the district as "state_district" ("Satara").
'Missing levels are simply skipped and names are never invented, so a point
'with no taluka still yields "Area, District".
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

'v1.9.4: read an optional _CLOUD_CFG cell. Never errors if the sheet or the
'cell is missing, so an unset key simply disables the optional tier.
Private Function ResourceConfigCell(ByVal CellAddress As String) As String
    ResourceConfigCell = ResourceSettingCell(CONFIG_SHEET, CellAddress)
End Function

'v1.9.7: read a cell from any sheet (case-insensitive sheet name). Never
'errors if the sheet or the cell is missing - returns "" instead.
'Used to read the Geocoding API key (SETTINGS!B11) and the Google Maps key
'(SETTINGS!B4) from the workbook's own SETTINGS sheet.
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

'v1.9.4: read one Google address_component by its type, e.g.
'  "locality"                     -> exact village / town
'  "administrative_area_level_3"  -> taluka (India)
'Google emits {"long_name":"..","short_name":"..","types":[..]} per component.
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

'v1.9.6: OPTIONAL Google reverse-geocode tier - the best shot at the EXACT
'Indian revenue village (OSM simply does not carry names like "Sonwadi Bk.").
'Active only when a key is found (SETTINGS!B11 -> SETTINGS!B4 -> _CLOUD_CFG!B5).
'With no key this returns "" and the Nominatim / BigDataCloud tiers run exactly
'as before.
'Every outcome is recorded in _CLOUD_CFG A14/B14 (status) and A15/B15 (detail)
'so a silent fallback can always be diagnosed instead of guessed at.
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

'v2.1: writes a label/value pair into a cell of the hidden _CLOUD_CFG sheet.
'Diagnostic only - no input field and no Excel Table is touched.
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

'v1.9.6: records why the Google tier succeeded / was skipped / failed into the
'hidden _CLOUD_CFG sheet. Diagnostic only - no input field or Table is touched.
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

'v2.0: distance in km between two WGS84 points (equirectangular approximation).
'Over a few km this is accurate to well under 0.1%, which is far more than
'enough for "which village is this site in".
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

'v2.0: THE MOST IMPORTANT TIER - the user's own VILLAGE_DB sheet.
'Layout (row 1 = headers, data from row 2):
'   A: Village   B: Taluka   C: District   D: Latitude   E: Longitude
'The closest village within VILLAGE_MAX_KM of the centroid wins. This is the
'only tier that reliably carries Indian revenue-village names such as
'"Sonwadi Bk.", which every free geocoder lacks. It needs no key, no
'billing and no internet.
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

'v2.0: one-time setup. Creates the VILLAGE_DB sheet with its headers.
Public Sub SolarEPC_ResourceMakeVillageDb()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(VILLAGE_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = VILLAGE_SHEET
        ws.Range("A1:E1").Value = Array("Village", "Taluka", "District", "Latitude", "Longitude")
        ws.Range("A1:E1").Font.Bold = True
        ws.Columns("A:C").ColumnWidth = 24
        ws.Columns("D:E").ColumnWidth = 14
        MsgBox "The VILLAGE_DB sheet has been created." & vbCrLf & vbCrLf & _
            "Add one row per village:" & vbCrLf & _
            "  A = Village    (e.g. Sonwadi Bk.)" & vbCrLf & _
            "  B = Taluka     (e.g. Phaltan)" & vbCrLf & _
            "  C = District   (e.g. Satara)" & vbCrLf & _
            "  D = Latitude   (right-click the village in Google Maps and" & vbCrLf & _
            "                  click the coordinates to copy them)" & vbCrLf & _
            "  E = Longitude" & vbCrLf & vbCrLf & _
            "Once a village is stored, every site drawn within 5 km of it is " & _
            "labelled automatically. No charge and no API key are required.", _
            vbInformation, "Solar EPC Resource"
    Else
        MsgBox "The VILLAGE_DB sheet already exists. (Villages stored: " & _
            CStr(ws.Cells(ws.Rows.Count, "A").End(xlUp).Row - 1) & ")", _
            vbInformation, "Solar EPC Resource"
    End If
End Sub

'v2.0: adds the last RESOURCE_DB site to VILLAGE_DB after asking for the
'village name. Taluka/District come from the Location already resolved by
'Nominatim, so only the village name has to be typed once.
Public Sub SolarEPC_ResourceAddVillage()
    Dim ws As Worksheet, vs As Worksheet, Tbl As ListObject
    Dim R As Long, NR As Long
    Dim cLoc As Long, cLat As Long, cLon As Long
    Dim LatitudeText As String, LongitudeText As String
    Dim Current As String, Parts As Variant
    Dim Village As String, Taluka As String, District As String

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

    'Last row that actually carries a centroid.
    R = 0
    Dim i As Long
    For i = Tbl.ListRows.Count To 1 Step -1
        If Len(Trim$(CStr(Tbl.ListRows(i).Range.Cells(1, cLat).Value2))) > 0 Then
            R = i
            Exit For
        End If
    Next i
    If R = 0 Then
        MsgBox "No RESOURCE_DB row with a latitude was found.", vbExclamation, "Solar EPC Resource"
        Exit Sub
    End If

    LatitudeText = Trim$(CStr(Tbl.ListRows(R).Range.Cells(1, cLat).Value2))
    LongitudeText = Trim$(CStr(Tbl.ListRows(R).Range.Cells(1, cLon).Value2))
    Current = Trim$(CStr(Tbl.ListRows(R).Range.Cells(1, cLoc).Value2))
    Parts = Split(Current, ",")
    If UBound(Parts) >= 1 Then
        Taluka = Trim$(CStr(Parts(1)))
        If UBound(Parts) >= 2 Then District = Trim$(CStr(Parts(2)))
    End If
    If Len(Taluka) = 0 And Len(District) = 0 And Len(Current) > 0 Then Taluka = Current

    Village = Trim$(InputBox( _
        "Enter the village or locality name for this site:" & vbCrLf & vbCrLf & _
        "  Latitude  : " & LatitudeText & vbCrLf & _
        "  Longitude : " & LongitudeText & vbCrLf & _
        "  Resolved  : " & Current & vbCrLf & vbCrLf & _
        "The name is saved to VILLAGE_DB and is reused automatically for " & _
        "every later site in the same area.", _
        "Solar EPC - Village name", vbNullString))
    If Len(Village) = 0 Then Exit Sub

    Set vs = ThisWorkbook.Worksheets(VILLAGE_SHEET)
    NR = vs.Cells(vs.Rows.Count, "A").End(xlUp).Row + 1
    If NR < 2 Then NR = 2
    vs.Cells(NR, "A").Value2 = Village
    vs.Cells(NR, "B").Value2 = Taluka
    vs.Cells(NR, "C").Value2 = District
    vs.Cells(NR, "D").Value2 = Val(LatitudeText)
    vs.Cells(NR, "E").Value2 = Val(LongitudeText)

    Tbl.ListRows(R).Range.Cells(1, cLoc).Value2 = _
        ResourceLocationCompose(Village, ResourceLocationCompose(Taluka, District))
    mLastLocKey = vbNullString
    mLastLocLabel = vbNullString

    MsgBox "Saved to VILLAGE_DB:" & vbCrLf & vbCrLf & _
        Village & IIf(Len(Taluka) > 0, ", " & Taluka, "") & _
        IIf(Len(District) > 0, ", " & District, ""), _
        vbInformation, "Solar EPC Resource"
    Exit Sub
Failed:
    MsgBox "The village could not be added: " & Err.Description & vbCrLf & vbCrLf & _
        "Run SolarEPC_ResourceMakeVillageDb first.", vbExclamation, "Solar EPC Resource"
End Sub

'v1.9.5: "EXACT AREA, TALUKA, DISTRICT" for the exact centroid
'(e.g. "Rahegaon, Vaijapur, Chhatrapati Sambhajinagar",
'"Sonwadi Bk., Phaltan, Satara"). STATE is NOT included.
'Tiers: 0) VILLAGE_DB (user's own sheet - free, offline, exact revenue
'village) -> 0b) Google (only if a working key is found) ->
'1) OSM Nominatim (free, no key) -> 2) BigDataCloud (free, no key).
'Only the exact centroid lat/lon is reverse-geocoded - there is no
'nearest-town search. Result is descriptive ONLY; identity stays
'centroid_lat/centroid_lon.
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

'==========================================================================
' v2.2 - EXACT DRAWN-SITE LOCATION (OFFLINE EXTRACTION)
'
' Answers: "what is the EXACT location of the place I drew on the map?"
'
' Everything below is READ-ONLY and works with no cloud configuration, no
' internet and no NASA/Worker call. The centroid algebra is the Worker's own
' polygonCentroidLocalProjection() ported statement by statement, so Excel and
' the Worker can never disagree about where the drawn site is.
'==========================================================================

'v2.2: shows the exact location of a drawn SITE - the area-weighted centroid
'(the authoritative coordinate), the drawn boundary corners, the site area and
'a Google Maps link for the centroid. Reads DRAWING_DATA.autoLWHTbl only.
Public Sub SolarEPC_ResourceShowDrawnLocation()
    Dim Site As Object
    Dim Answer As VbMsgBoxResult
    Dim Detail As String
    Dim MapsURL As String
    Dim i As Long

    On Error GoTo Failed
    Set Site = ResourceLatestSiteData()
    If Site Is Nothing Then
        MsgBox "No drawn SITE row was found in DRAWING_DATA." & vbCrLf & vbCrLf & _
            "Open the map, draw the site boundary and SAVE it first.", _
            vbExclamation, "Solar EPC - Exact Drawn Location"
        Exit Sub
    End If

    Detail = "EXACT DRAWN-SITE LOCATION" & vbCrLf & _
        "----------------------------------------" & vbCrLf & _
        "Project ID            : " & Site("ProjectID") & vbCrLf & _
        "Reference             : " & Site("ReferenceID") & vbCrLf & _
        "Boundary corners      : " & CStr(Site("VertexCount")) & vbCrLf & _
        "Site area             : " & Format$(CDbl(Site("AreaM2")), "#,##0.00") & " m2" & vbCrLf & vbCrLf & _
        "EXACT SITE CENTROID" & vbCrLf & _
        "Latitude  (" & ChrW(176) & ")       : " & ResourceDecimal(CDbl(Site("CentroidLat")), 8) & vbCrLf & _
        "Longitude (" & ChrW(176) & ")       : " & ResourceDecimal(CDbl(Site("CentroidLon")), 8) & vbCrLf & vbCrLf & _
        "Descriptive location  : " & Site("LocationLabel") & vbCrLf & _
        "Computed in Excel (offline) with the Worker's own centroid algorithm."

    MapsURL = "https://www.google.com/maps/search/?api=1&query=" & _
        ResourceDecimal(CDbl(Site("CentroidLat")), 8) & "," & _
        ResourceDecimal(CDbl(Site("CentroidLon")), 8)

    Answer = MsgBox(Detail & vbCrLf & vbCrLf & _
        "Yes    = open Google Maps at the exact centroid" & vbCrLf & _
        "No     = copy the coordinates to the clipboard" & vbCrLf & _
        "Cancel = close", _
        vbYesNoCancel + vbInformation, "Solar EPC - Exact Drawn Location")

    If Answer = vbYes Then
        ThisWorkbook.FollowHyperlink MapsURL
    ElseIf Answer = vbNo Then
        ResourceCopyToClipboard ResourceDecimal(CDbl(Site("CentroidLat")), 8) & ", " & _
            ResourceDecimal(CDbl(Site("CentroidLon")), 8)
        MsgBox "Copied to the clipboard:" & vbCrLf & vbCrLf & _
            ResourceDecimal(CDbl(Site("CentroidLat")), 8) & ", " & _
            ResourceDecimal(CDbl(Site("CentroidLon")), 8), _
            vbInformation, "Solar EPC - Exact Drawn Location"
    End If

    'Diagnostic only - the hidden _CLOUD_CFG sheet records what was extracted.
    ResourceSettingDiag "A18", "LAST DRAWN-SITE LOCATION", "B18", _
        Site("ProjectID") & " | " & ResourceDecimal(CDbl(Site("CentroidLat")), 8) & ", " & _
        ResourceDecimal(CDbl(Site("CentroidLon")), 8)
    Exit Sub
Failed:
    MsgBox "The drawn-site location could not be read." & vbCrLf & vbCrLf & _
        Err.Description, vbExclamation, "Solar EPC - Exact Drawn Location"
End Sub

'v2.2: fills BLANK RESOURCE_DB Latitude (deg) / Longitude (deg) cells from the
'exact drawn-site centroid, matched by Project ID, and then labels the row with
'the existing Location resolver. Existing values are NEVER overwritten and no
'resource value, status or NASA field is touched.
'Use it when a site was drawn but its resource run has not completed yet.
Public Sub SolarEPC_ResourceFillDrawnCentroids()
    Dim ws As Worksheet
    Dim Tbl As ListObject
    Dim R As Long
    Dim cProject As Long, cLat As Long, cLon As Long
    Dim RowRange As Range
    Dim ProjectID As String
    Dim Site As Object
    Dim Filled As Long, Skipped As Long
    Dim Answer As VbMsgBoxResult

    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(RESOURCE_SHEET)
    Set Tbl = ResourceFindHeaderTable(ws)
    If Tbl Is Nothing Then
        MsgBox "The RESOURCE_DB table was not found, or a required column is missing.", _
            vbExclamation, "Solar EPC - Exact Drawn Location"
        Exit Sub
    End If
    cProject = ResourceColumn(Tbl, "Project ID")
    cLat = ResourceColumn(Tbl, "Latitude (" & ChrW(176) & ")")
    cLon = ResourceColumn(Tbl, "Longitude (" & ChrW(176) & ")")
    If cProject = 0 Or cLat = 0 Or cLon = 0 Then
        MsgBox "RESOURCE_DB requires the Project ID, Latitude and Longitude columns.", _
            vbExclamation, "Solar EPC - Exact Drawn Location"
        Exit Sub
    End If

    Answer = MsgBox("Fill BLANK RESOURCE_DB Latitude/Longitude cells from the exact " & _
        "drawn-site centroid?" & vbCrLf & vbCrLf & _
        "Match: Project ID -> DRAWING_DATA.autoLWHTbl SITE row." & vbCrLf & _
        "Existing coordinates are never overwritten." & vbCrLf & _
        "No NASA value, Data Status or cache identity is changed.", _
        vbYesNo + vbQuestion, "Solar EPC - Exact Drawn Location")
    If Answer <> vbYes Then Exit Sub

    For R = 1 To Tbl.ListRows.Count
        Set RowRange = Tbl.ListRows(R).Range
        ProjectID = Trim$(CStr(RowRange.Cells(1, cProject).Value2))
        If Len(ProjectID) > 0 Then
            If Len(Trim$(CStr(RowRange.Cells(1, cLat).Value2))) = 0 And _
               Len(Trim$(CStr(RowRange.Cells(1, cLon).Value2))) = 0 Then

                Set Site = ResourceExactSiteData(ProjectID)
                If Not Site Is Nothing Then
                    With RowRange.Cells(1, cLat)
                        If Not .HasFormula Then .Value2 = CDbl(Site("CentroidLat"))
                    End With
                    With RowRange.Cells(1, cLon)
                        If Not .HasFormula Then .Value2 = CDbl(Site("CentroidLon"))
                    End With
                    'Descriptive label, blank cells only - the existing resolver.
                    ResourceFillLocationCell Tbl, RowRange, _
                        ResourceDecimal(CDbl(Site("CentroidLat")), 8), _
                        ResourceDecimal(CDbl(Site("CentroidLon")), 8)
                    Filled = Filled + 1
                Else
                    Skipped = Skipped + 1
                End If
            End If
        End If
    Next R

    MsgBox "Exact drawn-centroid fill complete." & vbCrLf & vbCrLf & _
        "Rows filled  : " & CStr(Filled) & vbCrLf & _
        "Rows skipped : " & CStr(Skipped) & " (no valid SITE drawing for that Project ID)" & vbCrLf & vbCrLf & _
        "Only blank coordinate cells were filled.", _
        vbInformation, "Solar EPC - Exact Drawn Location"
    Exit Sub
Failed:
    MsgBox "Exact drawn-centroid fill failed: " & Err.Description, _
        vbExclamation, "Solar EPC - Exact Drawn Location"
End Sub

'v2.2: PUBLIC read-only API. Returns the exact centroid of the SITE drawn for
'a Project ID. Usable from any module, add-in or the Immediate window:
'
'   Dim la As Double, lo As Double
'   If SolarEPC_ExactSiteLocation("PRJ-001", la, lo) Then Debug.Print la, lo
'
'Returns False (and leaves both outputs at 0) when the project has no valid
'SITE drawing. Never raises, never writes anything.
Public Function SolarEPC_ExactSiteLocation(ByVal ProjectID As String, _
    ByRef CentroidLatitude As Double, ByRef CentroidLongitude As Double) As Boolean

    Dim Site As Object
    On Error GoTo Failed
    CentroidLatitude = 0#
    CentroidLongitude = 0#
    'IncludeLabel:=False -> no reverse geocoding, no network, no cost.
    Set Site = ResourceExactSiteData(ProjectID, False)
    If Site Is Nothing Then Exit Function
    CentroidLatitude = CDbl(Site("CentroidLat"))
    CentroidLongitude = CDbl(Site("CentroidLon"))
    SolarEPC_ExactSiteLocation = True
    Exit Function
Failed:
    SolarEPC_ExactSiteLocation = False
End Function

'v2.2: PUBLIC read-only API - exact drawn area in square metres (0 on failure).
'The Worker's own local-projection area, so it matches the resource pipeline.
Public Function SolarEPC_ExactSiteArea(ByVal ProjectID As String) As Double
    Dim Site As Object
    On Error GoTo Failed
    Set Site = ResourceExactSiteData(ProjectID, False)
    If Site Is Nothing Then Exit Function
    SolarEPC_ExactSiteArea = CDbl(Site("AreaM2"))
    Exit Function
Failed:
    SolarEPC_ExactSiteArea = 0#
End Function

'v2.2: PUBLIC read-only API - one printable line for a report or a cell:
'   "Sonwadi Bk., Phaltan, Satara | 17.99088747, 74.43575012 | 11786.13 m2 | 4 corners"
'Returns "" when the project has no valid SITE drawing.
Public Function SolarEPC_ExactSiteLocationText(ByVal ProjectID As String) As String
    Dim Site As Object
    On Error GoTo Failed
    Set Site = ResourceExactSiteData(ProjectID)
    If Site Is Nothing Then Exit Function
    SolarEPC_ExactSiteLocationText = _
        Site("LocationLabel") & " | " & _
        ResourceDecimal(CDbl(Site("CentroidLat")), 8) & ", " & _
        ResourceDecimal(CDbl(Site("CentroidLon")), 8) & " | " & _
        Format$(CDbl(Site("AreaM2")), "#,##0.00") & " m2 | " & _
        CStr(Site("VertexCount")) & " corners"
    Exit Function
Failed:
    SolarEPC_ExactSiteLocationText = vbNullString
End Function

'v2.2: PUBLIC read-only API - the RAW drawn boundary of a SITE, offline.
'   BoundaryText = "lat,lon;lat,lon;..." exactly as the Worker receives it
'                  (8 decimals, duplicates and closing point already removed)
'   PolygonJSON  = [[lat,lon],[lat,lon],...] - the sitePolygon wire format
'   VertexCount  = number of corners actually drawn
'Returns False when the project has no valid SITE drawing.
'Use it to re-open the exact same boundary on the map, export it, or verify
'what was sent to the Worker.
Public Function SolarEPC_ExactSiteBoundary(ByVal ProjectID As String, _
    ByRef BoundaryText As String, ByRef PolygonJSON As String, _
    ByRef VertexCount As Long) As Boolean

    Dim Site As Object

    On Error GoTo Failed
    BoundaryText = vbNullString
    PolygonJSON = vbNullString
    VertexCount = 0
    Set Site = ResourceExactSiteData(ProjectID, False)
    If Site Is Nothing Then Exit Function

    VertexCount = CLng(Site("VertexCount"))
    BoundaryText = CStr(Site("CoordinateText"))
    'Same builder the resource request already uses -> identical wire format.
    PolygonJSON = ResourcePolygonJSON(BoundaryText)
    SolarEPC_ExactSiteBoundary = (Len(PolygonJSON) > 0)
    Exit Function
Failed:
    SolarEPC_ExactSiteBoundary = False
End Function

'v2.2: single internal extractor. Reads the SITE drawing for a Project ID and
'returns a Dictionary with the exact location, or Nothing when unavailable.
'Keys: ProjectID, ReferenceID, VertexCount, CentroidLat, CentroidLon, AreaM2,
'      LocationLabel, GoogleMapsURL, CoordinateText (the cleaned ring actually
'      used for the maths and sent to the Worker), RawCoordinateText (the cell
'      text exactly as the map stored it).
'Read-only: no sheet, table, cell or module state is modified.
'IncludeLabel=False keeps the call 100% offline (no reverse geocoding at all),
'which is what the coordinate-only public API uses.
Private Function ResourceExactSiteData(ByVal ProjectID As String, _
    Optional ByVal IncludeLabel As Boolean = True) As Object
    Dim Dict As Object
    Dim SiteRow As ListRow
    Dim CoordinateText As String
    Dim Latitudes() As Double
    Dim Longitudes() As Double
    Dim VertexCount As Long
    Dim CentroidLatitude As Double
    Dim CentroidLongitude As Double
    Dim AreaM2 As Double
    Dim ErrorText As String

    On Error GoTo Failed
    ProjectID = Trim$(ProjectID)
    If Len(ProjectID) = 0 Then Exit Function

    Set SiteRow = ResourceFindSiteRowByProject(ProjectID)
    If SiteRow Is Nothing Then Exit Function

    CoordinateText = ResourceRowCoordinateText(SiteRow)
    If Len(CoordinateText) = 0 Then Exit Function

    'Same parse + validation rules as the Worker's validateSitePolygon().
    If Not ResourceParseSitePolygon(CoordinateText, Latitudes, Longitudes, _
                                    VertexCount, ErrorText) Then Exit Function

    'Same area-weighted centroid as the Worker's polygonCentroidLocalProjection().
    If Not ResourceCentroidLocalProjection(Latitudes, Longitudes, VertexCount, _
        CentroidLatitude, CentroidLongitude, AreaM2, ErrorText) Then Exit Function
    If AreaM2 < SITE_MIN_AREA_M2 Then Exit Function

    Set Dict = CreateObject("Scripting.Dictionary")
    Dict.CompareMode = vbTextCompare
    Dict("ProjectID") = ProjectID
    Dict("ReferenceID") = ResourceRowReferenceText(SiteRow)
    Dict("RawCoordinateText") = CoordinateText
    'The cleaned ring actually used for the maths and sent to the Worker.
    Dict("CoordinateText") = ResourceBoundaryText(Latitudes, Longitudes, VertexCount)
    Dict("VertexCount") = VertexCount
    Dict("CentroidLat") = CentroidLatitude
    Dict("CentroidLon") = CentroidLongitude
    Dict("AreaM2") = AreaM2
    'Descriptive only - the identity stays the exact centroid coordinates.
    'Resolved lazily so a coordinate-only request never touches the network.
    If IncludeLabel Then
        Dict("LocationLabel") = ResourceLocationLabel( _
            ResourceDecimal(CentroidLatitude, 8), ResourceDecimal(CentroidLongitude, 8))
    Else
        Dict("LocationLabel") = vbNullString
    End If
    Dict("GoogleMapsURL") = "https://www.google.com/maps/search/?api=1&query=" & _
        ResourceDecimal(CentroidLatitude, 8) & "," & ResourceDecimal(CentroidLongitude, 8)
    Set ResourceExactSiteData = Dict
    Exit Function
Failed:
    Set ResourceExactSiteData = Nothing
End Function

'v2.2: the most recently imported SITE drawing, or Nothing. Used by
'SolarEPC_ResourceShowDrawnLocation so one click answers "where did I draw?".
Private Function ResourceLatestSiteData() As Object
    Dim Tbl As ListObject
    Dim R As Long
    Dim ProjectID As String
    On Error GoTo Failed
    Set Tbl = ResourceSiteTable()
    If Tbl Is Nothing Then Exit Function
    For R = Tbl.ListRows.Count To 1 Step -1
        If ResourceRowIsImportedSite(Tbl.ListRows(R)) Then
            ProjectID = ResourceRowProjectID(Tbl.ListRows(R))
            If Len(ProjectID) > 0 Then
                Set ResourceLatestSiteData = ResourceExactSiteData(ProjectID)
                'Keep scanning when the newest row exists but its drawing is
                'unreadable, so one bad row never hides the last good site.
                If Not ResourceLatestSiteData Is Nothing Then Exit Function
            End If
        End If
    Next R
    Exit Function
Failed:
    Set ResourceLatestSiteData = Nothing
End Function

'v2.2: DRAWING_DATA.autoLWHTbl, or Nothing when the sheet/table is absent.
Private Function ResourceSiteTable() As ListObject
    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(DRAWING_SHEET)
    Set ResourceSiteTable = ws.ListObjects(SITE_TABLE)
    Exit Function
Failed:
    Set ResourceSiteTable = Nothing
End Function

'v2.2: newest SITE row for a Project ID. Only rows imported from the map are
'accepted (Reference id "SITE-MAP-..."), so obstacle rows can never be used.
Private Function ResourceFindSiteRowByProject(ByVal ProjectID As String) As ListRow
    Dim Tbl As ListObject
    Dim R As Long
    On Error GoTo Failed
    Set Tbl = ResourceSiteTable()
    If Tbl Is Nothing Then Exit Function
    If ResourceColumn(Tbl, "Project ID") = 0 Then Exit Function
    If ResourceColumn(Tbl, "Coordinates") = 0 Then Exit Function
    For R = Tbl.ListRows.Count To 1 Step -1
        If ResourceRowIsImportedSite(Tbl.ListRows(R)) Then
            If StrComp(Trim$(CStr(ResourceRowProjectID(Tbl.ListRows(R)))), _
                       ProjectID, vbTextCompare) = 0 Then
                Set ResourceFindSiteRowByProject = Tbl.ListRows(R)
                Exit Function
            End If
        End If
    Next R
    Exit Function
Failed:
    Set ResourceFindSiteRowByProject = Nothing
End Function

'v2.2: a row is a MAP SITE only when its Reference id starts with "SITE-MAP-".
Private Function ResourceRowIsImportedSite(ByVal Row As ListRow) As Boolean
    Dim ReferenceID As String
    On Error GoTo Failed
    ReferenceID = ResourceRowReferenceText(Row)
    ResourceRowIsImportedSite = (UCase$(Left$(ReferenceID, 9)) = "SITE-MAP-")
    Exit Function
Failed:
    ResourceRowIsImportedSite = False
End Function

Private Function ResourceRowProjectID(ByVal Row As ListRow) As String
    ResourceRowProjectID = ResourceRowCellText(Row, "Project ID")
End Function

Private Function ResourceRowReferenceText(ByVal Row As ListRow) As String
    ResourceRowReferenceText = ResourceRowCellText(Row, "Reference id")
End Function

Private Function ResourceRowCoordinateText(ByVal Row As ListRow) As String
    ResourceRowCoordinateText = ResourceRowCellText(Row, "Coordinates")
End Function

'v2.2: read one named cell of a list row. Returns "" when the column is absent
'so a renamed/extra column can never raise inside a read-only helper.
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

'v2.2: parse the stored boundary text "lat,lon;lat,lon;..." into two Double
'arrays, applying EXACTLY the Worker's validateSitePolygon() rules:
'   1. every pair must be a plain decimal number inside the WGS84 range
'   2. every vertex is rounded to 8 decimals before any maths
'   3. consecutive duplicate vertices are dropped
'   4. a repeated closing vertex is dropped
'   5. at least 3 vertices and 3 distinct vertices must remain
'   6. at most SITE_MAX_VERTICES vertices
'Returns False with a readable reason when the drawing cannot be used.
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

'v2.2: rebuilds the cleaned boundary text "lat,lon;lat,lon;..." from the vertex
'arrays that were actually used for the centroid maths. This is byte-for-byte
'what the Worker receives as sitePolygon, so Excel and the Worker can never
'disagree about which corners define the drawn site.
Private Function ResourceBoundaryText(ByRef Latitudes() As Double, _
    ByRef Longitudes() As Double, ByVal VertexCount As Long) As String

    Dim ResultText As String
    Dim i As Long
    On Error GoTo Failed
    ResultText = vbNullString
    For i = 0 To VertexCount - 1
        If Len(ResultText) > 0 Then ResultText = ResultText & ";"
        ResultText = ResultText & ResourceDecimal(Latitudes(i), 8) & "," & _
                     ResourceDecimal(Longitudes(i), 8)
    Next i
    ResourceBoundaryText = ResultText
    Exit Function
Failed:
    ResourceBoundaryText = vbNullString
End Function

'v2.2: how many different vertices a ring really holds (Worker's Set check).
Private Function ResourceDistinctVertexCount(ByRef Latitudes() As Double, _
    ByRef Longitudes() As Double, ByVal VertexCount As Long) As Long

    Dim Keys As Object
    Dim i As Long
    On Error GoTo Failed
    Set Keys = CreateObject("Scripting.Dictionary")
    Keys.CompareMode = vbBinaryCompare
    For i = 0 To VertexCount - 1
        Keys(ResourceDecimal(Latitudes(i), 8) & "," & ResourceDecimal(Longitudes(i), 8)) = 1
    Next i
    ResourceDistinctVertexCount = Keys.Count
    Exit Function
Failed:
    ResourceDistinctVertexCount = 0
End Function

'v2.2: EXACT PORT of worker_k12.js polygonCentroidLocalProjection().
'Local equirectangular projection around the mean vertex, then the standard
'area-weighted polygon centroid, then back to WGS84. Statement order, constants
'and guards are identical to the Worker so both always produce the same numbers.
'Outputs: CentroidLatitude, CentroidLongitude (degrees) and AreaM2 (square metres).
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

'v2.2: copies text to the Windows clipboard (MSForms DataObject, late bound).
'Returns False when the clipboard is unavailable so the caller can say so.
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
    'v3.0: exact local run stamp (new column, auto-created once).
    ResourceStampLocalRunTime Tbl, Target
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
    Dim NewRow As ListRow

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

    For R = 1 To Tbl.ListRows.Count
        If Len(Trim$(CStr(Tbl.ListRows(R).Range.Cells(1, cProject).Value2))) = 0 Then
            Set ResourceProjectTableRow = Tbl.ListRows(R).Range
            Exit Function
        End If
    Next R
    Set NewRow = Tbl.ListRows.Add
    Set ResourceProjectTableRow = NewRow.Range
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

'--------------------------------------------------------------------------
' v2.3 - AUTOMATIC LOCATION FILL (no Alt+F8, no popups)
'
' The watcher starts when the workbook opens (Auto_Open) and inspects
' DRAWING_DATA.autoLWHTbl every AUTO_TICK_SECONDS. As soon as a new or
' modified SITE row appears:
'   1. the exact area-weighted centroid is computed offline
'   2. the RESOURCE_DB row for that Project ID is located
'   3. the exact centroid (identical to the Worker's) is written into the
'      BLANK Latitude / Longitude cells
'   4. a descriptive label ("Exact Area, Taluka, District") is written into
'      the BLANK Location cell - VILLAGE_DB -> Google -> Nominatim ->
'      BigDataCloud, using this workbook's existing rules
'
' Filled cells are NEVER overwritten, formula cells are skipped, and the
' NASA numbers/status inside RESOURCE_DB are never touched. The only
' feedback is a single status-bar line.
'--------------------------------------------------------------------------

Public Sub Auto_Open()
    On Error Resume Next
    SolarEPC_DrawnLocationAutoStart
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    SolarEPC_DrawnLocationAutoStop
End Sub

'Start the watcher and run one immediate sweep (already-drawn sites are filled too).
Public Sub SolarEPC_DrawnLocationAutoStart()
    On Error GoTo Failed
    If mAutoActive Then Exit Sub
    If mProcessed Is Nothing Then Set mProcessed = CreateObject("Scripting.Dictionary")
    mAutoActive = True
    DrawnLocationSweep
    DrawnLocationSchedule AUTO_TICK_SECONDS
    Application.StatusBar = "Solar EPC: drawn-site location AUTO-fill ON (scanning DRAWING_DATA every " & _
        CStr(AUTO_TICK_SECONDS) & " s)."
    Exit Sub
Failed:
    mAutoActive = False
End Sub

'Stop the watcher (cancels the scheduled OnTime).
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

'Run one sweep immediately without waiting - may be attached to a button.
Public Sub SolarEPC_DrawnLocationAutoNow()
    On Error GoTo Failed
    If mProcessed Is Nothing Then Set mProcessed = CreateObject("Scripting.Dictionary")
    DrawnLocationSweep
    Exit Sub
Failed:
End Sub

'Internal scheduler entry point (OnTime requires a Public name).
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

'Inspect all SITE rows once; auto-fill any new or re-drawn site.
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
    Set mFilledCache = ResourceBuildFilledMap()

    For R = 1 To Tbl.ListRows.Count
        ReferenceText = RowCellText(Tbl, R, "Reference id")
        If UCase$(Left$(ReferenceText, 9)) = "SITE-MAP-" Then
            ProjectID = RowCellText(Tbl, R, "Project ID")
            CoordinateText = RowCellText(Tbl, R, "Coordinates")
            If Len(ProjectID) > 0 And Len(CoordinateText) > 0 Then
                'The key includes the coordinates: re-drawing the site changes
                'the key and the fill is attempted again.
                KeyText = ReferenceText & "|" & CoordinateText
                If Not mProcessed.Exists(KeyText) Or _
                   Left$(CStr(mProcessed(KeyText)), 5) = "NOROW" Or _
                   Left$(CStr(mProcessed(KeyText)), 7) = "LOCPEND" Or _
                   Not ProjectFilledInCache(ProjectID) Then

                    StateText = CStr(mProcessed(KeyText) & "")
                    Attempts = DrawnLocationAttempts(StateText)

                    If ParsePolygon(CoordinateText, Latitudes, Longitudes, VertexCount, ErrorText) Then
                        If CentroidLocalProjection(Latitudes, Longitudes, VertexCount, _
                               CentroidLatitude, CentroidLongitude, AreaM2, ErrorText) Then
                            If AreaM2 >= SITE_MIN_AREA_M2 Then
                                ResultText = FillResourceDbForSite(ProjectID, CentroidLatitude, _
                                    CentroidLongitude, Attempts >= 2)
                                If InStr(1, ResultText, "LOCPEND", vbBinaryCompare) > 0 Then
                                    'lat/lon filled, label tier offline - limited retries.
                                    Application.StatusBar = "Solar EPC: " & ProjectID & _
                                        " lat/lon filled; Location label offline, retry " & _
                                        CStr(Attempts + 1) & "/" & CStr(AUTO_LABEL_RETRIES) & _
                                        " (check internet / VILLAGE_DB / geocoding key)"
                                    If Attempts < AUTO_LABEL_RETRIES Then _
                                        mProcessed(KeyText) = "LOCPEND:" & CStr(Attempts + 1)
                                ElseIf ResultText = "NOROW" Then
                                    'RESOURCE_DB row does not exist yet - retry, then auto-create.
                                    Application.StatusBar = "Solar EPC: no RESOURCE_DB row for " & ProjectID & _
                                        " (" & CStr(Attempts + 1) & ") - " & _
                                        IIf(Attempts + 1 >= 2, "creating the row now...", "retrying...")
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

'Extract the attempt count from states such as "NOROW:3" / "LOCPEND:1".
Private Function DrawnLocationAttempts(ByVal StateText As String) As Long
    Dim P As Long
    P = InStr(1, StateText, ":", vbBinaryCompare)
    If P > 0 Then DrawnLocationAttempts = Val(Mid$(StateText, P + 1))
End Function

'Record of the last auto-fill in the hidden _CLOUD_CFG sheet
'(diagnostic only; silently skipped when the sheet is absent).
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

'Locate the RESOURCE_DB row for this Project ID and fill BLANK lat/lon +
'BLANK Location. Returns what was filled, e.g.
'"lat lon location(Sonwadi Bk., Phaltan, Satara)", or SKIP-NO-BLANK /
'NOROW / LOCPEND.
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
    'whose lat/lon already equal this centroid (Location may still be blank).
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
    'v2.5: create the missing row - a drawn site must appear in RESOURCE_DB.
    'Remaining columns are filled by the NASA import or manually.
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

    If Len(DidText) > 0 And DidText <> "SKIP-NO-BLANK" Then _
        ResourceStampLocalRunTime Tbl, Target
    If Len(DidText) = 0 Then DidText = "SKIP-NO-BLANK"
    FillResourceDbForSite = DidText
    Exit Function
Failed:
    FillResourceDbForSite = "NOROW"
End Function

'The resource_db Excel Table by name, or Nothing.
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

'Column index by header name; 0 when the column does not exist.
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
 v2.4 adapters - connect the watcher's named helpers to the proven
' helpers of the complete module (no duplicated logic).
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

'--------------------------------------------------------------------------
' v2.5 - DIAGNOSTIC: reports where the chain is stuck, in one click.
'   Alt+F8 -> SolarEPC_DrawnLocationDebug
' Changes NOTHING (read-only report).
'--------------------------------------------------------------------------
Public Sub SolarEPC_DrawnLocationDebug()
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
    Dim cDT As Long
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
        P1 = P1 & "   >> Watcher is OFF. Run SolarEPC_DrawnLocationAutoStart" & vbCrLf & _
                  "      or save and REOPEN the workbook (Auto_Open starts it)." & vbCrLf
    End If

    Set Tbl = DrawnSiteTable()
    If Tbl Is Nothing Then
        P1 = P1 & "2) DRAWING_DATA / '" & SITE_TABLE & "' table NOT FOUND." & vbCrLf & _
                  "   >> Check the sheet/table name." & vbCrLf
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
            P1 = P1 & "   >> No SITE-MAP-* row found - was the drawing SAVED on the map?" & vbCrLf
        Else
            P1 = P1 & "3) Latest SITE: " & ProjectID & "  (" & CStr(Len(CoordinateText)) & " char coordinates)" & vbCrLf
            If Len(CoordinateText) = 0 Then
                P1 = P1 & "   >> Coordinates cell is BLANK - the map SAVE was incomplete." & vbCrLf
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
        P2 = "4) The 'resource_db' table was NOT FOUND in RESOURCE_DB." & vbCrLf & _
             "   >> The table name must be exactly 'resource_db'." & vbCrLf
    Else
        cProject = TableColumn(Rdb, "Project ID")
        cLat = TableColumn(Rdb, "Latitude (" & ChrW(176) & ")")
        cLon = TableColumn(Rdb, "Longitude (" & ChrW(176) & ")")
        cLoc = TableColumn(Rdb, "Location")
        cDT = ResourceDateTimeColumn(Rdb)
        P2 = "4) resource_db table: " & CStr(Rdb.ListRows.Count) & " row(s); columns: " & _
             IIf(cProject > 0, "P", "-") & IIf(cLat > 0, "Lat", "-") & _
             IIf(cLon > 0, "Lon", "-") & IIf(cLoc > 0, "Loc", "-") & vbCrLf
        P2 = P2 & "   Date/Time column: " & IIf(cDT > 0, _
             "found (" & Rdb.ListColumns(cDT).Name & ")", _
             "NOT INSIDE THE TABLE - add it inside the header row") & vbCrLf
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
                P2 = P2 & "5) >> No BLANK-lat/lon row found for " & ProjectID & "." & vbCrLf & _
                          "   (v2.5+ creates such a row automatically while the watcher is ON)" & vbCrLf
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
             "NO TIER RESOLVED (VILLAGE_DB sheet / internet / geocoding key)") & vbCrLf
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
        P2 = P2 & "   Date/Time stamp: " & CStr(ws.Range("B21").Value2) & vbCrLf
        End If
    End If
    P2 = P2 & vbCrLf & "This report has been copied to the clipboard."

    ResourceCopyToClipboard P1 & vbCrLf & P2
    MsgBox P1, vbInformation, "Solar EPC Debug 1/2"
    MsgBox P2, vbInformation, "Solar EPC Debug 2/2"
End Sub

'--------------------------------------------------------------------------
' v3.0: RESOURCE_DB me ek extra column "Run Date-Time (Local)" -
' jis DIN aur TIME (aapke computer ki local ghadi) pe NASA summary import
' stamp. The Worker's "Retrieval Date" keeps the UTC fetch time;
' this column is your own local run-time, refreshed on every import.
' If the column is absent it is created on the first import (end of table).
'--------------------------------------------------------------------------
Private Function ResourceDateTimeColumn(ByVal Tbl As ListObject) As Long
    Dim Names As Variant
    Dim i As Long
    Names = Array("Date/Time", "Run Date-Time (Local)", "Date-Time", _
                  "DateTime", "Date / Time")
    For i = LBound(Names) To UBound(Names)
        ResourceDateTimeColumn = ResourceColumn(Tbl, CStr(Names(i)))
        If ResourceDateTimeColumn > 0 Then Exit Function
    Next i
End Function

'v3.5: BLANK-ONLY local run stamp. Whichever event writes the first data
'into a row (watcher auto-fill or NASA import) stamps the local date and
'time. Never overwrites - a manual date entry is sacred.
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
    If C = 0 Then
        ResourceSettingDiag "A21", "DATE/TIME STAMP", "B21", "NO COLUMN"
        Exit Sub
    End If
    Set Cell = Target.Cells(1, C)
    If Cell Is Nothing Then Exit Sub
    If Cell.HasFormula Then
        ResourceSettingDiag "A21", "DATE/TIME STAMP", "B21", "FORMULA CELL (row " & CStr(Target.Row) & ")"
        Exit Sub
    End If
    If Len(Trim$(CStr(Cell.Value2))) > 0 Then     'blank-only, hamesha
        ResourceSettingDiag "A21", "DATE/TIME STAMP", "B21", _
            "NOT BLANK row " & CStr(Target.Row) & ": value=[" & Left$(CStr(Cell.Value2), 40) & "]"
        Exit Sub
    End If
    ResourceWriteDateTimeCell Cell, Now
    ResourceSettingDiag "A21", "DATE/TIME STAMP", "B21", _
        "OK row " & CStr(Target.Row) & " = " & Format$(Now, "dd-mm-yyyy hh:nn:ss")
End Sub


'--------------------------------------------------------------------------
' v3.2: safely upgrade previously filled Location rows when a better tier
' (VILLAGE_DB) becomes available. The blank-only rule is sacred for manual
' cells, so this macro runs ONLY on user demand and writes only when the
' new label is exactly the "village-plus" version of the old one (the new
' label ends with ", " & old label).
'   "Phaltan, Satara" -> "Sonwadi Bk., Phaltan, Satara"   = upgrade OK
'   "My Farm, Phaltan" (manual)                          = skipped, never touched
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
                "/" & CStr(Total) & " (manual entries are never overwritten)..."
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
        "Only rows were changed where the new label is the exact " & _
        "'village + ' version of the old one. Manual entries and everything else untouched.", _
        vbInformation, "Solar EPC Resource"
    Exit Sub
Failed:
    Application.StatusBar = False
    MsgBox "Label refresh failed: " & Err.Description, vbExclamation, "Solar EPC Resource"
End Sub

'v3.11: watcher session-memory self-heal, FAST edition. One bulk column read
'per sweep builds a ProjectID -> filled map; a DONE key is re-processed only
'when its row's lat/lon are blank again (user cleared them or an earlier
'session left them empty). No per-cell COM calls, no reopen needed.
Private Function ResourceBuildFilledMap() As Object
    Dim D As Object
    Dim Rdb As ListObject
    Dim cP As Long, cLat As Long, cLon As Long
    Dim ProjArr As Variant, LatArr As Variant, LonArr As Variant
    Dim n As Long, i As Long

    Set D = CreateObject("Scripting.Dictionary")
    D.CompareMode = 1                                  'vbTextCompare
    On Error GoTo Failed
    Set Rdb = ResourceDbTable()
    If Rdb Is Nothing Then GoTo DoneMap
    cP = TableColumn(Rdb, "Project ID")
    cLat = TableColumn(Rdb, "Latitude (" & ChrW(176) & ")")
    cLon = TableColumn(Rdb, "Longitude (" & ChrW(176) & ")")
    If cP = 0 Or cLat = 0 Or cLon = 0 Then GoTo DoneMap
    n = Rdb.ListRows.Count
    If n = 0 Then GoTo DoneMap
    ProjArr = Rdb.ListColumns(cP).Range.Value2         'one bulk read (incl. header)
    LatArr = Rdb.ListColumns(cLat).Range.Value2
    LonArr = Rdb.ListColumns(cLon).Range.Value2
    For i = 2 To n + 1
        If Len(Trim$(CStr(ProjArr(i, 1)))) > 0 Then
            If Len(Trim$(CStr(LatArr(i, 1)))) > 0 And _
               Len(Trim$(CStr(LonArr(i, 1)))) > 0 Then
                D(CStr(ProjArr(i, 1))) = True
            End If
        End If
    Next i
DoneMap:
    Set ResourceBuildFilledMap = D
    Exit Function
Failed:
    Set ResourceBuildFilledMap = D
End Function

Private Function ProjectFilledInCache(ByVal ProjectID As String) As Boolean
    If mFilledCache Is Nothing Then Exit Function
    ProjectFilledInCache = mFilledCache.Exists(ProjectID)
End Function

'--------------------------------------------------------------------------
' v3.8: THE BIG RED BUTTON. One click performs everything immediately -
' no session memory, no reopen, no waiting:
'   1. starts the watcher if it is OFF
'   2. runs an immediate sweep (blank lat/lon + Location + Date/Time stamp)
'   3. back-fills Date/Time on previously filled rows by converting the
'      NASA "Retrieval Date" (UTC) into the user's LOCAL time (the true
'      run time, not a fake Now)
'   4. reports everything in _CLOUD_CFG B21 and a MsgBox
'--------------------------------------------------------------------------

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
        MsgBox "The resource_db table was not found.", vbExclamation, "Solar EPC"
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
        MsgBox "The Date/Time or Latitude/Longitude columns were not found.", _
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
                        ResourceWriteDateTimeCell Cell, LocalTime
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
        "Left blank (lat/lon or Retrieval Date missing): " & CStr(LeftBlank) & vbCrLf & vbCrLf & _
        "The watcher is ON - new drawings will fill and stamp automatically.", _
        vbInformation, "Solar EPC Resource"
    Exit Sub
Failed:
    MsgBox "FillAndStampNow failed: " & Err.Description, vbExclamation, "Solar EPC Resource"
End Sub

'v3.10: NumberFormat can be blocked (sheet protection / table column
'setting). Try the format first; if Excel refuses, write clean text - the
'VALUE always lands, we never depend on the format.
Private Sub ResourceWriteDateTimeCell(ByVal Cell As Range, ByVal WhenTime As Date)
    Dim FmtOk As Boolean
    On Error Resume Next
    Cell.NumberFormat = "dd-mm-yyyy hh:nn:ss"
    FmtOk = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    If FmtOk Then
        Cell.Value2 = WhenTime
    Else
        Cell.Value2 = Format$(WhenTime, "dd-mm-yyyy hh:nn:ss")
    End If
End Sub
