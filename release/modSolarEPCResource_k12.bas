Option Explicit

'==========================================================================
' SOLAR EPC - NASA POWER HOURLY RESOURCE MODULE
' Version 2.3
'
' PURPOSE
'   Imports NASA POWER hourly solar and weather resource data for a site that
'   was drawn on the map, validates it, and records a summary row in the
'   RESOURCE_DB table.
'
' SCOPE
'   - Reads only the SITE row in DRAWING_DATA.autoLWHTbl. The obstacle table
'     (tblDrawingData) is never used for the resource centroid.
'   - Sends the complete SITE polygon to the authenticated Worker.
'   - The Worker validates the polygon, computes its area-weighted centroid,
'     downloads the configured resource period in hourly chunks, validates
'     every record, and returns a project summary.
'   - T2M minimum and maximum are derived from validated hourly T2M because
'     the NASA POWER Hourly API rejects T2M_MIN and T2M_MAX.
'==========================================================================
'
' LOCATION RESOLUTION
'   The descriptive "Location" column of RESOURCE_DB is filled automatically
'   from the exact centroid, in the format:
'
'       "Exact Area, Taluka, District"
'       e.g. "Sonwadi Bk., Phaltan, Satara"
'            "Rahegaon, Vaijapur, Chhatrapati Sambhajinagar"
'
'   STATE is deliberately excluded. No level is ever invented: any level a
'   resolver cannot supply is simply skipped.
'
'   Resolution order (first match wins):
'
'     0)  VILLAGE_DB        A sheet in this workbook that lists the villages
'                           you work in. Free, offline, no key, no billing.
'                           This is the only source that carries Indian
'                           revenue village names such as "Sonwadi Bk.",
'                           which every free geocoder lacks.
'     0b) Google Geocoding  Reverse geocode through the Google Geocoding API.
'                           Best Indian revenue-village coverage, but the
'                           Cloud project must have billing enabled.
'     1)  OSM Nominatim     Free, no key required.
'     2)  BigDataCloud      Free, no key required.
'
'   Only the exact centroid is reverse geocoded - there is no nearest-town
'   search. The label is descriptive only; the authoritative identity of a
'   row remains its centroid latitude and longitude.
'==========================================================================
'
' CONFIGURATION
'
'   Google API key (optional). Resolved in this order:
'       1. SETTINGS!B4      the workbook Google Maps key
'       2. SETTINGS!B11     spare slot for a second Google key
'       3. _CLOUD_CFG!B5    manual override
'   When no key is present that tier is skipped and the free tiers run.
'   NOTE: Google requires billing to be enabled on the Cloud project. Without
'   it the API answers REQUEST_DENIED and the free tiers below run instead.
'
'   IMPORTANT - manual corrections are permanent. A Location cell that already
'   holds text is never changed, so anything you type by hand survives every
'   later import and every back-fill run. To have a cell re-resolved, clear it
'   first, then run SolarEPC_ResourceFillLocations.
'
'   VILLAGE_DB sheet (optional). Row 1 holds the headers, data starts at row 2:
'       A = Village    B = Taluka    C = District
'       D = Latitude   E = Longitude
'   The nearest village within 5 km of the centroid wins. To obtain a
'   village's coordinates: right-click it in Google Maps and click the
'   coordinate pair that appears - it is copied to the clipboard.
'
'   Diagnostics are written to the hidden _CLOUD_CFG sheet:
'       A10/B10  last resource-start error timestamp
'       A11/B11  last resource-start error detail
'       A12/B12  processing mode (RANGE12 / RANGE6 / MONTH)
'       A13/B13  last range processing time (ms)
'       A14/B14  Google geocode status
'       A15/B15  Google geocode detail
'       A16/B16  VILLAGE_DB match
'       A17/B17  Google key source
'       A20/B20  resolver that produced the label
'==========================================================================
'
' PUBLIC MACROS
'   SolarEPC_ResourceConfigurePeriod    Set the NASA POWER date range.
'   SolarEPC_ResourceQueueImportedSite  Queue a resource run for a SITE.
'   SolarEPC_ResourceRetryLatestSite    Re-run the most recent imported SITE.
'   SolarEPC_ResourceFillLocations      Back-fill blank Location cells.
'   SolarEPC_ResourceMakeVillageDb      Create the VILLAGE_DB sheet.
'   SolarEPC_ResourceAddVillage         Record the village for the latest site.
'   SolarEPC_ResourceResume             Retry after an outage.
'   SolarEPC_ResourceStop               Stop scheduled processing.
'   SolarEPC_ResourceShowLastError      Display the last recorded error.
'   SolarEPC_ResourceResumePending      Auto-resume when the workbook opens.
'   SolarEPC_ResourceProcessNext        Scheduler entry point (internal).
'==========================================================================
'
' VERSION HISTORY
'   v2.3    Removed the Geoapify tier. A head-to-head measurement on three
'           Maharashtra sites showed Geoapify matched OSM Nominatim on two
'           and was WRONG on the third (it returned a railway-station hamlet
'           picked up from a nearby school), so it added rate-limit headroom
'           but no accuracy. SETTINGS!B11 is now a spare slot for a second
'           Google key: the Google tier resolves B4 -> B11 -> _CLOUD_CFG!B5.
'           _CLOUD_CFG A18/B18 and A19/B19 are no longer written.
'   v2.2    Added the Geoapify reverse-geocode tier (removed again in v2.3).
'   v2.1    Geocoding key is read from SETTINGS!B11 (a dedicated key), then
'           SETTINGS!B4, then _CLOUD_CFG!B5. The key source actually used is
'           reported in _CLOUD_CFG A17/B17.
'   v2.0    Added the VILLAGE_DB tier - an offline, user-owned village list
'           that supplies Indian revenue village names which no free
'           geocoder carries.
'   v1.9.7  Google key read from SETTINGS!B4 instead of _CLOUD_CFG!B5.
'   v1.9.6  Google tier diagnostics written to _CLOUD_CFG A14/B14, A15/B15.
'   v1.9.5  Location format is "Exact Area, Taluka, District".
'   v1.9.2  Location auto-fill added; SolarEPC_ResourceFillLocations.
'   v1.9.1  RESOURCE_DB columns matched by name; extra columns allowed.
'   v1.8    Range mode (K=6/K=12) with automatic month-by-month fallback.
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
Private Const SETTINGS_SHEET As String = "SETTINGS"
'v2.3: spare slot for a second Google key. The Google tier resolves
'SETTINGS!B4 -> SETTINGS!B11 -> _CLOUD_CFG!B5 and reports which one it used
'in _CLOUD_CFG A17/B17.
Private Const GOOGLE_SPARE_KEY_CELL As String = "B11"
'v2.0: user-owned village list. Beats every online source because the user
'knows the real revenue-village name, which no free geocoder carries.
Private Const VILLAGE_SHEET As String = "VILLAGE_DB"
Private Const VILLAGE_MAX_KM As Double = 5#

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
    Dim I As Long, N As Long
    S = Trim$(TextValue)
    If Len(S) = 0 Then Exit Function
    If UCase$(Right$(S, 7)) = ", INDIA" Then S = Trim$(Left$(S, Len(S) - 7))
    Suffixes = Array(" taluka", " taluk", " tehsil", " tahsil", " subdivision", _
                     " mandal", " block", " district")
    For I = LBound(Suffixes) To UBound(Suffixes)
        N = Len(CStr(Suffixes(I)))
        If Len(S) > N Then
            If UCase$(Right$(S, N)) = UCase$(CStr(Suffixes(I))) Then
                S = Trim$(Left$(S, Len(S) - N))
                Exit For
            End If
        End If
    Next I
    ResourceAdminClean = S
End Function

'v1.9.3: most specific LOCALITY actually returned by the reverse geocoder.
'Priority: village -> hamlet -> suburb -> locality -> neighbourhood -> town -> city.
'Returns "" when the geocoder exposes none of these, so the caller falls back to
'the next available level. A locality is NEVER invented.
Private Function ResourceLocalityName(ByVal JSONText As String) As String
    Dim Keys As Variant
    Dim I As Long, V As String
    Keys = Array("village", "hamlet", "suburb", "locality", _
                 "neighbourhood", "town", "city")
    For I = LBound(Keys) To UBound(Keys)
        V = Trim$(ResourceJSONValue(JSONText, CStr(Keys(I))))
        If Len(V) > 0 Then
            ResourceLocalityName = V
            Exit Function
        End If
    Next I
End Function

'v1.9.5: TALUKA + DISTRICT (two slots). STATE is deliberately NOT included -
'the label is "Exact Area, Taluka, District" and nothing more.
'In India Nominatim returns the taluka as "county" ("Phaltan",
'"Paithan taluka") and the district as "state_district" ("Satara").
'Missing levels are simply skipped and names are never invented, so a point
'with no taluka still yields "Area, District".
Private Function ResourceAdminNames(ByVal JSONText As String, ByVal Locality As String) As String
    Dim Keys As Variant
    Dim I As Long, V As String
    Dim A1 As String, A2 As String
    Keys = Array("county", "district", "state_district")   'NOTE: "state" excluded on purpose
    For I = LBound(Keys) To UBound(Keys)
        V = ResourceAdminClean(ResourceJSONValue(JSONText, CStr(Keys(I))))
        If Len(V) > 0 Then
            If Not ResourceSameName(V, Locality) Then
                If Len(A1) = 0 Then
                    A1 = V
                ElseIf Not ResourceSameName(V, A1) Then
                    If Len(A2) = 0 Then A2 = V Else Exit For
                End If
            End If
        End If
    Next I
    ResourceAdminNames = ResourceLocationCompose(A1, A2)
End Function

'v1.9.4: read an optional _CLOUD_CFG cell. Never errors if the sheet or the
'cell is missing, so an unset key simply disables the optional tier.
Private Function ResourceConfigCell(ByVal CellAddress As String) As String
    ResourceConfigCell = ResourceSettingCell(CONFIG_SHEET, CellAddress)
End Function

'v1.9.7: read a cell from any sheet (case-insensitive sheet name). Never
'errors if the sheet or the cell is missing - returns "" instead.
'Used to read the Google Maps key (SETTINGS!B4), the spare Google key
'(SETTINGS!B11) and the manual override (_CLOUD_CFG!B5) from this workbook.
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
'Active only when a key is found (SETTINGS!B4 -> SETTINGS!B11 -> _CLOUD_CFG!B5).
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
    'Resolve the Google key: SETTINGS!B4 (workbook map key), then
    'SETTINGS!B11 (spare slot), then _CLOUD_CFG!B5 (manual override).
    'The first non-empty value wins.
    Key = ResourceSettingCell(SETTINGS_SHEET, "B4")
    KeySource = "SETTINGS!B4"
    If Len(Key) = 0 Then
        Key = ResourceSettingCell(SETTINGS_SHEET, GOOGLE_SPARE_KEY_CELL)
        KeySource = "SETTINGS!" & GOOGLE_SPARE_KEY_CELL
    End If
    If Len(Key) = 0 Then
        Key = ResourceConfigCell("B5")
        KeySource = "_CLOUD_CFG!B5"
    End If
    If Len(Key) = 0 Then
        KeySource = "NONE"
        ResourceGoogleStatus "NO_KEY", _
            "No Google key found in SETTINGS!B4 or _CLOUD_CFG!B5, so the Google tier " & _
            "is skipped. Put a Google key in SETTINGS!B4 or SETTINGS!B11 " & _
            "(or in _CLOUD_CFG!B5). Remember: Google requires billing to be " & _
            "enabled on the Cloud project, otherwise it returns REQUEST_DENIED."
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
    Dim I As Long
    For I = Tbl.ListRows.Count To 1 Step -1
        If Len(Trim$(CStr(Tbl.ListRows(I).Range.Cells(1, cLat).Value2))) > 0 Then
            R = I
            Exit For
        End If
    Next I
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
'village) -> 0b) Google (only if a working key with billing is found) ->
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
    If Len(ResourceLocationLabel) > 0 Then
        ResourceSettingDiag "A20", "LOCATION RESOLVER", "B20", "VILLAGE_DB"
        GoTo StoreCache
    End If

    '0b) OPTIONAL Google tier - best revenue-village data, but the Cloud
    '    project must have billing enabled or it returns REQUEST_DENIED.
    ResourceLocationLabel = ResourceGoogleLocationLabel(Lat, Lon)
    If Len(ResourceLocationLabel) > 0 Then
        ResourceSettingDiag "A20", "LOCATION RESOLVER", "B20", "GOOGLE"
        GoTo StoreCache
    End If

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
            ResourceSettingDiag "A20", "LOCATION RESOLVER", "B20", "NOMINATIM"
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
        If Len(Primary) > 0 Then
            ResourceLocationLabel = Primary
            ResourceSettingDiag "A20", "LOCATION RESOLVER", "B20", "BIGDATACLOUD"
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
