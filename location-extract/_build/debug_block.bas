'--------------------------------------------------------------------------
' v2.5 - DIAGNOSTIC: chain kaha atki hai, ek click me batata hai.
'   Alt+F8 -> SolarEPC_DrawnLocationDebug
' Kuch bhi change NAHI karta (sirf padhta hai + report dikhata hai).
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

    CopyToClipboard P1 & vbCrLf & P2
    MsgBox P1, vbInformation, "Solar EPC Debug 1/2"
    MsgBox P2, vbInformation, "Solar EPC Debug 2/2"
End Sub
