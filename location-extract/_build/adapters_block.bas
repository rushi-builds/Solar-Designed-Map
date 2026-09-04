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
