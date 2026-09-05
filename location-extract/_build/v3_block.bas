'--------------------------------------------------------------------------
' v3.0: RESOURCE_DB me ek extra column "Run Date-Time (Local)" -
' jis DIN aur TIME (aapke computer ki local ghadi) pe NASA summary import
' hua, exact stamp. Worker ka "Retrieval Date" UTC fetch-time rakhta hai;
' ye column aapka apna local run-time hai, har import pe fresh update.
' Column na ho to pehli import par apne aap ban jaata hai (table ke end me).
'--------------------------------------------------------------------------
Private Sub ResourceStampLocalRunTime(ByVal Tbl As ListObject, ByVal Target As Range)
    Dim ColName As String
    Dim C As Long
    Dim Col As ListColumn
    Dim Cell As Range

    On Error Resume Next
    ColName = "Run Date-Time (Local)"
    C = ResourceColumn(Tbl, ColName)
    If C = 0 Then
        Set Col = Tbl.ListColumns.Add(Tbl.ListColumns.Count + 1)
        If Not Col Is Nothing Then Col.Name = ColName
        C = ResourceColumn(Tbl, ColName)
    End If
    On Error GoTo 0
    If C = 0 Then Exit Sub
    Set Cell = Target.Cells(1, C)
    If Cell Is Nothing Then Exit Sub
    If Cell.HasFormula Then Exit Sub
    Cell.NumberFormat = "@"
    Cell.Value2 = Format$(Now, "dd-mm-yyyy hh:nn:ss")
End Sub
