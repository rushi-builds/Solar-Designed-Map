import io
p = 'watcher_block.bas'
s = io.open(p, encoding='utf-8').read()

old_sig = '''Private Function FillResourceDbForSite(ByVal ProjectID As String, _
    ByVal CentroidLatitude As Double, ByVal CentroidLongitude As Double) As String'''
new_sig = '''Private Function FillResourceDbForSite(ByVal ProjectID As String, _
    ByVal CentroidLatitude As Double, ByVal CentroidLongitude As Double, _
    Optional ByVal AllowCreate As Boolean = False) As String'''
assert old_sig in s
s = s.replace(old_sig, new_sig, 1)

old_norow = '''    If Target Is Nothing Then
        FillResourceDbForSite = "NOROW"
        Exit Function
    End If

    DidText = ""'''
new_norow = '''    'v2.5: row hai hi nahi to bana do - draw kiya hai to location RESOURCE_DB
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
    End If'''
assert old_norow in s
s = s.replace(old_norow, new_norow, 1)

old_call = '''                                ResultText = FillResourceDbForSite(ProjectID, CentroidLatitude, CentroidLongitude)
                                If InStr(1, ResultText, "LOCPEND", vbBinaryCompare) > 0 Then
                                    'lat/lon bhar gaye, label offline tha - limited retries.
                                    If Attempts < AUTO_LABEL_RETRIES Then _
                                        mProcessed(KeyText) = "LOCPEND:" & CStr(Attempts + 1)
                                ElseIf ResultText = "NOROW" Then
                                    'RESOURCE_DB row abhi bani nahi - kuch ticks retry karo.
                                    If Attempts < AUTO_NOROW_RETRIES Then _
                                        mProcessed(KeyText) = "NOROW:" & CStr(Attempts + 1)'''
new_call = '''                                ResultText = FillResourceDbForSite(ProjectID, CentroidLatitude, _
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
                                        mProcessed(KeyText) = "NOROW:" & CStr(Attempts + 1)'''
assert old_call in s
s = s.replace(old_call, new_call, 1)

io.open(p, 'w', encoding='utf-8').write(s)
print("watcher block -> v2.5 OK")
