import io, re

ROOT = '/home/user/Solar-Designed-Map/location-extract/'
s = io.open(ROOT + 'modSolarEPCResource.bas', encoding='utf-8').read()

# 1) GeoNames tier call remove
call = """    '0a) v3.1: GeoNames free tier (no card/no billing/no download) - Indian
    '   census village names jo OSM/BigDataCloud ke paas nahi hote.
    ResourceLocationLabel = ResourceGeoNamesLabel(Lat, Lon)
    If Len(ResourceLocationLabel) > 0 Then GoTo StoreCache
"""
assert s.count(call) == 1
s = s.replace(call, '', 1)

# 2) GeoNames functions remove (label + http helper), comment-header se End Function tak
for marker in ["' v3.1: GeoNames FREE tier", "' v3.3: GeoNames ko 3 endpoints"]:
    i = s.index(marker)
    start = s.rindex("'-----", 0, i)
    end = s.index('\nEnd Function\n', i) + len('\nEnd Function\n')
    s = s[:start] + s[end:]

# 3) debug macro se GeoNames lines remove
for line in ['    Dim GeoNamesStatus As String\n',
             '            GeoNamesStatus = CStr(ws.Range("B20").Value2)\n',
             '                      "   GeoNames tier: " & IIf(Len(GeoNamesStatus) > 0, GeoNamesStatus, "-") & vbCrLf & _\n']:
    assert s.count(line) == 1, line
    s = s.replace(line, '', 1)

# 4) Date/Time column: user ka apna column preferred, real date value
old_stamp = """    On Error Resume Next
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
"""
new_stamp = """    On Error Resume Next
    C = ResourceColumn(Tbl, "Date/Time")              'user ka apna column
    If C = 0 Then C = ResourceColumn(Tbl, "Run Date-Time (Local)")
    If C = 0 Then
        Set Col = Tbl.ListColumns.Add(Tbl.ListColumns.Count + 1)
        If Not Col Is Nothing Then Col.Name = "Run Date-Time (Local)"
        C = ResourceColumn(Tbl, "Run Date-Time (Local)")
    End If
    On Error GoTo 0
    If C = 0 Then Exit Sub
    Set Cell = Target.Cells(1, C)
    If Cell Is Nothing Then Exit Sub
    If Cell.HasFormula Then Exit Sub
    Cell.NumberFormat = "dd-mm-yyyy hh:nn:ss"
    Cell.Value2 = Now                                  'real date value, sortable
"""
assert s.count(old_stamp) == 1
s = s.replace(old_stamp, new_stamp, 1)
s = s.replace("""    Dim ColName As String
    Dim C As Long
    Dim Col As ListColumn
    Dim Cell As Range
""", """    Dim C As Long
    Dim Col As ListColumn
    Dim Cell As Range
""", 1)

# 4b) purane GeoNames header bullets remove (tier hat chuka hai)
for bullet in ["""'   - v3.1: GeoNames FREE tier (koi card/billing/download nahi) - Indian
'     census village names; username SETTINGS!B12 (ya _CLOUD_CFG!B19).
""", """'   - v3.3: GeoNames tier ab 3 endpoints try karta hai (secure https,
'     api https, api http) + MSXML fallback; B20 me ASLI WinHttp error
'     number aata hai (12007 DNS / 12029 connect / 12002 timeout / TLS).
"""]:
    assert s.count(bullet) == 1, bullet
    s = s.replace(bullet, '', 1)

# 5) version -> 3.4 LEAN
assert "' Version 3.3" in s
s = s.replace("' Version 3.3", "' Version 3.4", 1)
hist = "'   - v3.2: SolarEPC_ResourceRefreshLabels"
assert hist in s
s = s.replace(hist, "'   - v3.4 LEAN: GeoNames tier HATAYA (unke database me Indian\n"
    "'     revenue villages hain hi nahi - prove ho gaya). Ab user ka apna\n"
    "'     \"Date/Time\" column bharta hai (local Now, real date value);\n"
    "'     column na ho to \"Run Date-Time (Local)\" apne aap banti hai.\n"
    "'     NASA fetch hamesha us row ke APNE centroid se hota hai - same\n"
    "'     town/village ho tab bhi coordinates se alag result aata hai.\n"
    + hist, 1)

io.open(ROOT + 'modSolarEPCResource.bas', 'w', encoding='utf-8').write(s)
print('v3.4 LEAN assembled:', s.count('\n') + 1, 'lines | GeoNames refs:', s.count('GeoNames'))
