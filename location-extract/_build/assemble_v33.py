import io

ROOT = '/home/user/Solar-Designed-Map/location-extract/'
s = io.open(ROOT + 'modSolarEPCResource.bas', encoding='utf-8').read()

old = '''    UrlText = "https://secure.geonames.org/findNearbyPlaceNameJSON?lat=" & _
              ResourceDecimal(Lat, 6) & "&lng=" & ResourceDecimal(Lon, 6) & _
              "&radius=3&maxRows=1&style=LONG&username=" & UserName
    JSONText = ResourceHttpGetJson(UrlText)
    If Len(JSONText) = 0 Then
        ResourceSettingDiag "A20", "GEONAMES STATUS", "B20", "HTTP_FAIL - network block"
        Exit Function
    End If
'''
new = '''    QueryText = "findNearbyPlaceNameJSON?lat=" & _
              ResourceDecimal(Lat, 6) & "&lng=" & ResourceDecimal(Lon, 6) & _
              "&radius=3&maxRows=1&style=LONG&username=" & UserName
    JSONText = ResourceGeoNamesHttp(QueryText, DiagText)
    If Len(JSONText) = 0 Then
        ResourceSettingDiag "A20", "GEONAMES STATUS", "B20", "HTTP_FAIL - " & DiagText
        Exit Function
    End If
'''
assert s.count(old) == 1
s = s.replace(old, new, 1)

old_dim = '''    Dim UserName As String, UrlText As String, JSONText As String
    Dim Village As String, Taluka As String, District As String
    Dim PlaceLat As String, PlaceLon As String
    Dim Km As Double
'''
new_dim = '''    Dim UserName As String, QueryText As String, JSONText As String
    Dim Village As String, Taluka As String, District As String
    Dim PlaceLat As String, PlaceLon As String
    Dim DiagText As String
    Dim Km As Double
'''
assert s.count(old_dim) == 1
s = s.replace(old_dim, new_dim, 1)

helper = '''
'--------------------------------------------------------------------------
' v3.3: GeoNames ko 3 endpoints pe try karo (secure https -> api https ->
' api http) aur ASLI error number/message report karo, taaki "network block"
' jaisa andha message dobara na aaye. 12007=DNS fail, 12029=connect fail,
' 12002=timeout, 12044/12030=certificate - B20 me sab likha aayega.
'--------------------------------------------------------------------------
Private Function ResourceGeoNamesHttp(ByVal QueryText As String, _
    ByRef DiagText As String) As String

    Dim Hosts As Variant
    Dim i As Long
    Dim HTTP As Object
    Dim LastError As String

    Hosts = Array("https://secure.geonames.org/", _
                  "https://api.geonames.org/", _
                  "http://api.geonames.org/")
    For i = LBound(Hosts) To UBound(Hosts)
        On Error GoTo NextHost
        Set HTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
        HTTP.Open "GET", CStr(Hosts(i)) & QueryText, False
        HTTP.setTimeouts 6000, 6000, 15000, 15000
        HTTP.setRequestHeader "User-Agent", "Solar-EPC-Resource/3.3 (Excel location label)"
        HTTP.setRequestHeader "Accept", "application/json"
        HTTP.send
        If HTTP.Status = 200 Then
            ResourceGeoNamesHttp = HTTP.ResponseText
            Exit Function
        End If
        LastError = CStr(Hosts(i)) & " HTTP " & CStr(HTTP.Status)
NextHost:
        If Err.Number <> 0 Then
            LastError = CStr(Hosts(i)) & " err " & CStr(Err.Number) & " " & Left$(Err.Description, 120)
            Err.Clear
        End If
    Next i

    'Aakhri koshish: MSXML2 (kabhi-kabhi proxy settings alag hoti hain).
    On Error GoTo FailedMsxml
    Set HTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    HTTP.Open "GET", "https://secure.geonames.org/" & QueryText, False
    HTTP.setTimeouts 6000, 6000, 15000, 15000
    HTTP.setRequestHeader "User-Agent", "Solar-EPC-Resource/3.3 (Excel location label)"
    HTTP.send
    If HTTP.Status = 200 Then
        ResourceGeoNamesHttp = HTTP.ResponseText
        Exit Function
    End If
    LastError = LastError & " | msxml HTTP " & CStr(HTTP.Status)
FailedMsxml:
    If Err.Number <> 0 Then
        LastError = LastError & " | msxml err " & CStr(Err.Number) & " " & Left$(Err.Description, 120)
        Err.Clear
    End If
    DiagText = LastError
    ResourceGeoNamesHttp = vbNullString
End Function
'''
s = s.rstrip() + '\n' + helper

assert "' Version 3.2" in s
s = s.replace("' Version 3.2", "' Version 3.3", 1)
hist = "'   - v3.2: SolarEPC_ResourceRefreshLabels"
assert hist in s
s = s.replace(hist, "'   - v3.3: GeoNames tier ab 3 endpoints try karta hai (secure https,\n"
    "'     api https, api http) + MSXML fallback; B20 me ASLI WinHttp error\n"
    "'     number aata hai (12007 DNS / 12029 connect / 12002 timeout / TLS).\n"
    + hist, 1)

io.open(ROOT + 'modSolarEPCResource.bas', 'w', encoding='utf-8').write(s)
print('v3.3 assembled:', s.count('\n') + 1, 'lines')
