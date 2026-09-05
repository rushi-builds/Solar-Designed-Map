import io

ROOT = '/home/user/Solar-Designed-Map/location-extract/'
s = io.open(ROOT + 'modSolarEPCResource.bas', encoding='utf-8').read()

# 1) GeoNames tier function (appended at EOF)
geonames_fn = '''
'--------------------------------------------------------------------------
' v3.1: GeoNames FREE tier - koi card nahi, koi billing nahi, koi download
' nahi. Sirf ek chhota HTTPS call (lat/lon) -> ~1KB JSON wapas. GeoNames ka
' database unke server pe hai; workbook me kuch store NAHIN hota.
' Username free hai (geonames.org signup, email only): SETTINGS!B12 me rakho,
' fallback _CLOUD_CFG!B19. Username na ho to tier chupchaap skip ho jaata hai.
' Sirf populated places (gaon/qasbe) accept hote hain, aur sirf 3 km ke andar.
'--------------------------------------------------------------------------
Private Function ResourceGeoNamesLabel(ByVal Lat As Double, ByVal Lon As Double) As String
    Dim UserName As String, UrlText As String, JSONText As String
    Dim Village As String, Taluka As String, District As String
    Dim PlaceLat As String, PlaceLon As String
    Dim Km As Double

    On Error GoTo Failed
    UserName = ResourceSettingCell(SETTINGS_SHEET, "B12")
    If Len(UserName) = 0 Then UserName = ResourceConfigCell("B19")
    If Len(UserName) = 0 Then
        ResourceSettingDiag "A20", "GEONAMES STATUS", "B20", _
            "NO_USER - geonames.org pe free signup karke username SETTINGS!B12 me rakho"
        Exit Function
    End If

    UrlText = "https://secure.geonames.org/findNearbyPlaceNameJSON?lat=" & _
              ResourceDecimal(Lat, 6) & "&lng=" & ResourceDecimal(Lon, 6) & _
              "&radius=3&maxRows=1&style=LONG&username=" & UserName
    JSONText = ResourceHttpGetJson(UrlText)
    If Len(JSONText) = 0 Then
        ResourceSettingDiag "A20", "GEONAMES STATUS", "B20", "HTTP_FAIL - network block"
        Exit Function
    End If
    If InStr(1, JSONText, """" & "status" & """", vbBinaryCompare) > 0 Then
        ResourceSettingDiag "A20", "GEONAMES STATUS", "B20", _
            "ERROR - " & Left$(ResourceJSONValue(JSONText, "message"), 200)
        Exit Function
    End If

    Village = Trim$(ResourceJSONValue(JSONText, "name"))
    If Len(Village) = 0 Then
        ResourceSettingDiag "A20", "GEONAMES STATUS", "B20", "NO_MATCH - 3 km me koi gaon nahi"
        Exit Function
    End If
    PlaceLat = Trim$(ResourceJSONValue(JSONText, "lat"))
    PlaceLon = Trim$(ResourceJSONValue(JSONText, "lng"))
    If IsNumeric(PlaceLat) And IsNumeric(PlaceLon) Then
        Km = ResourceDistanceKm(Lat, Lon, CDbl(PlaceLat), CDbl(PlaceLon))
        If Km > 3# Then
            ResourceSettingDiag "A20", "GEONAMES STATUS", "B20", _
                "TOO_FAR - " & Village & " " & Format$(Km, "0.0") & " km door tha, reject"
            Exit Function
        End If
    End If

    Taluka = ResourceAdminClean(ResourceJSONValue(JSONText, "adminName3"))
    District = ResourceAdminClean(ResourceJSONValue(JSONText, "adminName2"))
    If ResourceSameName(Taluka, Village) Then Taluka = vbNullString
    If ResourceSameName(District, Village) Then District = vbNullString
    If ResourceSameName(District, Taluka) Then District = vbNullString
    ResourceGeoNamesLabel = ResourceLocationCompose(Village, _
        ResourceLocationCompose(Taluka, District))
    ResourceSettingDiag "A20", "GEONAMES STATUS", "B20", "OK - " & ResourceGeoNamesLabel
    Exit Function
Failed:
    ResourceGeoNamesLabel = vbNullString
End Function
'''
s = s.rstrip() + '\n' + geonames_fn

# 2) tier call: VILLAGE_DB ke turant baad
anchor = """    ResourceLocationLabel = ResourceVillageDbLabel(Lat, Lon)
    If Len(ResourceLocationLabel) > 0 Then GoTo StoreCache
"""
assert s.count(anchor) == 1
s = s.replace(anchor, anchor + """
    '0a) v3.1: GeoNames free tier (no card/no billing/no download) - Indian
    '   census village names jo OSM/BigDataCloud ke paas nahi hote.
    ResourceLocationLabel = ResourceGeoNamesLabel(Lat, Lon)
    If Len(ResourceLocationLabel) > 0 Then GoTo StoreCache
""", 1)

# 3) debug macro me GeoNames status line
old_dbg = '            LastFill = CStr(ws.Range("B18").Value2)\n'
assert s.count(old_dbg) == 1
s = s.replace(old_dbg, old_dbg + '            GeoNamesStatus = CStr(ws.Range("B20").Value2)\n', 1)
old_dim = '    Dim GoogleStatus As String, GoogleDetail As String, VillageMatch As String, LastFill As String\n'
assert s.count(old_dim) == 1
s = s.replace(old_dim, old_dim + '    Dim GeoNamesStatus As String\n', 1)
old_p2 = '                      "   Last auto-fill: " & IIf(Len(LastFill) > 0, LastFill, "-") & vbCrLf\n'
assert s.count(old_p2) == 1
s = s.replace(old_p2, '                      "   GeoNames tier: " & IIf(Len(GeoNamesStatus) > 0, GeoNamesStatus, "-") & vbCrLf & _\n' + old_p2, 1)

# 4) version bump
assert "' Version 3.0" in s
s = s.replace("' Version 3.0", "' Version 3.1", 1)
hist = "'   - v3.0: SINGLE MODULE"
assert hist in s
s = s.replace(hist, "'   - v3.1: GeoNames FREE tier (koi card/billing/download nahi) - Indian\n"
    "'     census village names; username SETTINGS!B12 (ya _CLOUD_CFG!B19).\n"
    + hist, 1)

io.open(ROOT + 'modSolarEPCResource.bas', 'w', encoding='utf-8').write(s)
print('v3.1 assembled:', s.count('\n') + 1, 'lines')
