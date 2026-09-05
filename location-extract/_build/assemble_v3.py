import io

ROOT = '/home/user/Solar-Designed-Map/location-extract/'
s = io.open(ROOT + 'modSolarEPCResource.bas', encoding='utf-8').read()
v3 = io.open(ROOT + '_build/v3_block.bas', encoding='utf-8').read()

# 1) call the local run-time stamp at the end of ResourceWriteToDatabase
anchor = '    ResourceWriteTableField Tbl, Target, "Remarks", ResourceValue(D, "remarks")\n'
assert s.count(anchor) == 1
s = s.replace(anchor, anchor + '    \'v3.0: exact local run stamp (new column, auto-created once).\n    ResourceStampLocalRunTime Tbl, Target\n', 1)

# 2) append the v3 helper block at EOF
s = s.rstrip() + '\n\n' + v3.rstrip() + '\n'

# 3) header -> v3.0
assert "' Version 2.5" in s
s = s.replace("' Version 2.5", "' Version 3.0", 1)
hist = "'   - PLUS v2.3 ka AUTOMATIC watcher:"
assert hist in s
s = s.replace(hist,
    "'   - v3.0: SINGLE MODULE - NASA pipeline + watcher + auto row-create +\n"
    "'     debug, sab ek hi file me. Nayi RESOURCE_DB column\n"
    "'     \"Run Date-Time (Local)\" har NASA summary import pe aapke computer\n"
    "'     ka exact din + samay stamp karti hai (column khud ban jaati hai).\n"
    + hist, 1)

io.open(ROOT + 'modSolarEPCResource.bas', 'w', encoding='utf-8').write(s)
print('v3.0 assembled:', s.count('\n') + 1, 'lines')
