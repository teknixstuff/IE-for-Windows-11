Option Explicit
'~ On Error Resume Next
'RequireAdmin

Dim objReg
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "", "REG_SZ", ""
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "CurrentLevel", "REG_DWORD", 65536
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "Flags", "REG_DWORD", 67
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1200", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1400", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2001", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2004", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1001", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1004", "REG_DWORD", 1
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1201", "REG_DWORD", 1
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1206", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1207", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1208", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1209", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "120A", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "120B", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "120C", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1402", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1405", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1406", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1407", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1408", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1409", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "140A", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "140C", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "140D", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1601", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1604", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1605", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1606", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1607", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1608", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1609", "REG_DWORD", 1
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "160A", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "160B", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1802", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1803", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1804", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1809", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1812", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1A00", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1A02", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1A03", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1A04", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1A05", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1A06", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1A10", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "1C00", "REG_DWORD", 196608
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2000", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2005", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2007", "REG_DWORD", 65536
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2100", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2101", "REG_DWORD", 1
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2102", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2103", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2104", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2105", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2106", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2107", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2108", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2200", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2201", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2300", "REG_DWORD", 1
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2301", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2302", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2400", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2401", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2402", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2600", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2700", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2701", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2702", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2703", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2704", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2708", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2709", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "270B", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "270C", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "270D", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2500", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3", "2707", "REG_DWORD", 0





RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "", "REG_SZ", ""
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "CurrentLevel", "REG_DWORD", 65536
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "Flags", "REG_DWORD", 67
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1200", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1400", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2001", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2004", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1001", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1004", "REG_DWORD", 1
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1201", "REG_DWORD", 1
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1206", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1207", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1208", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1209", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "120A", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "120B", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "120C", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1402", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1405", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1406", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1407", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1408", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1409", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "140A", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "140C", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "140D", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1601", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1604", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1605", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1606", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1607", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1608", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1609", "REG_DWORD", 1
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "160A", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "160B", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1802", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1803", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1804", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1809", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1812", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1A00", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1A02", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1A03", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1A04", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1A05", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1A06", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1A10", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "1C00", "REG_DWORD", 196608
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2000", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2005", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2007", "REG_DWORD", 65536
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2100", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2101", "REG_DWORD", 1
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2102", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2103", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2104", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2105", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2106", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2107", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2108", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2200", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2201", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2300", "REG_DWORD", 1
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2301", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2302", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2400", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2401", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2402", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2600", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2700", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2701", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2702", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2703", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2704", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2708", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2709", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "270B", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "270C", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "270D", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2500", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2", "2707", "REG_DWORD", 0





RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "", "REG_SZ", ""
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "CurrentLevel", "REG_DWORD", 65536
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "Flags", "REG_DWORD", 67
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1200", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1400", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2001", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2004", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1001", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1004", "REG_DWORD", 1
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1201", "REG_DWORD", 1
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1206", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1207", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1208", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1209", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "120A", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "120B", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "120C", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1402", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1405", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1406", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1407", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1408", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1409", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "140A", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "140C", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "140D", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1601", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1604", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1605", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1606", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1607", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1608", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1609", "REG_DWORD", 1
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "160A", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "160B", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1802", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1803", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1804", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1809", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1812", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1A00", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1A02", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1A03", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1A04", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1A05", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1A06", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1A10", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "1C00", "REG_DWORD", 196608
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2000", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2005", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2007", "REG_DWORD", 65536
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2100", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2101", "REG_DWORD", 1
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2102", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2103", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2104", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2105", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2106", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2107", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2108", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2200", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2201", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2300", "REG_DWORD", 1
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2301", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2302", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2400", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2401", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2402", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2600", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2700", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2701", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2702", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2703", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2704", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2708", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2709", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "270B", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "270C", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "270D", "REG_DWORD", 0
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2500", "REG_DWORD", 3
RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1", "2707", "REG_DWORD", 0

CreateObject("InternetExplorer.Application").Visible=true

Function RegWrite(reg_keyname, reg_valuename,reg_type,ByVal reg_value)
	Dim aRegKey, Return
	aRegKey = RegSplitKey(reg_keyname)
	If IsArray(aRegKey) = 0 Then
		RegWrite = 0
		Exit Function
	End If

	Return = RegWriteKey(aRegKey)
	If Return = 0 Then
		RegWrite = 0
		Exit Function
	End If

	Select Case reg_type
		Case "REG_SZ"
			Return = objReg.SetStringValue(aRegKey(0),aRegKey(1),reg_valuename,reg_value)
		Case "REG_EXPAND_SZ"
			Return = objReg.SetExpandedStringValue(aRegKey(0),aRegKey(1),reg_valuename,reg_value)
		Case "REG_BINARY"
			If IsArray(reg_value) = 0 Then reg_value = Array()
			Return = objReg.SetBinaryValue(aRegKey(0),aRegKey(1),reg_valuename,reg_value)

		Case "REG_DWORD"
			If IsNumeric(reg_value) = 0 Then reg_value = 0
			Return = objReg.SetDWORDValue(aRegKey(0),aRegKey(1),reg_valuename,reg_value)

		Case "REG_MULTI_SZ"
			If IsArray(reg_value) = 0 Then
				If Len(reg_value) = 0 Then
					reg_value = Array()
				Else
					reg_value = Array(reg_value)
				End If
			End If
			Return = objReg.SetMultiStringValue(aRegKey(0),aRegKey(1),reg_valuename,reg_value)

		'Case "REG_QWORD"
			'Return = oReg.SetQWORDValue(aRegKey(0),aRegKey(1),reg_valuename,reg_value)
		Case Else
			RegWrite = 0
			Exit Function
	End Select

	If (Return <> 0) Or (Err.Number <> 0) Then
		RegWrite = 0
		Exit Function
	End If
	RegWrite = 1
End Function

Function RegWriteKey(RegKeyName)
	Dim Return
	If IsArray(RegKeyName) = 0 Then
		RegKeyName = RegSplitKey(RegKeyName)
	End If

	If (IsArray(RegKeyName) = 0) Or (UBound(RegKeyName) <> 1) Then
		RegWriteKey = 0
		Exit Function
	End If

	Return = objReg.CreateKey(RegKeyName(0),RegKeyName(1))
	If (Return <> 0) Or (Err.Number <> 0) Then
		RegWriteKey = 0
		Exit Function
	End If
	RegWriteKey = 1
End Function

Function RegDelete(reg_keyname, reg_valuename)
	Dim Return,aRegKey
	aRegKey = RegSplitKey(reg_keyname)
	If IsArray(aRegKey) = 0 Then
		RegDelete = 0
		Exit Function
	End If

	Return = objReg.DeleteValue(aRegKey(0),aRegKey(1),reg_valuename)
	If (Return <> 0) And (Err.Number <> 0) Then
		RegDelete = 0
		Exit Function
	End If
	RegDelete = 1
End Function

Function RegDeleteKey(reg_keyname)
	Dim Return,aRegKey
	aRegKey = RegSplitKey(reg_keyname)
	If IsArray(aRegKey) = 0 Then
		RegDeleteKey = 0
		Exit Function
	End If

	'On Error Resume Next
	Return = RegDeleteSubKey(aRegKey(0),aRegKey(1))
	'On Error Goto 0
	If Return = 0 Then
		RegDeleteKey = 0
		Exit Function
	End If
	RegDeleteKey = 1
End Function

Function RegDeleteSubKey(strRegHive, strKeyPath)
	Dim Return,arrSubkeys,strSubkey
    objReg.EnumKey strRegHive, strKeyPath, arrSubkeys
    If IsArray(arrSubkeys) <> 0 Then
        For Each strSubkey In arrSubkeys
            RegDeleteSubKey strRegHive, strKeyPath & "\" & strSubkey
        Next
    End If

	Return = objReg.DeleteKey(strRegHive, strKeyPath)
	If (Return <> 0) Or (Err.Number <> 0) Then
		RegDeleteSubKey = 0
		Exit Function
	End If
	RegDeleteSubKey = 1
End Function

Function RegSplitKey(RegKeyName)
	Dim strHive, strInstr, strLeft
	strInstr=InStr(RegKeyName,"\")
	If strInstr = 0 Then Exit Function
	strLeft=left(RegKeyName,strInstr-1)

	Select Case strLeft
		Case "HKCR","HKEY_CLASSES_ROOT"	strHive = &H80000000
		Case "HKCU","HKEY_CURRENT_USER"	strHive = &H80000001
		Case "HKLM","HKEY_LOCAL_MACHINE"	strHive = &H80000002
		Case "HKU","HKEY_USERS" 	strHive = &H80000003
		Case "HKCC","HKEY_CURRENT_CONFIG"	strHive = &H80000005
	  Case Else Exit Function
	End Select

    RegSplitKey = Array(strHive,Mid(RegKeyName,strInstr+1))
End Function

Function RequireAdmin()
	Dim reg_valuename, WShell, Cmd, CmdLine, I

	GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")_
	.EnumValues &H80000003, "S-1-5-19\Environment",  reg_valuename
	If IsArray(reg_valuename) <> 0 Then
		RequireAdmin = 1
		Exit Function
	End If

	Set Cmd = WScript.Arguments
	For I = 0 to Cmd.Count - 1
		If Cmd(I) = "/admin" Then
			Wscript.Echo "To script you must have administrator rights!"
			'RequireAdmin = 0
			'Exit Function
			WScript.Quit
		End If
		CmdLine = CmdLine & Chr(32) & Chr(34) & Cmd(I) & Chr(34)
	Next
	CmdLine = CmdLine & Chr(32) & Chr(34) & "/admin" & Chr(34)

	Set WShell= WScript.CreateObject( "WScript.Shell")
	CreateObject("Shell.Application").ShellExecute WShell.ExpandEnvironmentStrings(_
	"%SystemRoot%\System32\WScript.exe"),Chr(34) & WScript.ScriptFullName & Chr(34) & CmdLine, "", "runas"
	WScript.Quit
End Function