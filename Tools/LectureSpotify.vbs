On Error Resume Next
SpotifyActif = False
Set WS = WScript.CreateObject("WScript.Shell") 
Comandline = ws.ExpandEnvironmentStrings("%HOMEPATH%") & "\AppData\Roaming\Spotify\Spotify.exe"
Set fso = CreateObject("Scripting.FileSystemObject")
QueryLine = "select * from win32_process where Name = 'Spotify.exe'"
Computer = "."
Set WMIService = GetObject("winmgmts:\\" & Computer & "\root\cimv2")
Set Items = WMIService.ExecQuery(QueryLine, , 48)
For Each SubItems In Items
SpotifyActif = True
Next
If SpotifyActif = True then 'actif
strComputer = "." 
Set objWMIService = GetObject("winmgmts:" _ 
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
Set colProcessList = objWMIService.ExecQuery _ 
    ("Select * from Win32_Process Where Name = 'Spotify.exe'") 
For Each objProcess in colProcessList 
    objProcess.Terminate() 
Next
End if
  If fso.FileExists(Comandline) = true then
   WS.Run Comandline,3
            Else
   WS.Run "spotify:"
  End if 
  WScript.Sleep 9000
  WS.AppActivate Comandline
  WS.SendKeys " "
  WScript.Sleep 900
  WS.SendKeys "%{TAB}"
