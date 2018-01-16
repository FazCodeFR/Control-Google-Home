on error resume next

Set objArgs = WScript.Arguments
For I = 0 to objArgs.Count -1
a = a & " " & objArgs(I)	
Next

'inputbox "Valeur passer","valeur :",a

Set WshShell = CreateObject("WScript.Shell")

If a = " augmenter le son" then 
WshShell.SendKeys "{" & chr(175) & " 10}"
Elseif a =  " augmenter le volume" then 
WshShell.SendKeys "{" & chr(175) & " 10}"
Elseif a = " monte le sons" then 
WshShell.SendKeys "{" & chr(175) & " 10}"
Elseif a = " monte le sons au max" then 
WshShell.SendKeys "{" & chr(175) & " 50}"
Elseif a = " monte le sons au maximum" then 
WshShell.SendKeys "{" & chr(175) & " 50}"
Elseif a = " monte le volume au max" then 
WshShell.SendKeys "{" & chr(175) & " 50}"
Elseif a = " monte le volume au maximum" then 
WshShell.SendKeys "{" & chr(175) & " 50}"
Elseif a = " volume max" then 
WshShell.SendKeys "{" & chr(175) & " 50}"
Elseif a = " son au max" then 
WshShell.SendKeys "{" & chr(175) & " 50}"
Elseif a = " augmente le son au maximum" then 
WshShell.SendKeys "{" & chr(175) & " 50}"
Elseif a = " augmente le son" then 
WshShell.SendKeys "{" & chr(175) & " 10}"

'*Décodeur Orange* 

'Configuration requise : https://blubsy-news.blogspot.fr/2017/10/domotique-commander-la-livebox-avec-la.html
Elseif a = " allume le décodeur" then
Call onoffdecodeur ()
Elseif a = " allumer le décodeur" then
Call onoffdecodeur ()
Elseif a = " démarre le décodeur" then
Call onoffdecodeur ()
Elseif a = " allume le décodeur tv" then
Call onoffdecodeur ()
Elseif a = " allume la livebox" then
Call onoffdecodeur ()
Elseif a = " lance le décodeur" then
Call onoffdecodeur ()
Elseif a = " éteint le décodeur" then
Call onoffdecodeur ()
Elseif a = " éteint la livebox" then
Call onoffdecodeur ()
Elseif a = " mute le decodeur" then
Call mutedecodeur ()
Elseif a = " met en muet le décodeur" then
Call mutedecodeur ()
Elseif a = " coupe le son du décodeur" then
Call mutedecodeur ()
Elseif a = " remet le son du décodeur" then
Call mutedecodeur ()
Elseif a = " augmente le son du décodeur" then
Call VolumeUpdecodeur ()
Elseif a = " augmente le volume du décodeur" then
Call VolumeUpdecodeur ()
Elseif a = " met le son du décodeur plus fort" then
Call VolumeUpdecodeur ()
Elseif a = " monte le son du décodeur" then
Call VolumeUpdecodeur ()
Elseif a = " monte le volume du décodeur" then
Call VolumeUpdecodeur ()



Elseif a = " baisse le volume du décodeur" then
Call VolumeDowndecodeur ()
Elseif a = " met le son du décodeur moins fort" then
Call VolumeDowndecodeur ()
Elseif a = " baisse le son du décodeur" then
Call VolumeDowndecodeur ()
Elseif a = " baisse le volume du décodeur" then
Call VolumeDowndecodeur ()

End if 




sub onoffdecodeur ()
Dim IE
Set IE = Wscript.CreateObject("InternetExplorer.Application")
IE.Visible = 0
IE.navigate "http://192.168.1.10:8080/remoteControl/cmd?operation=01&key=116&mode=0"
WScript.Sleep 3000
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'iexplore.exe'")
For Each objProcess in colProcessList
    objProcess.Terminate()
Next
End sub

sub mutedecodeur ()
Dim IE
Set IE = Wscript.CreateObject("InternetExplorer.Application")
IE.Visible = 0
IE.navigate "http://192.168.1.10:8080/remoteControl/cmd?operation=01&key=113&mode=0"
WScript.Sleep 3000
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'iexplore.exe'")
For Each objProcess in colProcessList
    objProcess.Terminate()
Next
End sub

sub VolumeUpdecodeur ()
Dim IE
Set IE = Wscript.CreateObject("InternetExplorer.Application")
IE.Visible = 0
IE.navigate "http://192.168.1.10:8080/remoteControl/cmd?operation=01&key=115&mode=0"
WScript.Sleep 3000
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'iexplore.exe'")
For Each objProcess in colProcessList
    objProcess.Terminate()
Next
End sub


sub VolumeDowndecodeur ()
Dim IE
Set IE = Wscript.CreateObject("InternetExplorer.Application")
IE.Visible = 0
IE.navigate "http://192.168.1.10:8080/remoteControl/cmd?operation=01&key=114&mode=0"
WScript.Sleep 3000
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'iexplore.exe'")
For Each objProcess in colProcessList
    objProcess.Terminate()
Next
End sub



'inputbox "Valeur passer","valeur :",a
