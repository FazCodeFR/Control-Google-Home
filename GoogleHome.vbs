on error resume next

Set objArgs = WScript.Arguments
For I = 0 to objArgs.Count -1
a = a & " " & objArgs(I)	
Next

Set WshShell = CreateObject("WScript.Shell")

If a = " augmente le son" then 
WshShell.SendKeys "{" & chr(175) & " 10}"
Elseif a =  " augmente le volume" then 
WshShell.SendKeys "{" & chr(175) & " 10}"
Elseif a = " monte le son" then WshShell.SendKeys "{" & chr(175) & " 10}"
Elseif a = " monte le son au max" then WshShell.SendKeys "{" & chr(175) & " 50}"
Elseif a = " monte le son au maximum" then WshShell.SendKeys "{" & chr(175) & " 50}"
Elseif a = " monte le volume au max" then WshShell.SendKeys "{" & chr(175) & " 50}"
Elseif a = " monte le volume au maximum" then WshShell.SendKeys "{" & chr(175) & " 50}"
Elseif a = " volume max" then WshShell.SendKeys "{" & chr(175) & " 50}"
Elseif a = " son au max" then WshShell.SendKeys "{" & chr(175) & " 50}"
Elseif a = " augmente le son au maximum" then WshShell.SendKeys "{" & chr(175) & " 50}"
Elseif a = " baisse le son" then WshShell.SendKeys "{" & chr(174) & " 10}"
Elseif a = " descend le son" then WshShell.SendKeys "{" & chr(174) & " 10}"
Elseif a = " descend le volume" then WshShell.SendKeys "{" & chr(174) & " 10}"
Elseif a = " descend le son au max" then WshShell.SendKeys "{" & chr(174) & " 50}"
Elseif a = " baisse le son au max" then WshShell.SendKeys "{" & chr(174) & " 50}"
Elseif a = " baisse le volume" then WshShell.SendKeys "{" & chr(174) & " 10}"
Elseif a = " baisse le volume au max" then WshShell.SendKeys "{" & chr(174) & " 50}"
Elseif a = " baisse le son au maximum" then WshShell.SendKeys "{" & chr(174) & " 50}"
Elseif a = " baisse le volume au maximum" then WshShell.SendKeys "{" & chr(174) & " 50}"
Elseif a = " mute le volume" then WshShell.SendKeys chr(173)
Elseif a = " mute le son" then WshShell.SendKeys chr(173)
Elseif a = " muet" then WshShell.SendKeys chr(173)
Elseif a = " le son a 0" then WshShell.SendKeys chr(173)
Elseif a = " coupe le son" then WshShell.SendKeys chr(173)
Elseif a = " coupe le volume" then WshShell.SendKeys chr(173)
Elseif a = " coupe l'audio" then WshShell.SendKeys chr(173)
Elseif a = " remet le volume" then WshShell.SendKeys chr(173)
Elseif a = " remet le son" then WshShell.SendKeys chr(173)

Elseif a = " lance chrome" then WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Elseif a = " ouvre chrome" then WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Elseif a = " démarre chrome" then WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Elseif a = " exécute chrome" then WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Elseif a = " lance google chrome" then WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Elseif a = " ouvre google chrome" then WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Elseif a = " démarre google chrome" then WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Elseif a = " exécute google chrome" then WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Elseif a = " lance google" then WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Elseif a = " ouvre google" then WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Elseif a = " démarre google" then WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Elseif a = " exécute chrome" then WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Elseif a = " lorsque Rome" then WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Elseif a = " fait pause" then WshShell.SendKeys " "
Elseif a = " met pause" then WshShell.SendKeys " "
Elseif a = " fais une pause" then WshShell.SendKeys " "
Elseif a = " met en pause" then WshShell.SendKeys " "
Elseif a = " mais en pause" then WshShell.SendKeys " "
Elseif a = " fait pause" then WshShell.SendKeys " "
Elseif a = " fait stop" then WshShell.SendKeys " "
Elseif a = " stop" then WshShell.SendKeys " "
Elseif a = " pause" then WshShell.SendKeys " "
Elseif a = " relance" then WshShell.SendKeys " "
Elseif a = " enlève la pause" then WshShell.SendKeys " "
Elseif a = " met une pause" then WshShell.SendKeys " "
Elseif a = " lance" then WshShell.SendKeys " "
Elseif a = " lecture" then WshShell.SendKeys " "
Elseif a = " lance lecture" then WshShell.SendKeys " "
Elseif a = " lance la lecture" then WshShell.SendKeys " "
Elseif a = " éteint le" then CreateObject("Wscript.Shell").Run "CMD /C " & " shutdown /s /f"
Elseif a = " arrête le" then CreateObject("Wscript.Shell").Run "CMD /C " & " shutdown /s /f"
Elseif a = " éteint le pc" then CreateObject("Wscript.Shell").Run "CMD /C " & " shutdown /s /f"
Elseif a = " éteint l'ordinateur" then CreateObject("Wscript.Shell").Run "CMD /C " & " shutdown /s /f"
Elseif a = " arrête le système" then CreateObject("Wscript.Shell").Run "CMD /C " & " shutdown /s /f"
Elseif a = " éteint le système" then CreateObject("Wscript.Shell").Run "CMD /C " & " shutdown /s /f"

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
