IpDecodeur = ""

If IpDecodeur = "" then 
MsgBox "Aller sur le lien suivant sur votre navigateur : http://livebox/" & vbcr & vbcr & "Dans 'équipements connectés à la Livebox' trouver votre décodeur orange et copier son ip" & vbcr & vbcr & "Ecrire l'ip copier a la ligne 1 dans les guillemets du fichier DecodeurOrange.vbs", vbInformation+vbOKOnly, "Configurer le script"
Wscript.Quit ()
End if 

' Ok google, sur le décodeur XXX
' Fichier a mettre dans : C:\GoogleHome\
' Applet IFTTT : https://ifttt.com/applets/N8a5s4BE-controle-du-decodeur-orange-avec-google-assitant
' Aller sur le lien suivant sur votre navigateur : http://livebox/
' Dans 'équipements connectés à la Livebox' trouver votre décodeur orange et copier son ip
' Ecrire l'ip copier a la ligne 1 dans les "" du fichier DecodeurOrange.vbs

Set objArgs = WScript.Arguments
For I = 0 to objArgs.Count -1

'MsgBox objArgs(I)
If IsNumeric(objArgs(I)) Then 
chaine = objArgs(I) 
call chainevalue(chaine)
Exit for

End if 


a = a & " " & objArgs(I)	
Next
Set WshShell = CreateObject("WScript.Shell")

Select Case a
Case " allumer le"
Call onoffdecodeur ()
Case " allume"
Call onoffdecodeur ()
Case " allume le"
Call onoffdecodeur ()
Case " lance"
Call onoffdecodeur ()
Case " lance le"
Call onoffdecodeur ()
Case " démarre"
Call onoffdecodeur ()
Case " démarre le"
Call onoffdecodeur ()
Case " éteint"
Call onoffdecodeur ()
Case " éteint le"
Call onoffdecodeur ()
Case " arrête"
Call onoffdecodeur ()
Case " arrête le"
Call onoffdecodeur ()
Case " mute"
Call mutedecodeur ()
Case " mute le son"
Call mutedecodeur ()
Case " met en muet"
Call mutedecodeur ()
Case " mais en muet"
Call mutedecodeur ()
Case " mettre en muet"
Call mutedecodeur ()
Case " muet"
Call mutedecodeur ()
Case " coupe le son"
Call mutedecodeur ()
Case " coupe le volume"
Call mutedecodeur ()
Case "  enlève le son"
Call mutedecodeur ()
Case "  enlève le volume"
Call mutedecodeur ()
Case " remet le son"
Call mutedecodeur ()
Case " met le son"
Call mutedecodeur ()
Case " augmente le son"
Call VolumeUpdecodeur ()
Case " augmente le volume"
Call VolumeUpdecodeur ()
Case " met le son plus fort"
Call VolumeUpdecodeur ()
Case " monte le son"
Call VolumeUpdecodeur ()
Case " monte le volume"
Call VolumeUpdecodeur ()
Case " baisse le volume"
Call VolumeDowndecodeur ()
Case " baisse le son"
Call VolumeDowndecodeur ()
Case " descend le son"
Call VolumeDowndecodeur ()
Case " descend le volume"
Call VolumeDowndecodeur ()
Case " chaîne suivante"
Call ChaineUpdecodeur ()
Case " mets la chaîne suivante"
Call ChaineUpdecodeur ()
Case " met la chaîne suivante"
Call ChaineUpdecodeur ()
Case " mets la chaîne suivante"
Call ChaineUpdecodeur ()
Case " mets la chaîne suivante"
Call ChaineUpdecodeur ()
Case " chaîne précédente"
Call ChaineDowndecodeur ()
Case " mets la chaîne précédente"
Call ChaineDowndecodeur ()
Case " met la chaîne précédente"
Call ChaineDowndecodeur ()
Case " descend de chaîne"
Call ChaineDowndecodeur ()
Case " mets la chaîne précédente"
Call ChaineDowndecodeur ()
Case " met la chaîne précédente"
Call ChaineDowndecodeur ()
Case " clique sur OK"
call OKdecodeur ()
Case " clique sur ok"
call OKdecodeur ()
Case " met OK"
call OKdecodeur ()
Case " met ok"
call OKdecodeur ()
Case " valide"
call OKdecodeur ()
Case " ok"
call OKdecodeur ()
Case " OK"
call OKdecodeur ()
Case " cliquer sur ok"
call OKdecodeur ()
Case " cliquer sur OK"
call OKdecodeur ()
Case " cliquez sur ok"
call OKdecodeur ()
Case " cliquez sur OK"
call OKdecodeur ()
Case " clique ok"
call OKdecodeur ()
Case " clique OK"
call OKdecodeur ()
Case " clic ok"
call OKdecodeur ()
Case " clic OK"
call OKdecodeur ()
Case " mais ok"
call OKdecodeur ()
Case " mais OK"
call OKdecodeur ()
Case " mettre 0"
chaine = "0"
call chainevalue(chaine)
Case " 0"
chaine = "0"
call chainevalue(chaine)
Case " met la chaîne zéro"
chaine = "0"
call chainevalue(chaine)
Case " met la chaîne 0"
chaine = "0"
call chainevalue(chaine)
Case " mais zéro"
chaine = "0"
call chainevalue(chaine)
Case " mes héros"
chaine = "0"
call chainevalue(chaine)
Case " mettre zéro"
chaine = "0"
call chainevalue(chaine)
Case " mais la mosaïque"
chaine = "0"
call chainevalue(chaine)
Case " mettre la mosaïque"
chaine = "0"
call chainevalue(chaine)
Case " la mosaïque"
chaine = "0"
call chainevalue(chaine)
Case " mosaïque"
chaine = "0"
call chainevalue(chaine)
Case " Game One"
chaine = "76"
call chainevalue(chaine)
Case " six"
chaine = "6"
call chainevalue(chaine)
Case " une"
chaine = "1"
call chainevalue(chaine)
Case " met la une"
chaine = "1"
call chainevalue(chaine)
Case " un"
chaine = "1"
call chainevalue(chaine)
Case " met la un"
chaine = "1"
call chainevalue(chaine)
Case " neuf"
chaine = "9"
call chainevalue(chaine)
Case Else
'Inputbox "La valeur n'existe pas","Erreur : valeur n'existe pas",a
End Select



sub onoffdecodeur ()
on error resume next
Dim IE
Set IE = Wscript.CreateObject("InternetExplorer.Application")
IE.Visible = 0
IE.navigate "http://" & IpDecodeur & ":8080/remoteControl/cmd?operation=01&key=116&mode=0"
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
on error resume next
Dim IE
Set IE = Wscript.CreateObject("InternetExplorer.Application")
IE.Visible = 0
IE.navigate "http://" & IpDecodeur & ":8080/remoteControl/cmd?operation=01&key=113&mode=0"
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
on error resume next
Dim IE
Set IE = Wscript.CreateObject("InternetExplorer.Application")
IE.Visible = 0
IE.navigate "http://" & IpDecodeur & ":8080/remoteControl/cmd?operation=01&key=115&mode=0"
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
on error resume next
Dim IE
Set IE = Wscript.CreateObject("InternetExplorer.Application")
IE.Visible = 0
IE.navigate "http://" & IpDecodeur & ":8080/remoteControl/cmd?operation=01&key=114&mode=0"
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

sub ChaineUpdecodeur ()
on error resume next
Dim IE
Set IE = Wscript.CreateObject("InternetExplorer.Application")
IE.Visible = 0
IE.navigate "http://" & IpDecodeur & ":8080/remoteControl/cmd?operation=01&key=402&mode=0"
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


sub ChaineDowndecodeur ()
on error resume next
Dim IE
Set IE = Wscript.CreateObject("InternetExplorer.Application")
IE.Visible = 0
IE.navigate "http://" & IpDecodeur & ":8080/remoteControl/cmd?operation=01&key=403&mode=0"
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

sub OKdecodeur ()
on error resume next
Dim IE
Set IE = Wscript.CreateObject("InternetExplorer.Application")
IE.Visible = 0
IE.navigate "http://" & IpDecodeur & ":8080/remoteControl/cmd?operation=01&key=352&mode=0"
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






Sub chainevalue (chaine)

'MsgBox chaine

Select Case chaine
Case "0"
chaineid = "512"
Case "1"
chaineid = "*******192"
Case "2"
chaineid = "*********4"
Case "3"
chaineid = "********80"
Case "4"
chaineid = "********34"
Case "5"
chaineid = "********47"
Case "6"
chaineid = "*******118"
Case "7"
chaineid = "*******111"
Case "8"
chaineid = "*******445"
Case "9"
chaineid = "*******119"
Case "10"
chaineid = "*******195"
Case "11"
chaineid = "*******446"
Case "12"
chaineid = "*******444"
Case "13"
chaineid = "*******234"
Case "14"
chaineid = "********78"
Case "15"
chaineid = "*******481"
Case "16"
chaineid = "*******226"
Case "17"
chaineid = "*******458"
Case "18"
chaineid = "*******482"
Case "19"
chaineid = "*******160"
Case "20"
chaineid = "******1404"
Case "21"
chaineid = "******1401"
Case "22"
chaineid = "******1403"
Case "23"
chaineid = "******1402"
Case "24"
chaineid = "******1400"
Case "25"
chaineid = "******1399"
Case "26"
chaineid = "*******112"
Case "27"
chaineid = "******2111"
Case "28"
chaineid = "********-1"
Case "34"
chaineid = "*******191"
Case "40"
chaineid = "********33"
Case "41"
chaineid = "********35"
Case "42"
chaineid = "******1563"
Case "43"
chaineid = "*******657"
Case "44"
chaineid = "********30"
Case "45"
chaineid = "******1290"
Case "46"
chaineid = "******1304"
Case "47"
chaineid = "******1335"
Case "49"
chaineid = "********-1"
Case "50"
chaineid = "*******730"
Case "51"
chaineid = "*******733"
Case "52"
chaineid = "*******732"
Case "53"
chaineid = "*******734"
Case "76"
chaineid = "********87"
Case "77"
chaineid = "******1167"
Case "158"
chaineid = "******2153"

Case else 
MsgBox "autre " & chaine
Wscript.Quit ()
End select

on error resume next
Dim IE
Set IE = Wscript.CreateObject("InternetExplorer.Application")
IE.Visible = 0

If Len(chaineid) = "10" then 
IE.navigate "http://" & IpDecodeur & ":8080/remoteControl/cmd?operation=09&epg_id=" & chaineid & "&uui=1"
Else
IE.navigate "http://" & IpDecodeur & ":8080/remoteControl/cmd?operation=01&key=" & chaineid & "&mode=0"
End if
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'iexplore.exe'")
For Each objProcess in colProcessList
    objProcess.Terminate()
Next
End sub
