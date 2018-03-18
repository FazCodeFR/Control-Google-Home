'Ok Google, sur le pc XXX
'Ok Google, sur l'ordinateur XXX
' Applet IFTTT : https://ifttt.com/applets/jSNrZ4vJ-controle-de-l-ordinateur-avec-google-assitant
' Fichier GoogleHome.vbs a mettre dans : C:\GoogleHome
Dim MAJ
MAJ = "1.0" 'Version Actuelle du script

Dim WshShell,fso
On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell") 
checkregistre = WshShell.RegRead ("HKCU\Software\GoogleHome\Ok")

If err.Number<>0 or IsNull(checkregistre) Then
WshShell.RegWrite "HKCU\Software\GoogleHome\MAJ",MAJ,"REG_SZ"
WshShell.RegWrite "HKCU\Software\GoogleHome\Ok","1","REG_SZ"
MsgBox "Bienvenue dans mon script, il semblerait que vous lancer mon script pour la première fois ou que vous avez effectuer une mise à jour de celui-ci, pour faire fonctionner mon script dite : Ok Google, sur le pc xxx" & vbcr & "Par exemple Ok Google sur le pc test (pour tester la communication entre la Google homme est le PC)" & vbcr & " Dite des phrases simples et courtes" & vbcr & "Exercuté le script depuis l'ordinateur pour en savoir plus" & vbcr & vbcr & "Version Actuelle : " & MAJ ,vbInformation+vbOKOnly,"Control Google Home.vbs"
End if 

Set objArgs = WScript.Arguments
For I = 0 to objArgs.Count -1
a = a & " " & objArgs(I)	
Next
'inputbox a,a,a

If a = "" then
rep = InputBox ("Bienvenue dans mon script, pour faire fonctionner mon script dite : Ok Google, sur le pc xxx" & vbcr & "Par exemple Ok Google sur le pc test (pour tester la communication entre la Google homme est le PC)" & vbcr & vbcr & " Dite des phrases simples et courtes" & vbcr & vbcr & vbcr & "1 = Vérifier mise a jours" & vbcr & "2 = Envoyé un messsage au créateur (rapide & sans se logger)" & vbcr & vbcr & "Pour tester des commandes en ecrit il vous suffit de tapez une commande si dessous pour valoir si elle est comprise par le logiciel ","Control Google Home","augmente le son")
   If rep = "" then
   ElseIf rep = "1" then 
   MAJCheck ()
   ElseIf rep = "2" then 
   MeParler () 
   Else
   a = " " & rep
   End if 
end if 

Set WshShell = CreateObject("WScript.Shell")
Select Case a
Case " test"
MsgBox "La Google Home communique bien avec le pc !" & vbcr & vbcr & "Succès test",vbinformation+vbOKOnly,"Test"
Case " teste"
MsgBox "La Google Home communique bien avec le pc !" & vbcr & vbcr & "Succès test",vbinformation+vbOKOnly,"Test"
Case " augmente le son"
WshShell.SendKeys "{" & chr(175) & " 10}"
Case " augmente le volume"
WshShell.SendKeys "{" & chr(175) & " 10}"
Case " monte le son"
WshShell.SendKeys "{" & chr(175) & " 10}"
Case " monte le son au max"
WshShell.SendKeys "{" & chr(175) & " 50}"
Case " monte le son au maximum"
WshShell.SendKeys "{" & chr(175) & " 50}"
Case " monte le volume au maximum"
WshShell.SendKeys "{" & chr(175) & " 50}"
Case " volume max"
WshShell.SendKeys "{" & chr(175) & " 50}"
Case " son au max"
WshShell.SendKeys "{" & chr(175) & " 50}"
Case " augmente le son au maximum"
WshShell.SendKeys "{" & chr(175) & " 50}"
Case " baisse le son"
WshShell.SendKeys "{" & chr(174) & " 10}"
Case " descend le son"
WshShell.SendKeys "{" & chr(174) & " 10}"
Case " descend le volume"
WshShell.SendKeys "{" & chr(174) & " 10}"
Case " descend le son au max"
WshShell.SendKeys "{" & chr(174) & " 50}"
Case " baisse le son au max"
WshShell.SendKeys "{" & chr(174) & " 50}"
Case " baisse le volume"
WshShell.SendKeys "{" & chr(174) & " 10}"
Case " baisse le volume au max"
WshShell.SendKeys "{" & chr(174) & " 50}"
Case " baisse le son au maximum"
WshShell.SendKeys "{" & chr(174) & " 50}"
Case " baisse le volume au maximum"
WshShell.SendKeys "{" & chr(174) & " 50}"
Case " mute le volume"
WshShell.SendKeys chr(173)
Case " mute le son"
WshShell.SendKeys chr(173)
Case " muet"
WshShell.SendKeys chr(173)
Case " le son a 0"
WshShell.SendKeys chr(173)
Case " coupe le son"
WshShell.SendKeys chr(173)
Case " coupe le volume"
WshShell.SendKeys chr(173)
Case " coupe l'audio"
WshShell.SendKeys chr(173)
Case " remet le volume"
WshShell.SendKeys chr(173)
Case " remet le son"
WshShell.SendKeys chr(173)
Case " lance chrome"
WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Case " affiche chrome"
WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Case " ouvre chrome"
WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Case " démarre chrome"
WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Case " exécute chrome"
WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Case " lance google chrome"
WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Case " ouvre google chrome"
WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Case " démarre google chrome"
WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Case " exécute google chrome"
WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Case " lance google"
WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Case " ouvre google"
WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Case " démarre google"
WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Case " exécute chrome"
WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Case " lorsque Rome"
WshShell.Run """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Case " lance VLC"
WshShell.Run """C:\Program Files (x86)\VideoLAN\VLC\vlc.exe"""
Case " ouvre VLC"
WshShell.Run """C:\Program Files (x86)\VideoLAN\VLC\vlc.exe"""
Case " démarre VLC"
WshShell.Run """C:\Program Files (x86)\VideoLAN\VLC\vlc.exe"""
Case " exécute VLC"
WshShell.Run """C:\Program Files (x86)\VideoLAN\VLC\vlc.exe"""
Case " fait pause"
WshShell.SendKeys " "
Case " met pause"
WshShell.SendKeys " "
Case " fais une pause"
WshShell.SendKeys " "
Case " met en pause"
WshShell.SendKeys " "
Case " mais en pause"
WshShell.SendKeys " "
Case " fait pause"
WshShell.SendKeys " "
Case " fait stop"
WshShell.SendKeys " "
Case " stop"
WshShell.SendKeys " "
Case " pause"
WshShell.SendKeys " "
Case " relance"
WshShell.SendKeys " "
Case " enlève la pause"
WshShell.SendKeys " "
Case " met une pause"
WshShell.SendKeys " "
Case " lance"
WshShell.SendKeys " "
Case " lecture"
WshShell.SendKeys " "
Case " lance lecture"
WshShell.SendKeys " "
Case " lance la lecture"
WshShell.SendKeys " "
Case " éteint le"
CreateObject("Wscript.Shell").Run "CMD /C " & " shutdown /s /f"
Case " arrête le"
CreateObject("Wscript.Shell").Run "CMD /C " & " shutdown /s /f"
Case " éteint le pc"
CreateObject("Wscript.Shell").Run "CMD /C " & " shutdown /s /f"
Case " éteint l'ordinateur"
CreateObject("Wscript.Shell").Run "CMD /C " & " shutdown /s /f"
Case " arrête le système"
CreateObject("Wscript.Shell").Run "CMD /C " & " shutdown /s /f"
Case " éteint le système"
CreateObject("Wscript.Shell").Run "CMD /C " & " shutdown /s /f"
Case " verrouille le"
WshShell.Run "rundll32.exe user32.dll,LockWorkStation"
Case " verrouille la session"
WshShell.Run "rundll32.exe user32.dll,LockWorkStation"
Case " verrouille le pc"
WshShell.Run "rundll32.exe user32.dll,LockWorkStation"
Case " verrouille l'ordinateur"
WshShell.Run "rundll32.exe user32.dll,LockWorkStation"
Case " le verrouiller"
WshShell.Run "rundll32.exe user32.dll,LockWorkStation"
Case " met en veille"
WshShell.Run "rundll32.exe user32.dll,LockWorkStation"
Case " mettre en veille"
WshShell.Run "rundll32.exe user32.dll,LockWorkStation"
Case " met le en veille"
WshShell.Run "rundll32.exe user32.dll,LockWorkStation"
Case " veille"
WshShell.Run "rundll32.exe user32.dll,LockWorkStation"
Case " verrouillage"
WshShell.Run "rundll32.exe user32.dll,LockWorkStation"
Case " mot de passe wifi"
WifiPasswordsRecovery ()
Case " mot de passe du wifi"
WifiPasswordsRecovery ()
Case " code wifi"
WifiPasswordsRecovery ()
Case " code du wifi"
WifiPasswordsRecovery ()
Case " donne le mot de passe wifi"
WifiPasswordsRecovery ()
Case " donne le mot de passe du wifi"
WifiPasswordsRecovery ()
Case " retrouve le mot de passe wifi"
WifiPasswordsRecovery ()
Case " retrouve le mot de passe du wifi"
WifiPasswordsRecovery ()
Case " quel est le mot de passe wifi"
WifiPasswordsRecovery ()
Case " quel est le mot de passe du wifi"
WifiPasswordsRecovery ()
Case " ouvre le lecteur CD"
LecteurDVD ()
Case " ouvre le lecteur DVD"
LecteurDVD ()
Case " ouvrir lecteur"
LecteurDVD ()
Case " ouvrir le lecteur CD"
LecteurDVD ()
Case " ouvrir le lecteur DVD"
LecteurDVD ()
Case " ouvrir le lecteur"
LecteurDVD ()
Case " ouvre le lecteur"
LecteurDVD ()
Case " eject le CD"
LecteurDVD ()
Case " eject le DVD"
LecteurDVD ()
Case " eject CD"
LecteurDVD ()
Case " eject DVD"
LecteurDVD ()
Case " ferme le logiciel"
WshShell.SendKeys ("%{F4}")
Case " ferme le logiciel actif"
WshShell.SendKeys ("%{F4}")
Case " ferme l'application"
WshShell.SendKeys ("%{F4}")
Case " arrête le logiciel"
WshShell.SendKeys ("%{F4}")
Case " arrête l'application"
WshShell.SendKeys ("%{F4}")
Case " maj"
MAJCheck ()
Case " mise à jour"
MAJCheck ()
Case " vérifier mise à jour"
MAJCheck ()
Case " vérifie mise à jour"
MAJCheck ()
Case " mise à jour script"
MAJCheck ()
Case " vérifier"
MAJCheck ()
Case " mage"
MAJCheck ()
Case Else
'Inputbox "La valeur n'existe pas","Erreur : valeur n'existe pas",a
End Select

Sub LecteurDVD ()
Set oWMP = CreateObject("WMPlayer.OCX.7" ) 
Set colCDROMs = oWMP.cdromCollection 
if colCDROMs.Count >= 1 then 
For i = 0 to colCDROMs.Count - 1 
colCDROMs.Item(i).Eject 
colCDROMs.Item(i).Eject 
Next 
End if
End sub 


Sub MAJCheck ()
Dim IE,VersionActu, NewVersion,Note,WshShell
Set WshShell = WScript.CreateObject("WScript.Shell") 
VersionActu = WshShell.RegRead ("HKCU\Software\GoogleHome\MAJ")
Set IE = Wscript.CreateObject("InternetExplorer.Application")
IE.Visible = 0
IE.navigate "https://dl.dropboxusercontent.com/s/c4cjk3bz535jpvd/Version.txt?dl=0" 
While IE.ReadyState <> 4 : WScript.Sleep 100 : Wend
NewVersion = IE.Document.body.innerText
IE.Quit
If VersionActu = NewVersion Then 
MsgBox "Logiciel à jour" & vbcr & vbcr & "Version actuelle : " & VersionActu,vbInformation,"Logiciel a jour"
Else
Set IE = Wscript.CreateObject("InternetExplorer.Application")
IE.navigate "https://dl.dropboxusercontent.com/s/jthogx3h7mozpk2/Note.txt?dl=0" 
While IE.ReadyState <> 4 : WScript.Sleep 100 : Wend
Note =  IE.Document.body.innerText
IE.Quit
rep = MsgBox ("Votre version n'est pas à jour" & vbcr & vbcr & "Note du developpeur : " & Note & vbcr & vbcr &  "Voulez vous télécharger la derniere version ?" & vbcr &  "Version actuelle : " & VersionActu & vbcr & "Version disponible : " & NewVersion,vbInformation+vbYesNo)
  If rep = vbYes then
	Set IE = Wscript.CreateObject("InternetExplorer.Application")
	IE.navigate "https://dl.dropboxusercontent.com/s/gybtf2i13bglxh7/GoogleHome.txt?dl=0"
	While IE.ReadyState <> 4 : WScript.Sleep 100 : Wend
	Telecharger =  IE.Document.body.innerText
	IE.Quit
    Const ForWriting = 2 
    Dim f
    Set f = fso.OpenTextFile("C:\GoogleHome\GoogleHomeNew.txt", ForWriting,true) 
    f.write(Telecharger)
    f.close
    WshShell.RegDelete "HKCU\Software\GoogleHome\Ok"
    WshShell.Run "cmd /c chcp 28591 > nul & taskkill /F /IM wscript.exe & move C:\GoogleHome\GoogleHomeNew.txt C:\GoogleHome\GoogleHome.vbs & start C:\GoogleHome\GoogleHome.vbs",0
   Else 
    Wscript.Quit ()
   End if
End if
End sub

Sub MeParler ()
Dim IE
Set IE = Wscript.CreateObject("InternetExplorer.Application")
IE.Visible = 0
IE.navigate "https://aboatdev.sarahah.com/" 
Texte = Inputbox ("M'envoyer un message : (constructif de préférence)" & vbcr & vbcr & vbcr & "N'oubliez pas de rajouter un moyen de contact si vous voulez être recontacter." & vbcr & vbcr & "-Envoie rapide & anonyme-","Envoyer un message au créateur")
If Texte = "" then 
MsgBox "Message annuler."
Else
While IE.ReadyState <> 4 : WScript.Sleep 100 : Wend
WScript.Sleep 1000
IE.Document.All.Item("Text").Value = Texte & vbcr & vbcr & "Version : " & VersionActu
WScript.Sleep 1000
IE.Document.All.Item("Send").click
MsgBox "Message envoyé !",vbInformation+vbOKOnly,"Message Envoyé"
IE.Quit
End If 
End sub


Sub WifiPasswordsRecovery ()
If FSO.FileExists("C:\GoogleHome\WifiPasswordsRecovery.bat") = true then 
Else
Dim IE
Set IE = Wscript.CreateObject("InternetExplorer.Application")
IE.Visible = 0
IE.navigate "https://pastebin.com/raw/6MadNeRY"
While IE.ReadyState <> 4 : WScript.Sleep 100 : Wend
WScript.Sleep 1000
Telecharger =  IE.Document.body.innerText
IE.Quit
Const ForWriting = 2 
Dim f
Set f = fso.OpenTextFile("C:\GoogleHome\WifiPasswordsRecovery.bat", ForWriting,true) 
f.write(Telecharger)
f.close
WScript.Sleep 100
End if
	WshShell.Run "C:\GoogleHome\WifiPasswordsRecovery.bat"
End sub 
