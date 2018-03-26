'Ok Google, sur le pc XXX
'Ok Google, sur l'ordinateur XXX
' Applet IFTTT : https://ifttt.com/applets/jSNrZ4vJ-controle-de-l-ordinateur-avec-google-assitant
' Fichier GoogleHome.vbs a mettre dans : C:\GoogleHome
Dim MAJ
MAJ = "1.1" 'Version Actuelle du script

Dim WshShell,fso
On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell") 
checkregistre = WshShell.RegRead ("HKCU\Software\GoogleHome\Ok")

If err.Number<>0 or IsNull(checkregistre) Then
WshShell.RegWrite "HKCU\Software\GoogleHome\MAJ",MAJ,"REG_SZ"
WshShell.RegWrite "HKCU\Software\GoogleHome\Ok","1","REG_SZ"
MsgBox "Bienvenue dans mon script, il semblerait que vous lancer mon script pour la première fois ou que vous avez effectuer une mise à jour de celui-ci, pour faire fonctionner mon script dite : Ok Google, sur le pc xxx" & vbcr & "Par exemple Ok Google sur le pc test (pour tester la communication entre la Google homme est le PC)" & vbcr & " Dite des phrases simples et courtes" & vbcr & "Exercuté le script depuis l'ordinateur pour en savoir plus" & vbcr & vbcr & "Version Actuelle : " & MAJ ,vbInformation+vbOKOnly,"Control Google Home.vbs"

If MsgBox ("Voulez vous configuez le chemin d'accès pour la musiques ? " &vbcr & vbcr & "Sélectionner un dossier afin d'y rechercher des chansons dans ses sous-dossiers et ses sous-dossiers. Dossier par défaut" & vbcr & "Ok google sur le pc met de la musique" & vbcr & vbcr & "Si le dossier n'est pas configué, cela marchera quand même mais affichera un choix de dossier musique a chaque demande de musique" & vbcr & vbcr & "Oui = Configuer",vbyesno,"Configurez le dossier Musique") = vbYes Then
Dim objShell,objFolder,Message, user
user = wshShell.ExpandEnvironmentStrings( "%USERPROFILE%" )
	Message = "Veuillez sélectionner un dossier afin d'y rechercher des chansons dans ses sous-dossiers et ses sous-dossiers."
	Set objShell = CreateObject("Shell.Application")
	Set objFolder = objShell.BrowseForFolder(0,Message,1,user)
	If objFolder Is Nothing Then
		Wscript.Quit
	End If
	WshShell.RegWrite "HKCU\Software\GoogleHome\MUSIC",objFolder.self.path,"REG_SZ"
	MsgBox "Je conseil de tester la commande <musique> pour vérifier que tout fonctionne bien et que le lecteur media est compatible"
End if

If MsgBox ("Voulez vous configuez le chemin d'accès pour les videos ? " &vbcr & vbcr & "Sélectionner un dossier afin d'y rechercher des chansons dans ses sous-dossiers et ses sous-dossiers. Dossier par défaut" & vbcr & "Ok google sur le pc met de les vidéos" & vbcr & vbcr & "Si le dossier n'est pas configué, cela marchera quand même mais affichera un choix de dossier vidéos a chaque demande de musique" & vbcr & vbcr & "Oui = Configuer",vbyesno,"Configurez le dossier Viéo") = vbYes Then
user = wshShell.ExpandEnvironmentStrings( "%USERPROFILE%" )
	Message = "Veuillez sélectionner un dossier afin d'y rechercher des videos dans ses sous-dossiers et ses sous-dossiers."
	Set objShell = CreateObject("Shell.Application")
	Set objFolder = objShell.BrowseForFolder(0,Message,1,user)
	If objFolder Is Nothing Then
		Wscript.Quit
	End If
	WshShell.RegWrite "HKCU\Software\GoogleHome\VIDEO",objFolder.self.path,"REG_SZ"
	MsgBox "Je conseil de tester la commande <vidéo> pour vérifier que tout fonctionne bien et que le lecteur media est compatible"
End if
End if 

Set objArgs = WScript.Arguments
For I = 0 to objArgs.Count -1
a = a & " " & objArgs(I)	
Next
'inputbox a,a,a

If a = "" then
rep = InputBox ("Bienvenue dans mon script, pour faire fonctionner mon script dite : Ok Google, sur le pc xxx" & vbcr & "Par exemple Ok Google sur le pc test (pour tester la communication entre la Google homme est le PC)" & vbcr & vbcr & " Dite des phrases simples et courtes" & vbcr & vbcr & vbcr & "1 = Vérifier mise a jours" & vbcr & "2 = Envoyé un messsage au créateur (rapide & sans se logger)" & vbcr & "3 = Réinsalisé la configuration du script." & vbCr & "4 = Crédit" & vbcr & vbcr & "Pour tester des commandes en ecrit il vous suffit de tapez une commande si dessous pour valoir si elle est comprise par le logiciel ","Control Google Home","augmente le son")
   If rep = "" then
   ElseIf rep = "1" then 
   MAJCheck ()
   ElseIf rep = "2" then 
   MeParler () 
   ElseIf rep = "3" then 
   Reset ()
   ElseIf rep = "4" then 
   MsgBox "Crédits : " & vbNewLine & vbNewLine & "HackooFr" & vbNewLine & "Aymkdn",vbInformation+vbOKOnly,"Crédits"
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
Case " wifi"
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
Case " musique"
Musique ()
Case " met de la musique"
Musique ()
Case " lance de la musique"
Musique ()
Case " mais de la musique"
Musique ()
Case " lance musique"
Musique ()
Case " audio"
Musique ()
Case " met la musique"
Musique ()
Case " met la playlist"
Musique ()
Case " lance la playlist"
Musique ()
Case " met la playlist"
Musique ()
Case " vidéo"
Video ()
Case " film"
Video ()
Case " met vidéo"
Video ()
Case " film"
Video ()
Case " mais vidéo"
Video ()
Case " lance vidéo"
Video ()
Case " lance film"
Video ()
Case " met les vidéos"
Video ()
Case " met la vidéo"
Video ()
Case " lance la vidéo"
Video ()
Case " met le film"
Video ()
Case " met les films"
Video ()
Case " lance la vidéo"
Video ()
Case " met la vidéo"
Video ()



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
Set objHTTP=CreateObject("MSXML2.XMLHTTP") 
objHTTP.Open "GET", "https://dl.dropboxusercontent.com/s/c4cjk3bz535jpvd/Version.txt?dl=0", FALSE
objHTTP.Send
NewVersion = objHTTP.ResponseText
If VersionActu = NewVersion Then 
MsgBox "Logiciel à jour" & vbcr & vbcr & "Version actuelle : " & VersionActu,vbInformation,"Logiciel a jour"
Else
objHTTP.Open "GET", "https://dl.dropboxusercontent.com/s/jthogx3h7mozpk2/Note.txt?dl=0", FALSE
objHTTP.Send
Note = objHTTP.ResponseText
rep = MsgBox ("Votre version n'est pas à jour" & vbcr & vbcr & "Note du developpeur : " & Note & vbcr & vbcr &  "Voulez vous télécharger la derniere version ?" & vbcr &  "Version actuelle : " & VersionActu & vbcr & "Version disponible : " & NewVersion,vbInformation+vbYesNo)
  If rep = vbYes then
    objHTTP.Open "GET", "https://dl.dropboxusercontent.com/s/gybtf2i13bglxh7/GoogleHome.txt?dl=0", FALSE
    objHTTP.Send
    Telecharger = objHTTP.ResponseText
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
Set objHTTP=CreateObject("MSXML2.XMLHTTP") 
objHTTP.Open "GET", "https://pastebin.com/raw/6MadNeRY", FALSE
objHTTP.Send
Telecharger = objHTTP.ResponseText
Const ForWriting = 2 
Dim f
Set f = fso.OpenTextFile("C:\GoogleHome\WifiPasswordsRecovery.bat", ForWriting,true) 
f.write(Telecharger)
f.close
WScript.Sleep 100
End if
	WshShell.Run "C:\GoogleHome\WifiPasswordsRecovery.bat"
End sub 

Sub Video ()
If FSO.FileExists("C:\GoogleHome\LancerDossierVideo.vbs") = true then 
Else
Set objHTTP=CreateObject("MSXML2.XMLHTTP") 
objHTTP.Open "GET", "https://pastebin.com/raw/jthwXjWi", FALSE
objHTTP.Send
Telecharger = objHTTP.ResponseText
Const ForWriting = 2 
Dim f
Set f = fso.OpenTextFile("C:\GoogleHome\LancerDossierVideo.vbs", ForWriting,true) 
f.write(Telecharger)
f.close
WScript.Sleep 100
End if
	WshShell.Run "C:\GoogleHome\LancerDossierVideo.vbs"
End sub 

Sub Musique ()
If FSO.FileExists("C:\GoogleHome\LancerDossierMusique.vbs") = true then 
Else
Set objHTTP=CreateObject("MSXML2.XMLHTTP") 
objHTTP.Open "GET", "https://pastebin.com/raw/7qjZSaHR", FALSE
objHTTP.Send
Telecharger = objHTTP.ResponseText
Const ForWriting = 2 
Dim f
Set f = fso.OpenTextFile("C:\GoogleHome\LancerDossierMusique.vbs", ForWriting,true) 
f.write(Telecharger)
f.close
WScript.Sleep 100
End if
	WshShell.Run "C:\GoogleHome\LancerDossierMusique.vbs"
End sub 


Sub Reset ()
On Error Resume Next
WshShell.RegDelete "HKCU\Software\GoogleHome\Ok"
WshShell.RegDelete "HKCU\Software\GoogleHome\MUSIC"
WshShell.RegDelete "HKCU\Software\GoogleHome\VIDEO"
WshShell.RegDelete "HKCU\Software\GoogleHome\MAJ"
MsgBox "La configuration de GoogleHome.vbs a étais Réinsalisé.",vbInformation+vbOKOnly,"Succès Réinsalisation GoogleHome.vbs "
WshShell.Run "C:\GoogleHome\GoogleHome.vbs"
End sub
