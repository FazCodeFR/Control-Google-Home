' Ok Google, sur le pc XXX
' Ok Google, sur l'ordinateur XXX
' Applet IFTTT : https://ifttt.com/applets/jSNrZ4vJ-controle-de-l-ordinateur-avec-google-assitant
' Projet  : https://github.com/ABOATDev/Control-Google-Home

Dim MAJ, WS,fso,CheckMAJUser,f,IE,objHTTP
MAJ = "1.0.7" 'Version Actuelle du script

On Error Resume Next

Set fso = CreateObject("Scripting.FileSystemObject")
Set WS = WScript.CreateObject("WScript.Shell") 
Set IE = Wscript.CreateObject("InternetExplorer.Application")
Set objHTTP=CreateObject("MSXML2.XMLHTTP") 


If err.Number<>0 or IsNull(WS.RegRead ("HKCU\Software\GoogleHome\Ok")) Then
Call MAJCheck (CheckMAJUser, MAJ)
WS.RegWrite "HKCU\Software\GoogleHome\MAJ",MAJ,"REG_SZ"
WS.RegWrite "HKCU\Software\GoogleHome\Ok","1","REG_SZ"
objHTTP.Open "GET", "https://raw.githubusercontent.com/ABOATDev/Control-Google-Home/master/ListeCommande.txt", FALSE
objHTTP.Send
Const ForWriting = 2
Set f = fso.OpenTextFile("C:\GoogleHome\ListeCommande.txt", ForWriting,true) 
f.write(objHTTP.ResponseText)
f.close
MsgBox "Bienvenue dans mon script, il semblerait que vous lancer mon script pour la première fois ou que vous avez effectuer une mise à jour de celui-ci, pour faire fonctionner mon script dite : Ok Google, sur le pc xxx" & vbcr & "Par exemple Ok Google sur le pc test (pour tester la communication entre la Google homme est le PC)" & vbcr & " Dite des phrases simples et courtes" & vbcr & "Exercuté le script depuis l'ordinateur pour en savoir plus" & vbcr & vbcr & "Version Actuelle : " & MAJ ,vbInformation+vbOKOnly,"Control Google Home.vbs"
If err.Number<>0 or IsNull(WS.RegRead ("HKCU\Software\GoogleHome\MUSIC")) Then
If MsgBox ("Voulez vous configuez le chemin d'accès pour la musiques ? " &vbcr & vbcr & "Sélectionner un dossier afin d'y rechercher des chansons dans ses sous-dossiers et ses sous-dossiers. Dossier par défaut" & vbcr & "Ok google sur le pc met de la musique" & vbcr & vbcr & "Si le dossier n'est pas configué, cela marchera quand même mais affichera un choix de dossier musique a chaque demande de musique" & vbcr & vbcr & "Oui = Configuer",vbyesno,"Configurez le dossier Musique") = vbYes Then
Dim objShell,objFolder,Message, user
user = WS.ExpandEnvironmentStrings( "%USERPROFILE%" )
	Message = "Veuillez sélectionner un dossier afin d'y rechercher des chansons dans ses sous-dossiers et ses sous-dossiers."
	Set objShell = CreateObject("Shell.Application")
	Set objFolder = objShell.BrowseForFolder(0,Message,1,user)
	If objFolder Is Nothing Then Wscript.Quit
	WS.RegWrite "HKCU\Software\GoogleHome\MUSIC",objFolder.self.path,"REG_SZ"
	MsgBox "Je conseil de tester la commande <musique> pour vérifier que tout fonctionne bien et que le lecteur media est compatible"
End if
End if 
If err.Number<>0 or IsNull(WS.RegRead ("HKCU\Software\GoogleHome\VIDEO")) Then
If MsgBox ("Voulez vous configuez le chemin d'accès pour les vidéos ? " &vbcr & vbcr & "Sélectionner un dossier afin d'y rechercher des chansons dans ses sous-dossiers et ses sous-dossiers. Dossier par défaut" & vbcr & "Ok google sur le pc met de les vidéos" & vbcr & vbcr & "Si le dossier n'est pas configué, cela marchera quand même mais affichera un choix de dossier vidéos a chaque demande de musique" & vbcr & vbcr & "Oui = Configuer",vbyesno,"Configurez le dossier Vidéo") = vbYes Then
user = WS.ExpandEnvironmentStrings( "%USERPROFILE%" )
	Message = "Veuillez sélectionner un dossier afin d'y rechercher des videos dans ses sous-dossiers et ses sous-dossiers."
	Set objShell = CreateObject("Shell.Application")
	Set objFolder = objShell.BrowseForFolder(0,Message,1,user)
	If objFolder Is Nothing Then Wscript.Quit
	WS.RegWrite "HKCU\Software\GoogleHome\VIDEO",objFolder.self.path,"REG_SZ"
	MsgBox "Je conseil de tester la commande <vidéo> pour vérifier que tout fonctionne bien et que le lecteur media est compatible"
End if
End if 

End if 

Set objArgs = WScript.Arguments
For I = 0 to objArgs.Count -1
Select Case objArgs(I)
Case "écris", "écrit"
ecrit = true
Case "lance", "ouvre","affiche","démarre", "exécute","ouvrir","démarrer","exécuter"
lance = true
Case Else
a = a & " " & LCase(objArgs(I))
End Select
Next

If ecrit = true then
Call write(a)
End if

If lance = true then
Logiciel = right (a,len(a)-1)
Call launch (Logiciel)
End if 


'inputbox a,a,a

If a = "" then
Call MAJCheck (CheckMAJUser, MAJ)

rep = InputBox ("Bienvenue dans mon script, communication entre vos Assistants (Google Assistant, Google Home , Cortana, Alexa, ...) sur vos ordinateurs Windows" & vbNewLine &  "Pour faire fonctionner mon script dite : Ok Google, sur le pc xxx" & vbcr & "Par exemple Ok Google sur le pc test (pour tester la communication entre la Google homme est le PC)" & vbcr & vbcr & " Dite des phrases simples et courtes" & vbcr & vbcr & vbcr & "1 = Vérifier mise a jours" & vbcr & "2 = Envoyé un messsage au créateur (rapide & sans se logger)" & vbcr & "3 = Réinsalisé la configuration du script." & vbCr & "4 = Crédit" & vbcr & vbcr & "Pour tester des commandes en écrit, il vous suffit de taper une commande si dessous pour savoir si elle est comprise par le logiciel" & vbNewLine & "Version : " &  MAJ,"Control Google Home " & MAJ,"test")
   If rep = "" then
   WScript.Quit()
   ElseIf rep = "1" then 
   CheckMAJUser = true
   Call MAJCheck (CheckMAJUser, MAJ)
   Wscript.Quit
   ElseIf rep = "2" then 
   MeParler () 
   Wscript.Quit
   ElseIf rep = "3" then 
   Reset ()
   Wscript.Quit
   ElseIf rep = "4" then 
   MsgBox "Crédits : " & vbNewLine & vbNewLine & "HackooFr - Aide indirect pour le Script" & vbNewLine & "facebook.com/hackoo.crackoo" & vbNewLine & vbNewLine & "Aymkdn - Pour l'assistant-plugins" &  vbNewLine & " github.com/Aymkdn | paypal.me/aymkdn" & vbNewLine & vbNewLine & "Créateur du Contrôle de l'ordinateur avec Google Home : ABOAT " & vbNewLine & "facebook.com/aboat.hack",vbInformation+vbOKOnly,"Crédits"
   Wscript.Quit
   Else
   Dim i,tb 
   tb = split(rep," ") 
        For i = lbound(tb) to 0
		  if tb(i) = "lance" or tb(i) = "ouvre" or tb(i) = "affiche" or tb(i) = "démarre" or tb(i) = "exécute" or tb(i) = "ouvrir" or tb(i) =  "démarrer" or tb(i) = "exécuter" = True Then Call launch(right (rep,len(rep)-len(tb(i))-1))
		next
    a = " " & LCase(rep)
End if 
End if 

a = right (a,len(a)-1)
Select Case a


Case "test", "teste", "check", "ok","vérifie","vérification"
Call Check ()
Call MAJCheck (CheckMAJUser, MAJ)

Case "augmente le son","augmente le volume","monte le son"
WS.SendKeys "{" & chr(175) & " 10}"

Case "monte le son au max","monte le son au maximum","monte le volume au maximum","volume max","son au max","augmente le son au maximum"
WS.SendKeys "{" & chr(175) & " 50}"

Case "baisse le son","descend le son","descend le volume","baisse le volume"
WS.SendKeys "{" & chr(174) & " 10}"

Case "descend le son au max","baisse le son au max","baisse le volume au max","baisse le son au maximum","baisse le volume au maximum"
WS.SendKeys "{" & chr(174) & " 50}"

Case "mute le volume","mute le son","muet","le son a 0","coupe le son","coupe le volume","coupe l'audio","remet le volume","remet le son"
WS.SendKeys chr(173)

Case "fait pause","met pause","fais une pause","met en pause","mais en pause","fait pause","fait stop","stop","pause","relance","enlève la pause","met une pause","lance","lecture","lance lecture","lance la lecture"
WS.SendKeys " "

Case "éteint le","arrête le","éteint le pc","éteint l'ordinateur","arrête le système","éteint le système"," arrête"
CreateObject("Wscript.Shell").Run "CMD /C " & " shutdown /s /f",0

Case "verrouille le","verrouille la session","verrouille le pc","le verrouiller","met en veille","mettre en veille","met le en veille","veille","verrouillage","verrouille","metre en veille"
WS.Run "rundll32.exe user32.dll,LockWorkStation"

Case "mot de passe wifi","mot de passe du wifi","code wifi","wifi","code de la wifi","donne mot de passe wifi","code du wifi","donne le mot de passe wifi","donne le mot de passe du wifi","retrouve le mot de passe wifi","retrouve le mot de passe du wifi","quel est le mot de passe wifi","quel est le mot de passe du wifi"
WifiPasswordsRecovery ()

Case "ouvre le lecteur cd","ouvre le lecteur dvd","ouvrir lecteur","ouvrir le lecteur cd","ouvrir le lecteur dvd","ouvrir le lecteur","ouvre le lecteur","eject le cd","eject le dvd","eject cd","eject dvd","ejecter dvd","ejecter cd"," ejecter le dvd"
LecteurDVD ()

Case "ferme le logiciel","ferme le logiciel actif","ferme l'application","arrête le logiciel","arrête l'application"
WS.SendKeys ("%{F4}")

Case "maj","mise à jour","vérifier mise à jour","vérifie mise à jour","mise à jour script","vérifier"," mage"
CheckMAJUser = true
Call MAJCheck (CheckMAJUser, MAJ)

Case "musique","met de la musique","lance de la musique","mais de la musique","lance musique","audio","met la musique","met la playlist","lance la playlist","met la playlist"
Musique ()

Case "vidéo","film","met vidéo","film","mais vidéo","lance vidéo","lance film","met les vidéos","met la vidéo","lance la vidéo","met le film","met les films","lance la vidéo","met la vidéo"
Video ()

Case "eject usb", "eject clé usb", "eject la clé usb" , "retire usb" , "retire la clé usb","retire clé usb"
Eject_USB ()


case "liste des commandes", "liste commande", "liste des commandes" , "détail des commandes", "les commandes disponible"
ListeCommande ()

Case Else

Call  MAJCheck (CheckMAJUser, MAJ)
Call Suggestion (MAJ,a)
'Inputbox "La valeur n'existe pas","Erreur : valeur n'existe pas",a
End Select

Sub LecteurDVD ()
On Error Resume Next
Set oWMP = CreateObject("WMPlayer.OCX.7" ) 
Set colCDROMs = oWMP.cdromCollection 
if colCDROMs.Count >= 1 then 
For i = 0 to colCDROMs.Count - 1 
colCDROMs.Item(i).Eject 
colCDROMs.Item(i).Eject 
Next 
End if
End sub 


Sub MAJCheck (CheckMAJUser, MAJ)
On Error Resume Next
Dim VersionActu, NewVersion,Note
VersionActu = MAJ 
objHTTP.Open "GET", "https://raw.githubusercontent.com/ABOATDev/Control-Google-Home/master/Tools/Version", FALSE
objHTTP.Send
NewVersion = objHTTP.ResponseText
NewVersion = left(NewVersion, len(NewVersion) - 1) 
if NewVersion > VersionActu Then
	 If CheckMAJUser = true Then MsgBox "La version : " & NewVersion & " est disponible et va etre installé !" & vbNewLine & vbNewLine & "Notre version actuelle" & VersionActu,vbInformation+vbOKOnly,"Nouvelle version disponible"
     objHTTP.Open "GET", "https://raw.githubusercontent.com/ABOATDev/Control-Google-Home/master/GoogleHome.vbs", FALSE
     objHTTP.Send
     Telecharger = objHTTP.ResponseText
     Const ForWriting = 2 
     Dim f
     Set f = fso.OpenTextFile("C:\GoogleHome\GoogleHomeNew.txt", ForWriting,true) 
     f.write(Telecharger)
     f.close
	 CheckMAJUser = false
	 If err.Number<>0 or IsNull(WS.RegRead("HKCU\Software\GoogleHome\Ok")) Then
	 Else
	 WS.RegDelete "HKCU\Software\GoogleHome\Ok"
	 End if 
	 If err.Number<>0 or IsNull(WS.RegRead("HKCU\Software\GoogleHome\MAJ")) Then
	 Else
	 WS.RegDelete "HKCU\Software\GoogleHome\MAJ"
	 End if 
     Return = WS.Run ("cmd /k chcp 28591 > nul & taskkill /F /IM wscript.exe & move C:\GoogleHome\GoogleHomeNew.txt C:\GoogleHome\GoogleHome.vbs & start C:\GoogleHome\GoogleHome.vbs & exit",0,true)
	Else
	 If CheckMAJUser = true then MsgBox "Pas de nouvelle mise a jours a installer" & vbNewLine & "Vous êtes bien dans la derniere version disponible" & vbNewLine & vbNewLine & vbNewLine & "Votre version : " & VersionActu & vbNewLine & "Derniere version : " & NewVersion
	 CheckMAJUser = false
End if
End sub

Sub MeParler (MAJ)
On Error Resume Next
IE.Visible = 0
IE.navigate "https://aboatdev.sarahah.com/" 
Texte = Inputbox ("M'envoyer un message : (constructif de préférence)" & vbcr & vbcr & vbcr & "N'oubliez pas de rajouter un moyen de contact si vous voulez être recontacter." & vbcr & vbcr & "-Envoie rapide & anonyme-","Envoyer un message au créateur")
If Texte = "" then 
MsgBox "Message annuler."
Else
While IE.ReadyState <> 4 : WScript.Sleep 100 : Wend
WScript.Sleep 1000
IE.Document.All.Item("Text").Value = Texte & vbcr & vbcr & "Version : " & MAJ
WScript.Sleep 1000
IE.Document.All.Item("Send").click
While IE.ReadyState <> 4 : WScript.Sleep 100 : Wend
WScript.Sleep 1000
MsgBox "Message envoyé !",vbInformation+vbOKOnly,"Message Envoyé"
IE.Quit
End If 
End sub

Sub WifiPasswordsRecovery ()
On Error Resume Next
If FSO.FileExists("C:\GoogleHome\WifiPasswordsRecovery.bat") = true then 
Else
objHTTP.Open "GET", "https://raw.githubusercontent.com/ABOATDev/Control-Google-Home/master/Tools/WifiPasswordsRecovery.bat", FALSE
objHTTP.Send
Telecharger = objHTTP.ResponseText
Const ForWriting = 2 
Dim f
Set f = fso.OpenTextFile("C:\GoogleHome\WifiPasswordsRecovery.bat", ForWriting,true) 
f.write(Telecharger)
f.close
WScript.Sleep 100
End if
	WS.Run "C:\GoogleHome\WifiPasswordsRecovery.bat"
End sub 

Sub Video ()
On Error Resume Next
If FSO.FileExists("C:\GoogleHome\LancerDossierVideo.vbs") = true then 
Else
objHTTP.Open "GET", "https://raw.githubusercontent.com/ABOATDev/Control-Google-Home/master/Tools/LancerDossierVideo.vbs", FALSE
objHTTP.Send
Telecharger = objHTTP.ResponseText
Const ForWriting = 2
Set f = fso.OpenTextFile("C:\GoogleHome\LancerDossierVideo.vbs", ForWriting,true) 
f.write(Telecharger)
f.close
WScript.Sleep 100
End if
	WS.Run "C:\GoogleHome\LancerDossierVideo.vbs"
End sub 

Sub Musique ()
On Error Resume Next
If FSO.FileExists("C:\GoogleHome\LancerDossierMusique.vbs") = true then 
Else
objHTTP.Open "GET", "https://raw.githubusercontent.com/ABOATDev/Control-Google-Home/master/Tools/LancerDossierMusique.vbs", FALSE
objHTTP.Send
Telecharger = objHTTP.ResponseText
Const ForWriting = 2
Set f = fso.OpenTextFile("C:\GoogleHome\LancerDossierMusique.vbs", ForWriting,true) 
f.write(Telecharger)
f.close
WScript.Sleep 100
End if
	WS.Run "C:\GoogleHome\LancerDossierMusique.vbs"
End sub 


Sub Eject_USB ()
On Error Resume Next
If FSO.FileExists("C:\GoogleHome\Eject_USB.vbs") = true then 
Else
objHTTP.Open "GET", "https://raw.githubusercontent.com/ABOATDev/Control-Google-Home/master/Tools/Eject_USB.vbs", FALSE
objHTTP.Send
Telecharger = objHTTP.ResponseText
Const ForWriting = 2 
Dim f
Set f = fso.OpenTextFile("C:\GoogleHome\Eject_USB.vbs", ForWriting,true) 
f.write(Telecharger)
f.close
WScript.Sleep 100
End if
	WS.Run "C:\GoogleHome\Eject_USB.vbs"
End sub 

Sub ListeCommande ()
On Error Resume Next
objHTTP.Open "GET", "https://raw.githubusercontent.com/ABOATDev/Control-Google-Home/master/ListeCommande.txt", FALSE
objHTTP.Send
Telecharger = objHTTP.ResponseText
Const ForWriting = 2 
Dim f
Set f = fso.OpenTextFile("C:\GoogleHome\ListeCommande.txt", ForWriting,true) 
f.write(Telecharger)
f.close
WScript.Sleep 100
WS.Run "C:\GoogleHome\ListeCommande.txt"
End sub 

Sub write(a)
WScript.Sleep 300
WS.SendKeys right(a,len(a)-1)
WScript.Quit ()
End sub

Sub Reset ()
On Error Resume Next
WS.RegDelete "HKCU\Software\GoogleHome\Ok"
WS.RegDelete "HKCU\Software\GoogleHome\MUSIC"
WS.RegDelete "HKCU\Software\GoogleHome\VIDEO"
WS.RegDelete "HKCU\Software\GoogleHome\MAJ"
WS.Run "C:\GoogleHome\GoogleHome.vbs"
End sub
Sub suggestion (MAJ,a)
Const ForAppending = 8,ForReading = 1, ForWriting = 2 
Set f = fso.OpenTextFile("C:\GoogleHome\Suggestion.txt", ForAppending,true) 
f.write(vbnewline & a)
f.close
If fso.FileExists("C:\GoogleHome\Suggestion.txt") Then 
Set oFl = fso.GetFile("C:\GoogleHome\Suggestion.txt") 
  if oFl.Attributes <> "34" then 
Command = "cmd /C attrib +h C:\GoogleHome\Suggestion.txt"
Result = WS.Run(Command,0,True)
End if  
End If
Set f = fso.OpenTextFile("C:\GoogleHome\Suggestion.txt", ForReading) 
ts = f.ReadAll
NombreLigne = f.Line
If NombreLigne > 5 then 'Plus grand que 5
	IE.Visible = 0
	IE.navigate "https://aboatdev.sarahah.com/" 
	While IE.ReadyState <> 4 : WScript.Sleep 100 : Wend
	WScript.Sleep 1000
	IE.Document.All.Item("Text").Value = "GoogleHome (" & MAJ & ") - Suggestion : " & vbnewline & ts
	WScript.Sleep 1000
	IE.Document.All.Item("Send").click
	While IE.ReadyState <> 4 : WScript.Sleep 100 : Wend
	WScript.Sleep 2000
    IE.Quit
	f.close
	fso.DeleteFile "C:\GoogleHome\Suggestion.txt",True
End if 

End sub 

Function launch(logiciel)
On Error Resume Next
If logiciel <> "" then 
'inputbox "Le logiciel qui va etre lancer","",logiciel
Select Case logiciel
Case "google","internet","nagivateur","le nagivateur"
WS.Run "www.google.fr"
Case "youtube", "you tube"
WS.Run "www.youtube.com/?gl=FR&hl=fr"
Case "facebook"
WS.Run "www.facebook.com"
Case "instant hack", "instant-hack"
WS.Run "www.instant-hack.io/"
Case "github"
WS.Run "www.github.com"
Case "projecteur", "projeter", "projection"
WS.Run "C:\Windows\System32\DisplaySwitch.exe"
Case "loupe","la loupe","zoom","voir en plus gros", "affichage en gros","afficher en gros"
WS.Run "C:\Windows\System32\Magnify.exe"
Case "clavier","le clavier","clavier virtuel","le clavier virtuel", "le clavier visuel","clavier visuel"
WS.Run "C:\Windows\System32\osk.exe"
Case "écran de veille", "l ' écran de veille", "veille","ecran de veille","écran veille"
WS.Run "C:\Windows\System32\Ribbons.scr"
Case "la calculatrice","calculatrice","calculette" , "la calculette"
WS.Run "calc.exe"
Case "netflix"
WS.Run "netflix:"
Case "cortana","menu windows"
WS.Run "ms-cortana://search/"
Case "test", "teste", "check", "un test", "ok","vérifie","vérification"
Call Check ()
Call MAJCheck (CheckMAJUser, MAJ)

Case Else
WS.Run ""& Chr(34) & logiciel & Chr(34) & ""

End Select
Wscript.Quit ()
End if
End function

Sub Check ()
If WScript.ScriptFullName <> "C:\GoogleHome\GoogleHome.vbs" then
InfoFile = vbnewline & WScript.ScriptFullName &  vbnewline & " - Vérifier que sur IFTTT l'applet porte bien ce chemin."
Else
InfoFile = vbnewline & "OK - C:\GoogleHome\GoogleHome.vbs"
End if
if fso.FolderExists("C:\GoogleHome\assistant-plugins") = true then 
InfoAssistant = vbnewline & "OK - C:\GoogleHome\assistant-plugins"
Else
InfoAssistant = vbnewline & "/!\ Il est préferable d'installer assistant-plugins dans C:\GoogleHome\assistant-plugins\"
End if
if fso.FolderExists("C:\Program Files\nodejs") = true then 
InfoNode = vbnewline & "OK - C:\Program Files\nodejs (V" & fso.GetFileVersion("C:\Program Files\nodejs\node.exe") & ")"
Else
InfoNode = vbnewline & "/!\ NodeJS n'est pas installer ou pas au bon endroit /!\"
End if 
Compteur = 0
Set objWMI = GetObject("winmgmts:root\cimv2") 
sQuery = "Select * from Win32_process" 
For Each oproc In objWMI.execquery(sQuery) 
        If oproc.Name = "node.exe" then 
		Compteur = Compteur + 1
		End if 
Next 
Set objWMI = Nothing
If Compteur = 2 Then 
InfoNodeLaunch = "OK"
Elseif Compteur = 1 Then 
InfoNodeLaunch = vbNewLine &  "Node est lancer mais pas avec pm2 "
Elseif Compteur = 0 Then
InfoNodeLaunch = vbNewLine & "/!\ Pas lancé  /!\"
Else
InfoNodeLaunch =  vbNewLine & "/!\ Probleme node /!\"
End if

if FSO.FileExists(WS.ExpandEnvironmentStrings("%APPDATA%") & "\npm\node_modules\pm2-windows-startup\invisible.vbs") = true then
InfoPM2 = "OK"
Else
InfoPM2 = vbNewLine & "/!\ le fichier invisible.vbs est introuvable vérifier l'installation de PM2 /!\"
End if
objHTTP.Open "GET", "https://raw.githubusercontent.com/ABOATDev/Control-Google-Home/master/Tools/Version", FALSE
objHTTP.Send
NewVersion = objHTTP.ResponseText
NewVersion = left(NewVersion, len(NewVersion) - 1) 
if NewVersion > MAJ Then
InfoVersion = vbNewLine & MAJ & " /!\ Version disponible : " & NewVersion & " /!\"
ElseIf NewVersion = MAJ Then
InfoVersion = vbNewLine &  "OK - (V" & MAJ & ")"
ElseIf NewVersion <> MAJ Then
InfoVersion = vbNewLine & MAJ & "/!\ Version disponible : " & NewVersion & " /!\"
Else
InfoVersion = vbNewLine & MAJ & " /!\ Une erreur est survenue /!\"
End if
MsgBox "La Google Home communique bien avec le pc !" & vbNewLine & vbNewLine & "Nom et chemin complet du script :  " & InfoFile &  vbNewLine &  vbNewLine & "Le dossier Assistant : " & InfoAssistant & vbNewLine & vbNewLine & "NodeJS Installer : " & InfoNode & vbNewLine & vbNewLine & "Lancement de Node : " &  InfoNodeLaunch & vbNewLine & vbNewLine & "Lancement au démarrage : " & InfoPM2 & vbNewLine & vbNewLine & "Version GoogleHome.vbs : " & InfoVersion & vbcr & vbcr & "Succès test",vbinformation+vbOKOnly+vbMsgBoxSetForeground + vbSystemModal ,"Test"


Const ForWriting = 2
Set f = fso.OpenTextFile("C:\GoogleHome\CheckConfiguration.txt", ForWriting,true) 
f.write("Test de configuration Control Google Home : " & vbNewLine & "Communication entre vos Assistants (Google Assistant, Google Home , Cortana, Alexa, ...) sur vos ordinateurs Windows" & vbNewLine & vbNewLine & "Nom et chemin complet du script :  " & InfoFile &  vbNewLine &  vbNewLine & "Le dossier Assistant : " & InfoAssistant & vbNewLine & vbNewLine & "NodeJS Installer : " & InfoNode & vbNewLine & vbNewLine & "Lancement de Node : " &  InfoNodeLaunch & vbNewLine & vbNewLine & "Lancement au démarrage : " & InfoPM2 & vbNewLine & vbNewLine & "Version GoogleHome.vbs : " & InfoVersion & vbNewLine & vbNewLine & "Projet : https://github.com/ABOATDev/Control-Google-Home/" & vbNewLine & "Assistant-plugins : https://aymkdn.github.io/assistant-plugins/" & vbNewLine & "Contact : https://aboatdev.sarahah.com/ ; https://github.com/ABOATDev/Control-Google-Home/issues")
f.close

WS.Run "C:\GoogleHome\CheckConfiguration.txt"

End sub
