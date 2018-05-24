' Ok Google, sur le pc XXX
' Ok Google, sur l'ordinateur XXX
' Applet IFTTT : https://ifttt.com/applets/jSNrZ4vJ-controle-de-l-ordinateur-avec-google-assitant
' Projet  : https://github.com/ABOATDev/Control-Google-Home

Dim MAJ, WS,fso,CheckMAJUser,f,IE,objHTTP,ScriptChemin
MAJ = "1.0.9" 'Version Actuelle du script

On Error Resume Next


Set fso = CreateObject("Scripting.FileSystemObject")
Set WS = WScript.CreateObject("WScript.Shell") 
Set objHTTP=CreateObject("MSXML2.XMLHTTP")
Const ForWriting = 2
ScriptChemin = Left(WScript.ScriptFullName, InStr(WScript.ScriptFullName, WScript.ScriptName)-1)


if fso.FileExists(ScriptChemin & "Config.ini") = false then 
Set f = fso.OpenTextFile(ScriptChemin & "Config.ini", ForWriting,true) 
f.write(" ")
f.close
End if 

Set oFile = fso.GetFile(ScriptChemin & "Config.ini")

If WriteReadIni(oFile,"CONFIG","OK",Null) = False Then
WriteReadIni oFile,"CONFIG","OK","1"
Call MAJCheck (CheckMAJUser, MAJ)
objHTTP.Open "GET", "https://raw.githubusercontent.com/ABOATDev/Control-Google-Home/master/ListeCommande.txt", FALSE
objHTTP.Send
Set f = fso.OpenTextFile(ScriptChemin & "ListeCommande.txt", ForWriting,true) 
f.write(objHTTP.ResponseText)
f.close
MsgBox "Bienvenue dans mon script, il semblerait que vous lancer mon script pour la première fois ou que vous avez effectuer une mise à jour de celui-ci, pour faire fonctionner mon script dite : Ok Google, sur le pc xxx" & vbcr & "Par exemple Ok Google sur le pc test (pour tester la communication entre la Google homme est le PC)" & vbcr & " Dite des phrases simples et courtes" & vbcr & "Exercuté le script depuis l'ordinateur pour en savoir plus" & vbcr & vbcr & "Version Actuelle : " & MAJ ,vbInformation+vbOKOnly,"Control Google Home.vbs"
If WriteReadIni(oFile,"CONFIG","MUSIC",Null) = False Then
If MsgBox ("Voulez vous configuez le chemin d'accès pour la musiques ? " &vbcr & vbcr & "Sélectionner un dossier afin d'y rechercher des chansons dans ses sous-dossiers et ses sous-dossiers. Dossier par défaut" & vbcr & "Ok google sur le pc met de la musique" & vbcr & vbcr & "Si le dossier n'est pas configué, cela marchera quand même mais affichera un choix de dossier musique a chaque demande de musique" & vbcr & vbcr & "Oui = Configuer",vbyesno,"Configurez le dossier Musique") = vbYes Then
Dim objShell,objFolder,Message
	Message = "Veuillez sélectionner un dossier afin d'y rechercher des chansons dans ses sous-dossiers et ses sous-dossiers."
	Set objShell = CreateObject("Shell.Application")
	Set objFolder = objShell.BrowseForFolder(0,Message,1)
	If objFolder Is Nothing Then Wscript.Quit
	WriteReadIni oFile,"CONFIG","MUSIC",objFolder.self.path
	MsgBox "Je conseil de tester la commande <musique> pour vérifier que tout fonctionne bien et que le lecteur media est compatible",vbInformation+vbOKOnly,"Ok"
End if
End if 
If WriteReadIni(oFile,"CONFIG","VIDEO",Null) = False Then
If MsgBox ("Voulez vous configuez le chemin d'accès pour les vidéos ? " &vbcr & vbcr & "Sélectionner un dossier afin d'y rechercher des chansons dans ses sous-dossiers et ses sous-dossiers. Dossier par défaut" & vbcr & "Ok google sur le pc met de les vidéos" & vbcr & vbcr & "Si le dossier n'est pas configué, cela marchera quand même mais affichera un choix de dossier vidéos a chaque demande de musique" & vbcr & vbcr & "Oui = Configuer",vbyesno,"Configurez le dossier Vidéo") = vbYes Then
	Message = "Veuillez sélectionner un dossier afin d'y rechercher des videos dans ses sous-dossiers et ses sous-dossiers."
	Set objShell = CreateObject("Shell.Application")
	Set objFolder = objShell.BrowseForFolder(0,Message,1)
	If objFolder Is Nothing Then Wscript.Quit
	WriteReadIni oFile,"CONFIG","VIDEO",objFolder.self.path
	MsgBox "Je conseil de tester la commande <vidéo> pour vérifier que tout fonctionne bien et que le lecteur media est compatible",vbInformation+vbOKOnly,"Ok"
End if
End if 
End if

Set objArgs = WScript.Arguments
For I = 0 to objArgs.Count -1
Select Case objArgs(I)
Case "écris", "écrit","marque"
ecrit = true
Case "lance", "ouvre","affiche","démarre", "exécute","ouvrir","démarrer","exécuter","lancer"
lance = true
Case "message","messagebox"
message = true
Case Else
a = a & " " & LCase(objArgs(I))
End Select
Next

If ecrit = true then Call write(a)
If message = true then Call MsgBoxtexte(a)
If lance = true then Call launch (right (a,len(a)-1)) '(Logiciel)
'inputbox a,a,a

If a = "" then
Call MAJCheck (CheckMAJUser, MAJ)

rep = InputBox ("Bienvenue dans mon script, communication entre vos Assistants (Google Assistant, Google Home , Cortana, Alexa, ...) sur vos ordinateurs Windows" & vbNewLine &  "Pour faire fonctionner mon script dite : Ok Google, sur le pc xxx" & vbcr & "Par exemple Ok Google sur le pc test (pour tester la communication entre la Google homme est le PC)" & vbcr & vbcr & " Dite des phrases simples et courtes" & vbcr & vbcr & vbcr & "1 = Vérifier mise a jours" & vbcr & "2 = Envoyé un messsage au créateur (rapide & sans se logger)" & vbcr & "3 = Réinsalisé la configuration du script." & vbCr & "4 = Crédit" & vbcr & "5 = Rajouter un logiciel a la liste" & vbCr & vbCr & "Pour tester des commandes en écrit, il vous suffit de taper une commande si dessous pour savoir si elle est comprise par le logiciel" & vbNewLine & "Version : " &  MAJ,"Control Google Home " & MAJ,"test")
   If rep = "" then
   WScript.Quit()
   ElseIf rep = "1" then 
   CheckMAJUser = true
   Call MAJCheck (CheckMAJUser, MAJ)
   Wscript.Quit
   ElseIf rep = "2" then 
   WS.Run "https://aboatdev.sarahah.com/" 
   Wscript.Quit
   ElseIf rep = "3" then 
   Reset ()
   Wscript.Quit
   ElseIf rep = "4" then
   MsgBox "Crédits : " & vbNewLine & vbNewLine & "HackooFr - Aide indirect pour le Script" & vbNewLine & "facebook.com/hackoo.crackoo" & vbNewLine & vbNewLine & "Aymkdn - Pour l'assistant-plugins" &  vbNewLine & " github.com/Aymkdn | paypal.me/aymkdn" & vbNewLine & vbNewLine & "Créateur du Contrôle de l'ordinateur avec Google Home : ABOAT " & vbNewLine & "facebook.com/aboat.hack",vbInformation+vbOKOnly,"Crédits"
   Wscript.Quit
   ElseIf rep = "5" then 
   nomfile = Inputbox ("Le nom du fichier a ouvrir ?" & vbcr & "Le nom que vous direz vocalement a votre assistant vocal" & vbCr & "Ne pas mettre de majuscule !","Nom du fichier Pages 1/2")
   cheminfile = Inputbox ("Le chemin complet du fichier " & nomfile & vbcr, "Chemin de : " & nomfile & "Pages 2/2")
     WriteReadIni oFile,"Logiciel",nomfile,cheminfile
	 If fso.FileExists(cheminfile) = true Then MsgBox "Le logiciel " & nomfile & " rajouter !",vbOKOnly+vbInformation,"Fichier rajouté !" 
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
Case "augmente le son","augmente le volume","monte le son","news le son","mieux que le son" : WS.SendKeys "{" & chr(175) & " 10}"
Case "monte le son au max","monte le son au maximum","monte le volume au maximum","volume max","volume maximum","son au max","augmente le son au maximum","mais le son au max","mais le son au maximum","mais le volume au max","mais le volume au maximum","mets le son à fond","le son à fond","son à fond" : WS.SendKeys "{" & chr(175) & " 50}"
Case "baisse le son","descend le son","descend le volume","baisse le volume" : WS.SendKeys "{" & chr(174) & " 10}"
Case "descend le son au max","baisse le son au max","baisse le volume au max","baisse le son au maximum","volume minimum","volume au minimum","baisse le volume au maximum" : WS.SendKeys "{" & chr(174) & " 50}"
Case "mute","mute le volume","mute le son","muet","le son a 0","coupe le son","coupe le volume","coupe l'audio","remets le volume","remets le son","remets le son","arrête le son","stop le son","stop le v","désactive le son","désactive le volume","allume le son","éteint le son","allume le volume","éteint le volume" : WS.SendKeys chr(173)
Case "pause","fait pause","met pause","mais pause","fais une pause","met en pause","mais en pause","fait pause","fait stop","stop","pause","relance","enlève la pause","met une pause","lance","lecture","mais play","lance lecture","lance la lecture","mais en pause","lecture","mais plaît","se pose" : WS.SendKeys " "
Case "éteint le","arrête le","éteint le pc","éteint l'ordinateur","arrête le pc","éteint l ' ordinateur","arrête le système","éteint le système"," arrête","arrête l ' ordinateur","arrêter le système","éteint","éteint le","arrêt du système" : CreateObject("Wscript.Shell").Run "CMD /C " & " shutdown /s /f /t 01",0
Case "verrouille le","verrouiller le","verrouille la session","verrouiller la session","verrouille le pc","le verrouiller","met en veille","mettre en veille","met le en veille","veille","verrouillage","verrouille","metre en veille","verrouiller la session","verrouille la session","mais en veille","verrouiller","verrouillé","verrouiller le pc" : WS.Run "rundll32.exe user32.dll,LockWorkStation"
Case "mot de passe wifi","mot de passe du wifi","code wifi","wifi","code de la wifi","donne mot de passe wifi","code du wifi","donne le mot de passe wifi","donne le mot de passe du wifi","retrouve le mot de passe wifi","retrouve le mot de passe du wifi","quel est le mot de passe wifi","quel est le mot de passe du wifi","donne le mot de passe" : Call TelechargerTools ("WifiPasswordsRecovery.bat","https://raw.githubusercontent.com/ABOATDev/Control-Google-Home/master/Tools/WifiPasswordsRecovery.bat")
Case "ejecte le cd","eject cd","eject le dvd","eject cd","eject dvd","ejecter dvd","ejecter cd"," ejecter le dvd","eject dvd" : LecteurDVD ()
Case "ferme le logiciel","ferme le logiciel actif","arrête l ' application","arrête le logiciel","arrête l ' application","ferme l ' application","ferme le programme","arrête le programme","quitte le programme" : WS.SendKeys ("%{F4}")
Case "eject usb", "eject clé usb", "eject la clé usb" , "retire usb" , "retire la clé usb","retire clé usb" : Call TelechargerTools ("Eject_USB.vbs","https://raw.githubusercontent.com/ABOATDev/Control-Google-Home/master/Tools/Eject_USB.vbs")
Case "écran de veille", "l ' écran de veille", "veille","ecran de veille","écran veille", "met l ' écran de veille","mais l ' écran de veille" : WS.Run "C:\Windows\System32\Ribbons.scr"
case "liste des commandes", "liste commande", "donne la liste des commandes" , "détail des commandes", "les commandes disponible", "liste des commandes disponible" : Call TelechargerTools ("ListeCommande.txt","https://raw.githubusercontent.com/ABOATDev/Control-Google-Home/master/ListeCommande.txt")
Case "spotify","la lecture spotify","lecture spotify","musique spotify","la musique spotify","spotify musique","spotify lecture" : Call TelechargerTools ("LectureSpotify.vbs","https://raw.githubusercontent.com/ABOATDev/Control-Google-Home/master/Tools/LectureSpotify.vbs")
Case "musique","met de la musique","lance de la musique","mais de la musique","lance musique","audio","met la musique","met la playlist","lance la playlist","met la playlist" : Call TelechargerTools ("LancerDossierMusique.vbs","https://raw.githubusercontent.com/ABOATDev/Control-Google-Home/master/Tools/LancerDossierMusique.vbs")
Case "vidéo","film","met vidéo","film","mais vidéo","lance vidéo","lance film","met les vidéos","met la vidéo","lance la vidéo","met le film","met les films","lance la vidéo","met la vidéo" : Call TelechargerTools ("LancerDossierVideo.vbs","https://raw.githubusercontent.com/ABOATDev/Control-Google-Home/master/Tools/LancerDossierVideo.vbs")

Case "maj","mise à jour","vérifier mise à jour","vérifie mise à jour","mise à jour script","vérifier","mage"
CheckMAJUser = true
Call MAJCheck (CheckMAJUser, MAJ)

Case Else

Call  MAJCheck (CheckMAJUser, MAJ)
Call Suggestion (MAJ,a)
'Inputbox "La valeur n'existe pas","Erreur : valeur n'existe pas",a
End Select


Function launch(logiciel)
On Error Resume Next
If logiciel <> "" then 
'inputbox "Le logiciel qui va etre lancer","",logiciel
Select Case logiciel
Case "google","internet","nagivateur","le nagivateur" : WS.Run "www.google.fr"
Case "youtube", "you tube" : WS.Run "www.youtube.com/?gl=FR&hl=fr"
Case "facebook" : WS.Run "www.facebook.com"
Case "instant hack", "instant-hack" : WS.Run "www.instant-hack.io/"
Case "github" : WS.Run "www.github.com"
Case "projecteur", "projeter", "projection","le projecteur" : WS.Run "C:\Windows\System32\DisplaySwitch.exe"
Case "se connecter", "connection", "connection","connexion","connexion sans fil" : WS.Run "ms-projection:"
Case "loupe","la loupe","zoom","voir en plus gros", "affichage en gros","afficher en gros" : WS.Run "C:\Windows\System32\Magnify.exe"
Case "clavier","le clavier","clavier virtuel","le clavier virtuel", "le clavier visuel","clavier visuel" : WS.Run "C:\Windows\System32\osk.exe"
Case "écran de veille", "l ' écran de veille", "veille","ecran de veille","écran veille" : WS.Run "C:\Windows\System32\Ribbons.scr"
Case "la calculatrice","calculatrice","calculette" , "la calculette" : WS.Run "calc.exe"
Case "netflix" : WS.Run "netflix:"
Case "spotify","la lecture spotify","lecture spotify","musique spotify","la musique spotify","spotify musique","spotify lecture" : Call TelechargerTools ("LectureSpotify.vbs","https://raw.githubusercontent.com/ABOATDev/Control-Google-Home/master/Tools/LectureSpotify.vbs")
Case "cortana","menu windows" : WS.Run "ms-cortana://search/"
Case "le lecteur cd","le lecteur cd","lecteur","le lecteur cd","le lecteur dvd","lecteur dvd","lecteur cd" : LecteurDVD ()
Case "bureau","desktop","bureaux","le bureau" : CreateObject("Shell.Application").ToggleDesktop
Case "test", "teste", "check", "un test", "ok","vérifie","vérification"
Call Check ()
Call MAJCheck (CheckMAJUser, MAJ)
Case Else
'Msgbox WriteReadIni(oFile,"Logiciel",logiciel,Null)
If WriteReadIni(oFile,"Logiciel",logiciel,Null) <> False then 
WS.Run ""& Chr(34) & WriteReadIni(oFile,"Logiciel",logiciel,Null) & Chr(34) & ""
else
WS.Run ""& Chr(34) & logiciel & Chr(34) & ""
End if
End Select
Wscript.Quit ()
End if
End function

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

Sub write(a)
WScript.Sleep 300
WS.SendKeys right(a,len(a)-1)
WScript.Quit ()
End sub

Sub MsgBoxtexte(a)
MsgBox "Message reçus de votre assistant vocal à " & Hour(Now)& ":"& Minute(Now) & vbnewline & vbnewline &  a,vbinformation+vbOKOnly, Hour(Now)& ":"& Minute(Now)
WScript.Quit ()
End sub

Sub Reset ()
On Error Resume Next
If fso.FileExists(ScriptChemin & "Config.ini") = true then
fso.DeleteFile ScriptChemin & "Config.ini",True
WS.Run "cmd /k chcp 28591 > nul & taskkill /F /IM wscript.exe & start " & ScriptChemin & WScript.ScriptName & " & exit",0,true
Else
MsgBox "Le fichier Config.ini n'a pas pu être supprimé.",vbCritical+vbOKOnly,"Reset non effectué"
End if
End sub


Sub TelechargerTools (NomFile,URL)
'Call TelechargerTools (NomFile,URL)
If FSO.FolderExists(ScriptChemin & "Tools") = false Then FSO.CreateFolder (ScriptChemin & "Tools")
If FSO.FileExists(ScriptChemin & "Tools\" & NomFile) = false then 
 objHTTP.Open "GET", URL, FALSE
 objHTTP.Send
 Telecharger = objHTTP.ResponseText
 Set f = fso.OpenTextFile(ScriptChemin & "Tools\" & NomFile, ForWriting,true) 
 f.write(Telecharger)
 f.close
 WScript.Sleep 100
End if
	WS.Run ScriptChemin & "Tools\" & NomFile
End sub 


Sub MAJCheck (CheckMAJUser, MAJ)
'On Error Resume Next
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
     Set f = fso.OpenTextFile(ScriptChemin & "GoogleHomeNew.txt", ForWriting,true) 
     f.write(Telecharger)
     f.close
	 CheckMAJUser = false
     Return = WS.Run ("cmd /k chcp 28591 > nul & taskkill /F /IM wscript.exe & move " & ScriptChemin & "GoogleHomeNew.txt " & ScriptChemin & WScript.ScriptName & " & start " & ScriptChemin & WScript.ScriptName & " & exit",0,true)
	Else
	 If CheckMAJUser = true then MsgBox "Pas de nouvelle mise a jours a installer" & vbNewLine & "Vous êtes bien dans la derniere version disponible" & vbNewLine & vbNewLine & vbNewLine & "Votre version : " & VersionActu & vbNewLine & "Derniere version : " & NewVersion
	 CheckMAJUser = false
End if
End sub

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
MsgBox "Votre assistant vocal semple bien communiquer bien avec l'ordinateur ! (si vous avez configurez WEBHOOKS votre assistant vocal devrais faire un retour vocal dans quelque instant) " & vbNewLine & vbNewLine & "Nom et chemin complet du script :  " & InfoFile &  vbNewLine &  vbNewLine & "Le dossier Assistant : " & InfoAssistant & vbNewLine & vbNewLine & "NodeJS Installer : " & InfoNode & vbNewLine & vbNewLine & "Lancement de Node : " &  InfoNodeLaunch & vbNewLine & vbNewLine & "Lancement au démarrage : " & InfoPM2 & vbNewLine & vbNewLine & "Version GoogleHome.vbs : " & InfoVersion & vbcr & vbcr & "Succès test",vbinformation+vbOKOnly+vbMsgBoxSetForeground + vbSystemModal ,"Test"
Const ForWriting = 2
Set f = fso.OpenTextFile(ScriptChemin & "CheckConfiguration.txt", ForWriting,true) 
f.write("Test de configuration Control Google Home : " & vbNewLine & "Communication entre vos Assistants (Google Assistant, Google Home , Cortana, Alexa, ...) sur vos ordinateurs Windows" & vbNewLine & vbNewLine & "Nom et chemin complet du script :  " & InfoFile &  vbNewLine &  vbNewLine & "Le dossier Assistant : " & InfoAssistant & vbNewLine & vbNewLine & "NodeJS Installer : " & InfoNode & vbNewLine & vbNewLine & "Lancement de Node : " &  InfoNodeLaunch & vbNewLine & vbNewLine & "Lancement au démarrage : " & InfoPM2 & vbNewLine & vbNewLine & "Version GoogleHome.vbs : " & InfoVersion & vbNewLine & vbNewLine & "Projet : https://github.com/ABOATDev/Control-Google-Home/" & vbNewLine & "Assistant-plugins : https://aymkdn.github.io/assistant-plugins/" & vbNewLine & "Contact : https://aboatdev.sarahah.com/ ; https://github.com/ABOATDev/Control-Google-Home/issues")
f.close
WS.Run ScriptChemin & "CheckConfiguration.txt"
End sub

Sub suggestion (MAJ,a)
On Error Resume Next
Set IE = Wscript.CreateObject("InternetExplorer.Application")
Const ForAppending = 8,ForReading = 1, ForWriting = 2 
Set f = fso.OpenTextFile(ScriptChemin & "Suggestion.txt", ForAppending,true) 
f.write(vbnewline & a)
f.close
If fso.FileExists(ScriptChemin & "Suggestion.txt") Then 
Set oFl = fso.GetFile(ScriptChemin & "Suggestion.txt") 
  if oFl.Attributes <> "34" then 
Command = "cmd /C attrib +h " & ScriptChemin & "Suggestion.txt"
Result = WS.Run(Command,0,True)
End if  
End If
Set f = fso.OpenTextFile(ScriptChemin & "Suggestion.txt", ForReading) 
ts = f.ReadAll
NombreLigne = f.Line
If NombreLigne > 7 then 'Plus grand que 5
	IE.Visible = 0
	IE.navigate "https://aboatdev.sarahah.com/" 
	While IE.ReadyState <> 4 : WScript.Sleep 100 : Wend
	WScript.Sleep 1000
	IE.Document.All.Item("Text").Value = "GoogleHome (" & MAJ & ") - Suggestion : " & vbnewline & ts & vbcr & "Suggestion auto par : "  & CreateObject("WScript.Network").username
	WScript.Sleep 1000
	IE.Document.All.Item("Send").click
	While IE.ReadyState <> 4 : WScript.Sleep 100 : Wend
	WScript.Sleep 2000
    IE.Quit
	f.close
	fso.DeleteFile ScriptChemin & "Suggestion.txt",True
	
		strComputer = "." 
Set objWMIService = GetObject("winmgmts:" _ 
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
Set colProcessList = objWMIService.ExecQuery _ 
    ("Select * from Win32_Process Where Name = 'ielowutil.exe'") 
  
For Each objProcess in colProcessList 
    objProcess.Terminate() 
Next
Set objWMIService2 = GetObject("winmgmts:" _ 
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
Set colProcessList2 = objWMIService.ExecQuery _ 
    ("Select * from Win32_Process Where Name = 'iexplore.exe'") 

For Each objProcess2 in colProcessList2
    objProcess2.Terminate() 
Next
End if	
End sub 

 Function WriteReadIni(oFile,section,key,value)
' *******************************************************************************************
' omen999 - mars 2018 v 1.1 - http://omen999.developpez.com/
' ********************************************************************************************
Dim oText,iniText,sectText,newSectText,keyText
  Set reg = New RegExp
  Set regSub = New RegExp
  reg.MultiLine=True
  reg.IgnoreCase = True
  regSub.IgnoreCase = True
  Set oText = oFile.OpenAsTextStream(1,0)
  iniText = oText.ReadAll
  oText.Close
  reg.Pattern = "^\[" & section & "\]((.|\n[^\[])+)":regSub.Pattern = "\b" & key & " *= *([^;\f\n\r\t\v]*)"
  On Error Resume Next
  If IsNull(value) Then
    WriteReadIni = regSub.Execute(reg.Execute(iniText).Item(0).SubMatches(0)).Item(0).SubMatches(0)
    If Err.Number = 5 then WriteReadIni = False
  Else
    sectText = reg.Execute(iniText).Item(0).SubMatches(0)
    If Err.Number = 5 Then
      iniText = iniText & vbCrLf & "[" & section & "]" & vbCrLf & key & "=" & value
    Else
      newSectText = regSub.Replace(sectText,key & "=" & value)
      If newSectText = sectText Then
        If regSub.Test(sectText) Then
          WriteReadIni = False
          Exit Function
        End If
        If Right(sectText,1) = vbCr Then keyText = key & "=" & value Else keyText = vbCrLf & key & "=" & value
        newSectText = sectText & keyText
      End If
      iniText = reg.Replace(iniText,"[" & section & "]" & newSectText)
    End If
    Set oText = oFile.OpenAsTextStream(2,0)
    oText.Write iniText
    oText.Close
    WriteReadIni = True
  End If
End Function
