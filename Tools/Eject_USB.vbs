'----------------------------Eject_USB.vbs---------------------------------
'Modification du code NumSerie_Usb.vbs by © Hackoo pour Eject_USB.vbs
'Tester et vérifier si votre clé USB est connecté ou non, 
'si cette dernière est connectée alors le script l'eject.
' ABOAT
'--------------------------------------------------------------------------

'----------------------------Original : NumSerie_Usb.vbs-------------------
'Tester et vérifier si votre clé USB est connecté ou non, 
'si cette dernière est connectée alors le script nous donne son N° de Série.
'© Hackoo
'--------------------------------------------------------------------------

  Sub Eject_Usb()
  On Error Resume Next
  Dim NumSerie, Patch
  Set fso = CreateObject("Scripting.FileSystemObject")
  For Each Drive In fso.Drives
  If Drive.IsReady Then
  If Drive.DriveType=1 Then
  Patch = fso.Drives(Drive + "\").Path
  'WScript.CreateObject("WScript.Shell").Popup "La clé USB : {" & fso.Drives(Drive + "\").VolumeName & " (" & Patch & ")} va être ejecté.",2,"Ejection " & fso.Drives(Drive + "\").VolumeName & " (" & Patch & ") ",vbOKOnly+vbInformation
  CreateObject("Shell.Application").NameSpace(fso.Drives(Drive + "\").Path).Self.InvokeVerb "Eject"
  If err.Number<>0 Then 
  WScript.CreateObject("WScript.Shell").Popup "La clé USB :  " & vbNewLine & fso.Drives(Drive + "\").VolumeName & " (" & Patch & ") n'a pas etait ejecté ! ",5,"Ejection " & fso.Drives(Drive + "\").VolumeName & " (" & Patch & ") ",vbOKOnly+vbCritical
    Else
  WScript.CreateObject("WScript.Shell").Popup "La clé USB : " & vbNewLine & fso.Drives(Drive + "\").VolumeName & " (" & Patch & ")} Ejecté !" & vbNewLine & vbNewLine & "Vous pouvez la retirer en toute sécurité.",3,"Ejection " & fso.Drives(Drive + "\").VolumeName & " (" & Patch & ") ",vbOKOnly+vbInformation
  End if 
  End if
  End If
  Next
  End Sub
 
'------------------------------checkUSB----------------------------
Sub checkUSB
strComputer = "."
On Error Resume Next
Set WshShell = CreateObject("Wscript.Shell")
beep = chr(007)
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_DiskDrive WHERE InterfaceType='USB'",,48)
intCount = 0
For Each drive In colItems
    If drive.mediaType <> "" Then
        intCount = intCount + 1
    End If
Next
If intCount > 0 Then
   ' MsgBox "Votre Clé USB Personnelle est bien Connectée !",64,"Flash Drive Check © Hackoo!"
	Call Eject_Usb() ' Appelle a la procédure Eject_Usb()
Else
WScript.CreateObject("WScript.Shell").Popup "Aucune Clé USB est détecté",3,"Flash Drive Check",vbOKOnly+vbCritical
End If
End Sub
'---------------------------Fin du checkUSB----------------------------
Call checkUSB ' Appelle a la procédure checkUSB
