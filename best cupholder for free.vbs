Option Explicit

Dim objShell, intResponse, objWMIService, colDrives, objDrive, hasCDDrive

Set objShell = CreateObject("WScript.Shell")
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colDrives = objWMIService.ExecQuery("Select * from Win32_CDROMDrive")

hasCDDrive = False

For Each objDrive in colDrives
    hasCDDrive = True
    Exit For
Next

intResponse = objShell.Popup("Do you want a free cupholder?", 0, "Totally legit free cupholder", 4 + 32)

If intResponse = 6 Then ' Yes
    If hasCDDrive Then
        objShell.Run "wmic cdrom where driveletter!=null call eject"
    Else
        objShell.Popup "No free cupholder for you lil bro.", 0, "Best cupholder for free!"
    End If
ElseIf intResponse = 7 Then ' No
    objShell.Popup "Ok lil bro", 0, "Fuck off lil bro"
End If

