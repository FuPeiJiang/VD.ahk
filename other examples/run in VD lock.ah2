#SingleInstance force
ListLines 0
SendMode "Input"
SetWorkingDir A_ScriptDir
KeyHistory 0

#Include %A_LineFile%\..\..\VD.ah2

VD.startShellMessage()
VD.goToDesktopNum(1)

Run "shell:AppsFolder\Microsoft.Office.OneNote_8wekyb3d8bbwe!microsoft.onenoteim"
Sleep 5000
VD.Run("shell:AppsFolder\Microsoft.Office.OneNote_8wekyb3d8bbwe!microsoft.onenoteim", "", "OneNote for Windows 10", "ApplicationFrameWindow", "ApplicationFrameHost.exe", 2)
VD.Run("shell:AppsFolder\Microsoft.Office.OneNote_8wekyb3d8bbwe!microsoft.onenoteim", "", "OneNote for Windows 10", "ApplicationFrameWindow", "ApplicationFrameHost.exe", 3)

return

f3::Exitapp
