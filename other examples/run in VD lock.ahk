#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance force
ListLines Off
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
SetBatchLines -1
#KeyHistory 0

#Include %A_LineFile%\..\..\VD.ahk

VD.startShellMessage()
VD.goToDesktopNum(1)
;assume no OneNote window exists
;first window needs 5 seconds to ready
Run % "shell:AppsFolder\Microsoft.Office.OneNote_8wekyb3d8bbwe!microsoft.onenoteim"
Sleep 5000 ;subsequent runs will be ignored without waiting 5000ms
VD.Run("shell:AppsFolder\Microsoft.Office.OneNote_8wekyb3d8bbwe!microsoft.onenoteim","","OneNote for Windows 10","ApplicationFrameWindow","ApplicationFrameHost.exe",2)
VD.Run("shell:AppsFolder\Microsoft.Office.OneNote_8wekyb3d8bbwe!microsoft.onenoteim","","OneNote for Windows 10","ApplicationFrameWindow","ApplicationFrameHost.exe",3)

return

f3::Exitapp