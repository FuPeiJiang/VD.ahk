#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance, force
ListLines Off
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
SetBatchLines, -1
#KeyHistory 0

#Include %A_LineFile%\..\..\VD.ahk

VD_Init()
; VD_AppIdFromHwnd(WinExist("A"))
; VD_AppIdFromView(VD_ViewFromHwnd("A"))
VD_IsAppPinned("A")

return

f3::Exitapp