#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance force
ListLines Off
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetBatchLines -1
#KeyHistory 0

#Include %A_LineFile%\..\..\VD.ahk
VD.RegisterDesktopNotifications()
VD.CurrentVirtualDesktopChanged:=Func("CurrentVirtualDesktopChanged")
CurrentVirtualDesktopChanged(desktopNum_Old, desktopNum_New) {
    ToolTip % desktopNum_Old ", " desktopNum_New
}
return

f3::Exitapp