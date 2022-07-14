#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance force
ListLines Off
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
SetBatchLines -1
#KeyHistory 0

#Include ..\VD.ahk

VD.startShellMessage()
VD.Run_lock_VD("""C:\Program Files (x86)\Hourglass\Hourglass.exe""","","","","Hourglass.exe",2,3)
; VD.Run_lock_VD("""C:\Program Files (x86)\Hourglass\Hourglass.exe""","","","","Hourglass.exe",2,VD.getCurrentDesktopNum()) ;I like to be explicit


return

f3::Exitapp