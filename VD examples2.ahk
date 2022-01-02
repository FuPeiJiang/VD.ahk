;you should first Run this, then Read this
;Ctrl + F: jump to #useful stuff

;#SETUP START
#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance force
ListLines Off
SetBatchLines -1
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
#KeyHistory 0
#WinActivateForce

Process, Priority,, H

SetWinDelay -1
SetControlDelay -1


;include the library
#Include VD.ahk

VD.init()

; MsgBox % VD._strGUID_from_Hwnd(WinExist("A"))

; MsgBox % VD.getCount()

return

numpad1::VD.goToDesktop(1)
numpad2::VD.goToDesktop(2)
numpad3::VD.goToDesktop(3)

f3::Exitapp
