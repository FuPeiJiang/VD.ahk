#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance, force
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
SetBatchLines, -1
#KeyHistory 0
ListLines Off

DetectHiddenWindows, off
SetWinDelay, -1

Gui,Font, s12, Segoe UI
gui, add, text,, Press ^d to come back to this window`nPress ^f to see which desktop this window is in`nPress f1 to see your current virtual desktop`nPress Numpad8 to move this to Desktop2`nPress Numpad2 to go to Desktop2`nPress Numpad6 to move this to Desktop3 and go to Desktop 3 (follow the window)
gui, show,, VD_examplesWinTile

;include the library
#Include VD.ahk
vd_init() ;call this when you want to init global vars, takes 0.04 seconds for me.
return

;getters and stuff
^d::
    VD_goToDesktopOfWindow("VD_examplesWinTile")
    ; VD_goToDesktopOfWindow("ahk_exe code.exe")
return
^f::
    msgbox % VD_getDesktopOfWindow("VD_examplesWinTile")
    ; msgbox % VD_getDesktopOfWindow("ahk_exe GitHubDesktop.exe")
return
f1::
    msgbox % VD_getCurrentDesktop()
return
f2::
    msgbox % VD_getCount()
return

;useful stuff
pleaseSwitchDesktop:
    VD_goToDesktop(theDesktopToSwitchTo) ;get last character from numpad{N}
return
numpad1::
numpad2::
numpad3::
    WindowisFullScreen:=VD_isWindowFullScreen("A")  ;"A" specially means active window
    theDesktopToSwitchTo:=SubStr(A_ThisHotkey, 0)
    VD_goToDesktop(theDesktopToSwitchTo) ;get last character from numpad{N}
    if (WindowisFullScreen)
        SetTimer, pleaseSwitchDesktop, -50
return

numpad4::
numpad5::
numpad6::
    wintitleOfActiveWindow:="ahk_id " WinActive("A")
    whichDesktop:=SubStr(A_ThisHotkey, 0) - 3
    VD_sendToDesktop(wintitleOfActiveWindow,whichDesktop,true) ;get last character from numpad{N} and minus 6
return
numpad7::
numpad8::
numpad9::
    wintitleOfActiveWindow:="ahk_id " WinActive("A")
    whichDesktop:=SubStr(A_ThisHotkey, 0) - 6 ;get last character from numpad{N} and minus 3
    VD_sendToDesktop(wintitleOfActiveWindow,whichDesktop) 
return

f3::Exitapp