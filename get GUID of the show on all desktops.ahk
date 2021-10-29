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

Gui,Font, s12, Segoe UI

explanation=
(
^ is Ctrl
f4 to come back to this window
f6 to see which desktop this window is in
f1 to see your current virtual desktop
f2 to see the total number of virtual desktops
Numpad8 to move the active window to Desktop2
Numpad2 to go to Desktop2
Numpad6 to move the active window to Desktop3 and go to Desktop 3 (follow the window)

Numpad0 to Toggle "Show this window on all desktops"(Pin/UnPin)
NumpadDot to Toggle "Show windows from this app on all desktops"(Pin/UnPin)
)

defaultBackgroundColor:=get_DefaultBackgroundColor()
; p(defaultBackgroundColor)

; gui, add, Text,, % explanation
; https://www.autohotkey.com/boards/viewtopic.php?t=3956#p21359
gui, add, Edit, -vscroll -E0x200 +hwndHWndExplanation_Edit, % explanation
; set background color of Edit
; CtlColors.Attach(HWndExplanation_Edit, "red")
; CtlColors.Attach(HWndExplanation_Edit, hex_RGBtoBGR(decimalToHex(16776960)))
CtlColors.Attach(HWndExplanation_Edit, hex_RGBtoBGR(decimalToHex(0+defaultBackgroundColor)))

Gui, Color, 
gui, show,, VD_examplesWinTile
WinSet, Redraw
; Postmessage,0xB1,-1,-1,, % "ahk_id " HWndExplanation_Edit ;move caret to end
Postmessage,0xB1,9,9,, % "ahk_id " HWndExplanation_Edit ;move caret to: ^ is Ctrl|


;include the library
#Include VD.ahk
vd_init() ;call this when you want to init global vars, takes 0.04 seconds for me.

;you should WinHide invisible programs that have a window.
WinHide, % "Malwarebytes Tray Application"

;autoexecute section start

VD_UnPinWindow("VD_examplesWinTile")
VD_UnPinApp("VD_examplesWinTile")
; VD_PinApp("VD_examplesWinTile")
msgbox % VD_getDesktopOfWindow("VD_examplesWinTile")
; ALL {BB64D5B7-4DE3-4AB2-A87C-DB7601AEA7DC}

; DesktopManager.VirtualDesktopManagerInternal.GetAdjacentDesktop(ivd, 3, out fallbackdesktop); // 3 = LeftDirection

CurrentIVirtualDesktop:=VD_getCurrentIVirtualDesktop()
d(CurrentIVirtualDesktop)
; 10378552
; 10378776

; 9924392
; 9925624

DllCall(GetAdjacentDesktop, "UPtr", IVirtualDesktopManagerInternal, "Ptr", CurrentIVirtualDesktop, "UInt", 4, "Ptr*", rightDesktop)
d(rightDesktop)

DllCall(GetAdjacentDesktop, "UPtr", IVirtualDesktopManagerInternal, "Ptr", CurrentIVirtualDesktop, "UInt", 3, "Ptr*", leftDesktop)
d(leftDesktop)
; 1345048
; 1344936
; 1346504
; so it's dynamically allocating memory
; so it's newly allocated memory
; no wait, RIGHT, is SMALLER than center. so it is NOT.
; 1345048 - 1344936 = 112
; 1346504 - 1345048 = 1456 (maybe because left has more windows)

; 1205032
; 1205704
; 1206152

; 1205032 - 1205704 = -672, now right is bigger, theory disproven


return

;getters and stuff
f4::
    VD_goToDesktopOfWindow("VD_examplesWinTile")
    ; VD_goToDesktopOfWindow("ahk_exe code.exe")
return
f6::
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
    WindowisFullScreen:=VD_isWindowFullScreen("A") ;"A" specially means active window
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

;Pin Window
numpad0::
    VD_TogglePinWindow("A")
return
^numpad0::
    VD_PinWindow("A")
return
!numpad0::
    VD_UnPinWindow("A")
return
#numpad0::
    MsgBox % VD_IsWindowPinned("A")
return

;Pin App
numpadDot::
    VD_TogglePinApp("A")
return
^numpadDot::
    VD_PinApp("A")
return
!numpadDot::
    VD_UnPinApp("A")
return
#numpadDot::
    MsgBox % VD_IsAppPinned("A")
return

f3::Exitapp

hex_RGBtoBGR(str) {
    R:=SubStr(str, 1, 2)
    G:=SubStr(str, 3, 2)
    B:=SubStr(str, 5, 2)
    return B G R
}

decimalToHex(num) {
    static map1 := {0:"0",1:"1",2:"2",3:"3":"4":"4",5:"5",6:"6",7:"7",8:"8",9:"9",10:"A",11:"B",12:"C",13:"D",14:"E",15:"F"}
    ; https://www.rapidtables.com/convert/number/decimal-to-hex.html
    ; 15790320
    ; F0F0F0

    ; (16776960)/16 1048560  0   0
    ; (1048560)/16  65535    0   1
    ; (65535)/16    4095     15  2
    ; (4095)/16     255      15  3
    ; (255)/16      15       15  4
    ; (15)/16       0        15  5
    letterArr:=[]
    finalHex:=""

    quotient:=num
    while (quotient > 0) {
        remainder:=Mod(quotient, 16)
        quotient:=Floor(quotient/16)
        letterArr.Push(map1[remainder])
    }
    ; [0, 0, "F", "F", "F", "F"]

    ;reverse iterate
    len:=letterArr.Length() + 1
    while (--len > 0) {
        finalHex.=letterArr[len]
    }
    ; "FFFF00"

return finalHex
}

; just me - https://www.autohotkey.com/board/topic/79600-how-to-get-default-gui-color/#post_id_505635
; http://msdn.microsoft.com/en-us/library/ms724371%28VS.85%29.aspx
get_DefaultBackgroundColor() {
    COLOR_3DFACE := 15 ; Face color for three-dimensional display elements [color=#800000]and for dialog box backgrounds[/color].
    DefaultGUIColor := DllCall("User32.dll\GetSysColor", "Int", COLOR_3DFACE, "UInt")
    ; R := DefaultGUIColor & 0xFF
    ; G := (DefaultGUIColor >> 8) & 0xFF
    ; B := (DefaultGUIColor >> 16) & 0xFF
    ; MsgBox, 0, DefaultGUIColor,  %DefaultGUIColor%:`r`nR = %R%`r`nG = %G%`r`nB = %B%
    ; return R G B
return DefaultGUIColor
}

#Include other examples\CtlColors.ahk
