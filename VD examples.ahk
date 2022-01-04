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

Gui,Font, s12, Segoe UI

explanation=
(
Numpad0 to pin this Window on all desktops
you can spam (Numpad2,Numpad1,Numpad2,Numpad1) for fun

here's a challenge (you might lose this window):
Unpin this using Numpad0
go to Desktop 3 (Numpad3)
this time, use Win + * on Numpad to come back to this window wherever you are
(and wherever this window is)
so you can move this window to desktop 2 (Numpad5), you go to desktop 1, and use Win + * on Numpad
(if you want to search in this script, the hotkey is #NumpadMult)

Numpad9 to throw a window to Desktop 3 (and not follow it)

getters:
f1 to see which desktop you currently are in
f6 to see which desktop this window is in
f2 to see the total number of virtual desktops

(You might want to pin this window for this part):
!NumpadAdd (Alt + Numpad+) to createDesktop and go to it
f1 to see which desktop you currently are in

but at this point, just use Win + Tab..
these functions are mostly for script only,
for example: I used VD.createUntil(3)
at the start of this tutorial, to make sure we have at least 3 VD

^+NumpadAdd (Ctrl Alt + Numpad+) to create until you have 3 desktops
!NumpadSub (Alt + Numpad-) to remove the current desktop
^+NumpadSub (Ctrl ALt + Numpad-) to delete the 3rd desktop

more below, look at the hotkeys in code.
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
gui, show,, VD.ahk examples WinTitle
WinSet, Redraw
Postmessage,0xB1,0,StrLen("Numpad0"),, % "ahk_id " HWndExplanation_Edit ;move caret to end
; Postmessage,0xB1,-1,-1,, % "ahk_id " HWndExplanation_Edit ;move caret to end
; Postmessage,0xB1,9,9,, % "ahk_id " HWndExplanation_Edit ;move caret to: ^ is Ctrl|

;include the library
#Include VD.ahk
; VD.init() ;optional

;you should WinHide invisible programs that have a window.
WinHide, % "Malwarebytes Tray Application"
;#SETUP END

VD.createUntil(3) ;create until we have at least 3 VD

return

;#useful stuff
numpad1::VD.goToDesktopNum(1)
numpad2::VD.goToDesktopNum(2)
numpad3::VD.goToDesktopNum(3)

;followYourWindow:=true
numpad4::VD.sendWindowToDesktopNum("A",1,true)
numpad5::VD.sendWindowToDesktopNum("A",2,true)
numpad6::VD.sendWindowToDesktopNum("A",3,true)

;followYourWindow:=false
numpad7::VD.sendWindowToDesktopNum("A",1,false)
numpad8::VD.sendWindowToDesktopNum("A",2,false)
numpad9::VD.sendWindowToDesktopNum("A",3,false)

;to come back to this window
#NumpadMult:: ;#*
    VD.goToDesktopOfWindow("VD.ahk examples WinTitle")
    ; VD.goToDesktopOfWindow("ahk_exe code.exe")
return

;getters and stuff
f6::
    msgbox % VD.getDesktopNumOfWindow("VD.ahk examples WinTitle")
    ; msgbox % VD.getDesktopNumOfWindow("ahk_exe GitHubDesktop.exe")
return
f1::
    msgbox % VD.getCurrentDesktopNum()
return
f2::
    msgbox % VD.getCount()
return

;Create/Remove Desktop
!NumpadAdd::VD.createDesktop()
#NumpadAdd::VD.createDesktop(false) ;don't go to newly created

!NumpadSub::VD.removeDesktop(VD.getCurrentDesktopNum())
#!NumpadSub::VD.removeDesktop(VD.getCount()) ;removes 3rd desktop if there are 3 desktops

^+NumpadAdd::
VD.createUntil(3) ;create until we have at least 3 VD
return

^+NumpadSub::
VD.createUntil(3) ;create until we have at least 3 VD
sleep 1000
;FALLBACK IS ONLY USED IF YOU ARE CURRENTLY ON THAT VD
VD.removeDesktop(3, 1)
return

;Pin Window
numpad0::
    VD.TogglePinWindow("A")
return
^numpad0::
    VD.PinWindow("A")
return
!numpad0::
    VD.UnPinWindow("A")
return
#numpad0::
    MsgBox % VD.IsWindowPinned("A")
return

;Pin App
numpadDot::
    VD.TogglePinApp("A")
return
^numpadDot::
    VD.PinApp("A")
return
!numpadDot::
    VD.UnPinApp("A")
return
#numpadDot::
    MsgBox % VD.IsAppPinned("A")
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
