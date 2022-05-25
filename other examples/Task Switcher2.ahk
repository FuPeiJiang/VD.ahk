#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance, force
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetBatchLines, -1
#KeyHistory 0
ListLines Off

SetWinDelay, -1
SetControlDelay, -1

#Include ..\_VD.ahk
VD.init()

MenuItemTitleLength:=100

OnMessage( 0x0006, "HandleMessage" ) ;to detect gui lose focus

f1::
arrayOfWindowsInfo:=[] ;to store {desktopNum:number, str:INFO}

DetectHiddenWindows, on
WinGet windows, List
Loop %windows%
{
    id := windows%A_Index%
    ahk_idId := "ahk_id " id
    desktopOfWindow:=VD.getDesktopNumOfWindow(ahk_idId)
    if (desktopOfWindow > -1)
    {
        whichDesktop:="Desktop " desktopOfWindow

        WinGetTitle, OutputTitle, % ahk_idId
        WinGet, OutputProcessPath, ProcessPath, % ahk_idId

        arrayOfWindowsInfo.Push({desktopNum:desktopOfWindow, title:OutputTitle, processPath:OutputProcessPath, hwnd:id})
    }
}

arrayOfWindowsInfo:=sortArrByKey(arrayOfWindowsInfo,"desktopNum")

ArrForMenuItemPos:=[]

Gui, Destroy
Gui, New, -0xC00000

guiWidth:=600

lastDesktopNum:=-1
for k, v in arrayOfWindowsInfo {

    if (!(v.desktopNum == lastDesktopNum)) {
        lastDesktopNum:=v.desktopNum
        Gui, Add, Text, % "x10 hwndOMG", % "Desktop " v.desktopNum
        ControlGetPos, Xpos, Ypos, Width, Height,, % "ahk_id " OMG
        Gui, Add, Text, % "x+3 y+-5 0x00000005 h1 w" guiWidth - (Xpos + Width)
    }

    title:=SubStr(v.title, 1, MenuItemTitleLength)
    Gui, Add, Text, % "x20", % title
}
DetectHiddenWindows, off

CoordMode, Menu, Screen
WinGetPos,,, Width, Height,
Xm := (0.4*A_ScreenWidth)
Ym := (0.6*A_ScreenHeight)
; MouseGetPos, OutputVarX, OutputVarY
Gui, Show, % "X" Xm " Y" Ym

return

HandleMessage( p_w, p_l, p_m, p_hw )
{
    global
    ; ToolTip % p_w ", " p_l
    if (p_w==0) {
        Gui, Destroy
    }
}

sortArrByKey(arr, key, sortType:="N") {
    str:=""
    for k,v in arr {
        str.=v[key] "+" k "|"
    }
    length:=arr.Length()
    Sort, str, % "D| " sortType
    finalAr:=[]
    finalAr.SetCapacity(length)
    barPos:=1
    loop %length% {
        plusPos:=InStr(str, "+",, barPos)
        barPos:=InStr(str, "|",, plusPos)

        num:=SubStr(str, plusPos + 1, barPos - plusPos - 1)
        finalAr.Push(arr[num])
    }
    return finalAr
}

f3::Exitapp