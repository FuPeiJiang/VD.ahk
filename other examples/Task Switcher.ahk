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

activeWindowTitle:=""
MenuItemTitleLength:=100

f1::
global activeWindowTitle
arrayOfWindowsInfo:=[] ;to store {desktopNum:number, str:INFO}
currentDesktop:=VD.getCurrentDesktopNum()
WinGetTitle, activeWindowTitle, A

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

; i_:=arrayOfWindowsInfo.Length()
; lastDesktopNum:=arrayOfWindowsInfo[i_].desktopNum
; while (i_ > 0) {
;
    ; if (!(arrayOfWindowsInfo[i_].desktopNum == lastDesktopNum)) {
        ; lastDesktopNum:=arrayOfWindowsInfo[i_].desktopNum
        ; arrayOfWindowsInfo.InsertAt(i_ + 1, {desktopNum:-2})
    ; }
    ; i_--
; }

ArrForMenuItemPos:=[]
Try
    Menu, windows, DeleteAll

lastDesktopNum:=arrayOfWindowsInfo[1].desktopNum
for k, v in arrayOfWindowsInfo {

    if (!(v.desktopNum == lastDesktopNum)) {
        lastDesktopNum:=v.desktopNum
        Menu, windows, Add
        ArrForMenuItemPos.Push("")
        Menu, windows, Add
        ArrForMenuItemPos.Push("")
    }

    title:=SubStr(v.title, 1, MenuItemTitleLength)
    Menu, windows, Add, % title, ActivateTitle
    ArrForMenuItemPos.Push(v)
    Menu, windows, Add
    ArrForMenuItemPos.Push("")
    Try
        Menu, windows, Icon, % title, % v.ProcessPath,, 0
    Catch
        Menu, windows, Icon, % title, %A_WinDir%\System32\SHELL32.dll, 3, 0
}
DetectHiddenWindows, off

Menu, windows, Color, Silver
defaultItemTitle:=SubStr(activeWindowTitle, 1, MenuItemTitleLength)
Menu, windows, Default, % defaultItemTitle
CoordMode, Menu, Screen
WinGetPos,,, Width, Height,
Xm := (0.4*A_ScreenWidth)
Ym := (0.6*A_ScreenHeight)
; MouseGetPos, OutputVarX, OutputVarY
Menu, windows, Show, % Xm, % Ym

return

ActivateTitle:
    global ArrForMenuItemPos
    ; Tooltip % ArrForMenuItemPos[A_ThisMenuItemPos].title
    VD.goToDesktopOfWindow("ahk_id " ArrForMenuItemPos[A_ThisMenuItemPos].hwnd)
return

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