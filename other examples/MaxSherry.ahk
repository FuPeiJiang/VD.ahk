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

lastWindowTitle:=""
activeWindowTitle:=""
ArrayStreamArray:=[]
MenuItemTitleLength:=100

F1::
global activeWindowTitle,lastWindowTitle,ArrayStreamArray
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

        ;useful
        WinGetTitle, OutputTitle, %ahk_idId%
        WinGetClass, OutputClass, %ahk_idId%
        WinGet, OutputEXE, ProcessName, %ahk_idId%
        useFulStr:="`nWinTitle: " OutputTitle "`nclass: " OutputClass "`nEXE: " OutputEXE

        ;not that useful
        WinGet, OutputFULLPATH, ProcessPath, %ahk_idId%
        WinGet, OutputPID, PID, %ahk_idId%

        notThatUseFulStr:="`n`nFULLPATH: " OutputFULLPATH "`nPID: " OutputPID "`nID: " id
        WinGet, OutputVar, ProcessPath, A

        arrayOfWindowsInfo.Push({desktopNum:desktopOfWindow, title:OutputTitle, cls:OutputClass, str:whichDesktop useFulStr notThatUseFulStr})
    }
}

;below is just to print it
arrayOfWindowsInfo:=sortArrByKey(arrayOfWindowsInfo,"desktopNum")

ArrayStreamArray:=[]
for k, v in arrayOfWindowsInfo {
    if (lastWindowTitle != "") and (lastWindowTitle = v["title"]) and WinExist(lastWindowTitle)
    {
;MsgBox,% v["title"] . "aa"
        ArrayStreamArray.push(v)
    }
}
for k, v in arrayOfWindowsInfo {
    if (activeWindowTitle = v["title"])
    {
;MsgBox,% v["title"] . "bb"
        ArrayStreamArray.push(v)
    }
}
for k, v in arrayOfWindowsInfo {
    if (currentDesktop = v["desktopNum"]) and (v["title"] != activeWindowTitle) and (v["title"] != lastWindowTitle)
    {
        ArrayStreamArray.push(v)
    }
}
for k, v in arrayOfWindowsInfo {
    if (currentDesktop != v["desktopNum"]) and (v["title"] != activeWindowTitle) and (v["title"] != lastWindowTitle)
    {
        ArrayStreamArray.push(v)
    }
}

desktopNum:=currentDesktop
skipItem:=1
if (lastWindowTitle != "")
{
    Menu, windows, DeleteAll
    skipItem:=2
}
for k, v in ArrayStreamArray{
    if (k > skipItem) and (desktopNum != v["desktopNum"])
    {
        Menu, windows, Add
        desktopNum:=v["desktopNum"]
    }
    title:=SubStr(v["title"],1, MenuItemTitleLength)
    Menu, windows, Add, %title%, ActivateTitle
    Menu, windows, Add
    WinGet, Path, ProcessPath, %title%
    Try
        Menu, windows, Icon, %title%, %Path%,, 0
    Catch
        Menu, windows, Icon, %title%, %A_WinDir%\System32\SHELL32.dll, 3, 0
}
DetectHiddenWindows, off

Menu, windows, Color, Silver		;
defaultItemTitle:=SubStr(activeWindowTitle, 1, MenuItemTitleLength)
Menu, windows, Default, %defaultItemTitle%
;CoordMode, Mouse, Screen
;MouseMove, (0.4*A_ScreenWidth), (0.35*A_ScreenHeight)
CoordMode, Menu, Screen
Xm := (0.25*A_ScreenWidth)
Ym := (0.25*A_ScreenHeight)
Menu, windows, Show, %Xm%, %Ym%

return

ActivateTitle:
    global activeWindowTitle, lastWindowTitle,ArrayStreamArray
    DetectHiddenWindows, on
    SetTitleMatchMode 1
    if WinExist(A_ThisMenuItem)
        WinActivate
    DetectHiddenWindows, off
    activeTitle:=SubStr(activeWindowTitle, 1, MenuItemTitleLength)
    for k, v in ArrayStreamArray {
        if (A_ThisMenuItem != activeTitle) and (activeWindowTitle = v["title"])
        {
            lastWindowTitle:=v["title"]
        }
    }
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