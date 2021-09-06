#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance, force
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetBatchLines, -1
#KeyHistory 0
ListLines Off

SetWinDelay, -1
SetControlDelay, -1

#Include ..\VD.ahk
VD_init()

lastWindowTitle:=""
activeWindowTitle:=""
ArrayStreamArray:=[]

F1::
global activeWindowTitle,lastWindowTitle,ArrayStreamArray
arrayOfWindowsInfo:=[] ;to store {desktopNum:number, str:INFO}
currentDesktop:=VD_getCurrentDesktop()
WinGetTitle, activeWindowTitle, A

DetectHiddenWindows, on
WinGet windows, List
Loop %windows%
{
    id := windows%A_Index%
    IfEqual, False, % VD_isValidWindow(id), continue
    ahk_idId := "ahk_id " id
    desktopOfWindow:=VD_getDesktopOfWindow(ahk_idId)
    if (desktopOfWindow)
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

if (lastWindowTitle != "")
{
    Menu, windows, DeleteAll
}
for k, v in ArrayStreamArray{
    title:=v["title"]
;MsgBox,%title%
    Menu, windows, Add, %title%, ActivateTitle  
    Menu, windows, Add
;    Menu, windows, Color, Yellow		; 
    WinGet, Path, ProcessPath, %title%
    Try 
        Menu, windows, Icon, %title%, %Path%,, 0
    Catch 
        Menu, windows, Icon, %title%, %A_WinDir%\System32\SHELL32.dll, 3, 0 
}
DetectHiddenWindows, off

Menu, windows, Color, Silver		; 
Menu, windows, Default, %activeWindowTitle%
;CoordMode, Mouse, Screen
;MouseMove, (0.4*A_ScreenWidth), (0.35*A_ScreenHeight)
CoordMode, Menu, Screen
Xm := (0.25*A_ScreenWidth)
Ym := (0.25*A_ScreenHeight)
Menu, windows, Show, %Xm%, %Ym%

;streamArray(Test,1100,200)
return

ActivateTitle:
    global activeWindowTitle, lastWindowTitle,ArrayStreamArray
    DetectHiddenWindows, on
    SetTitleMatchMode 1
    WinActivate, %A_ThisMenuItem%
    DetectHiddenWindows, off
    for k, v in ArrayStreamArray{
        if (activeWindowTitle = v["title"])
        {
            lastWindowTitle:=v["title"]
        }
    }
return

streamArray(Byref arr,Byref width,Byref height)
{
    global ArrayStreamArray, ArrayStreamIndex, ArrayStreamGuiId, ArrayStreamTextId, ArrayStreamIndexTextId, ArrayStreamLength

    ArrayStreamLength:=arr.Length()
    if (ArrayStreamLength)
    {
        ArrayStreamArray:=arr

        Gui, main:New, +hwndArrayStreamHwnd
        ArrayStreamGuiId:="ahk_id " ArrayStreamHwnd
        Gui,Font, s12 Normal, Segoe UI

        gui, add, Text,, Index:
        gui, add, Text, hwndArrayStreamIndexText x+10 w300, 1
        ArrayStreamIndexTextId:="ahk_id " ArrayStreamIndexText
        Gui, Font, s12 Bold

        gui, add, Text, x20 w%width% h%height% hwndArrayStreamTextBox, % ArrayStreamArray[1]

        Gui, Font, s18 Bold
        gui, add, button,w70 h35 gArrayStreamGoLeft, 🠔
        gui, add, button,w70 h35 Default gArrayStreamGoRight x+10, 🠖
        Gui,Font, s12 Normal

        heightPlus:=height+90
        gui, show, w%width% h%heightPlus%
        ArrayStreamTextId:="ahk_id " ArrayStreamTextBox
        ArrayStreamIndex:=1
    }
}
#if winactive(ArrayStreamGuiId)
left::
ArrayStreamGoLeft:
    if (ArrayStreamIndex < 2) {
        SoundPlay, *-1
        return
    }
    ArrayStreamIndex--
    ControlSetText,,% ArrayStreamArray[ArrayStreamIndex], %ArrayStreamTextId%
    ControlSetText,,% ArrayStreamIndex, %ArrayStreamIndexTextId%
return

right::
ArrayStreamGoRight:
    if (ArrayStreamIndex = ArrayStreamLength) {
        SoundPlay, *-1
        return
    }
    ArrayStreamIndex++
    ControlSetText,,% ArrayStreamArray[ArrayStreamIndex], %ArrayStreamTextId%
    ControlSetText,,% ArrayStreamIndex, %ArrayStreamIndexTextId%
return
#if


sortArrByKey(ar,byref key) {
    str=
    for k,v in ar {
        str.=v[key] "+" k "|"
    }
    length:=ar.Length()
    firstValue:=ar[1][key]
    if firstValue is number
    {
        sortType := "N"
    }
    Sort, str, % "D| " sortType
    finalAr:=[]
    finalAr.SetCapacity(length)
    barPos:=1
    loop %length% {
        plusPos:=InStr(str, "+",, barPos)
        barPos:=InStr(str, "|",, plusPos)

        num:=SubStr(str, plusPos + 1, barPos - plusPos - 1)
        finalAr.Push(ar[num])
    }
    return finalAr
}

f3::Exitapp