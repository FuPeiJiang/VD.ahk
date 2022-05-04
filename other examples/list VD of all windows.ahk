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

#Include ..\_VD.ahk
VD.init()

arrayOfWindowsInfo:=[] ;to store {desktopNum:number, str:INFO}

DetectHiddenWindows, on
WinGet windows, List
Loop % windows
{
    hwnd := windows%A_Index%
    desktopOfWindow:=VD.getDesktopNumOfWindow("ahk_id " hwnd)
    if (desktopOfWindow > -1)
    {

        if (desktopOfWindow==0) {
            whichDesktop:="DesktopNum: Show on all desktops"
        } else {
            whichDesktop:="DesktopNum: " desktopOfWindow
        }

        WinGetTitle, Title_, % "ahk_id " hwnd
        WinGetClass, Class_, % "ahk_id " hwnd
        WinGet, Exe_, ProcessName, % "ahk_id " hwnd
        WinGet, PID_, PID, % "ahk_id " hwnd
        WinGet, ID_, ID, % "ahk_id " hwnd
        finalStr:=whichDesktop "`n" Title_ "`nahk_class " Class_ "`nahk_exe " Exe_ "`nahk_pid " PID_ "`nahk_id " ID_ " || " Format("0x{:X}", ID_)

        arrayOfWindowsInfo.Push({desktopNum:desktopOfWindow, str:finalStr})
    }
}
DetectHiddenWindows, off

;below is just to print it
arrayOfWindowsInfo:=sortArrByKey(arrayOfWindowsInfo,"desktopNum")

ArrayStreamArray:=[]
for k, v in arrayOfWindowsInfo {
    ArrayStreamArray.push(v["str"])
}

streamArray(ArrayStreamArray,1100,250)
return

f3::Exitapp

streamArray(arr, width, height)
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

        gui, add, Text, % "x20 w" width - 20 " h" height " hwndArrayStreamTextBox", % ArrayStreamArray[1]

        Gui, Font, s18 Bold
        gui, add, button,w70 h35 gArrayStreamGoLeft, 🠔
        gui, add, button,w70 h35 Default gArrayStreamGoRight x+10, 🠖
        Gui,Font, s12 Normal

        heightPlus:=height+90
        gui, show, % "w" width " h" heightPlus
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


sortArrByKey(ar, key, sortType:="N") {
    str:=""
    for k,v in ar {
        str.=v[key] "+" k "|"
    }
    length:=ar.Length()
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