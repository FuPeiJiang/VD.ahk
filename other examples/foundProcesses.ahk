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

foundProcessesArr := []

; Make sure to get all windows from all virtual desktops
DetectHiddenWindows On
WinGet, id, List
Loop %id%
{
    hwnd := id%A_Index%
    ;VD.getDesktopNumOfWindow will filter out invalid windows
    desktopNum_ := VD.getDesktopNumOfWindow("ahk_id" hwnd)
    If (desktopNum_ > -1) ;-1 for invalid window, 0 for "Show on all desktops", 1 for Desktop 1
    {
        WinGet, exe, ProcessName, % "ahk_id" hwnd
        foundProcessesArr.Push({exe:exe, desktopNum_:desktopNum_})
    }
}

foundProcessesArr:=sortArrByKey(foundProcessesArr, "desktopNum_")

finalStr:="('0' for ""Show on all desktops"", '1' for Desktop 1)`n`n"

for unused, v_ in foundProcessesArr {
    finalStr .= v_.desktopNum_ " " v_.exe "`n"
}

MsgBox % finalStr

f3::Exitapp

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