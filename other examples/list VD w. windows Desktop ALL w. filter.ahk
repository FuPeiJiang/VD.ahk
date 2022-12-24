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

#Include %A_LineFile%\..\..\VD.ahk
VD_init()

arrayOfWindowsInfo:=[] ;to store {desktopNum:number, str:INFO}

DetectHiddenWindows, on
WinGet windows, List
Loop %windows%
{
    id := windows%A_Index%
    IfEqual, False, % VD_isValidWindow(id), continue
    ahk_idId := "ahk_id " id
    desktopOfWindow:=VD_getDesktopOfWindow(ahk_idId)

    if (!desktopOfWindow)
        desktopOfWindow:="ALL"

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

    arrayOfWindowsInfo.Push({desktopNum:desktopOfWindow
        , str:whichDesktop useFulStr notThatUseFulStr
        , WinTitle: OutputTitle
        , class: OutputClass
    , EXE: OutputEXE})
}
DetectHiddenWindows, off

;filter it
filter:=[{EXE:"mbamtray.exe"}
,{WinTitle:"Microsoft Store", EXE:"ApplicationFrameHost.exe"}
,{EXE:"WinStore.App.exe"}
,{WinTitle:"Settings", EXE:"ApplicationFrameHost.exe"}
,{EXE:"SystemSettings.exe"}
,{EXE:"WindowsInternal.ComposableShell.Experiences.TextInput.InputApp.exe"}]

filterArrOfObj(arrayOfWindowsInfo,filter)

arrayOfWindowsInfo:=sortArrByKey(arrayOfWindowsInfo,"desktopNum")

;below is just to print it
ArrayStreamArray:=[]
for k, v in arrayOfWindowsInfo {
    ArrayStreamArray.push(v["str"])
}

streamArray(ArrayStreamArray,1100,200)
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

filterArrOfObj(arr, filter)
{
    ;reverse iterate to remove
    i := arr.Length() + 1
    while (--i)
    {
        v:=arr[i]
        reverseI := length-

        for n, obj in filter {
            for key, value in obj {

                if (value!=v[key])
                    continue 2

            }
            ;if respects the filter, all values match values of filter
            arr.remove(i)
            continue
        }
    }
}

f3::Exitapp
