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

arrayOfWindowsInfo:=[] ;to store {desktopNum:number, str:INFO}

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

        arrayOfWindowsInfo.Push({desktopNum:desktopOfWindow, str:whichDesktop useFulStr notThatUseFulStr})
    }
}
DetectHiddenWindows, off

;below is just to print it
arrayOfWindowsInfo:=sortArrByKey(arrayOfWindowsInfo,"desktopNum")

ArrayStreamArray:=[]
for k, v in arrayOfWindowsInfo {
    ArrayStreamArray.push(v["str"])
}

streamArray(ArrayStreamArray,1000,800)
return


streamArray(Byref arr,Byref width,Byref height)
{
    global ArrayStreamArray, inputStreamIndex, inputStreamTextId, ArrayStreamLength

    ArrayStreamLength:=arr.Length()
    if (ArrayStreamLength)
    {
        ArrayStreamArray:=arr


        Gui, main:New, +hwndShowTextHwnd
        Gui,Font, s12, Segoe UI
        gui, add, text, w%width% h%height% hwndarrayStreamTextBox, % ArrayStreamArray[1]
        gui, add, button,Default gcontinueArrayStream, continue
        heightPlus:=height+60
        gui, show, w%width% h%heightPlus%
        inputStreamTextId:=ahk_id %inputStreamTextHwnd%
        ControlSetText,, %NewText%, %inputStreamTextId%
        inputStreamIndex:=2
    }

}
continueArrayStream:
    ControlSetText,,% ArrayStreamArray[inputStreamIndex], %inputStreamTextId%
    if (inputStreamIndex=ArrayStreamLength) {
        WinClose, %inputStreamTextId%
        return
    }
    inputStreamIndex++
return


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
