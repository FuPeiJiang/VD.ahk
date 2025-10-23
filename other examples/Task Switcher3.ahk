#Requires AutoHotkey v2.0
#SingleInstance Force
SendMode "Input"
SetWorkingDir A_ScriptDir
A_KeyHistory := 0
ListLines False

SetWinDelay -1
SetControlDelay -1

#Include "..\VD.ah2"

MenuItemTitleLength := 100
OnMessage(0x0006, HandleMessage) ; to detect gui lose focus

F1:: {
    global myGui
    arrayOfWindowsInfo := [] ; to store {desktopNum:number, str:INFO}

    DetectHiddenWindows True
    windows := WinGetList()

    for id in windows {
        ahk_idId := "ahk_id " id
        desktopOfWindow := VD.getDesktopNumOfWindow(ahk_idId)
        if (desktopOfWindow > -1) {
            OutputTitle := WinGetTitle(ahk_idId)
            OutputProcessPath := WinGetProcessPath(ahk_idId)

            arrayOfWindowsInfo.Push({desktopNum: desktopOfWindow, title: OutputTitle, processPath: OutputProcessPath, hwnd: id})
        }
    }

    arrayOfWindowsInfo := sortArrByKey(arrayOfWindowsInfo, "desktopNum")

    ArrForMenuItemPos := []

    if IsSet(myGui) && IsObject(myGui) {
        myGui.Destroy()
    }
    myGui := Gui("+ToolWindow -Caption")

    guiWidth := 600

    lastDesktopNum := -1
    for index, winInfo in arrayOfWindowsInfo {
        if (!(winInfo.desktopNum == lastDesktopNum)) {
            lastDesktopNum := winInfo.desktopNum
            _name := VD.getNameFromDesktopNum(lastDesktopNum)
            if (_name == "") {
                _name := "Desktop " winInfo.desktopNum
            }

            OMG := myGui.Add("Text", "x10", _name)
            OMG.GetPos(&Xpos, &Ypos, &Width, &Height)
            myGui.Add("Text", "x" (Xpos + Width + 3) " y" (Ypos - 5) " 0x00000005 h1 w" (guiWidth - (Xpos + Width)))
        }

        title := SubStr(winInfo.title, 1, MenuItemTitleLength)
        myGui.Add("Text", "x20", title)
    }
    DetectHiddenWindows False

    CoordMode "Menu", "Screen"
    Xm := (0.4 * A_ScreenWidth)
    Ym := (0.6 * A_ScreenHeight)
    myGui.Show("X" Xm " Y" Ym)
}

F3::ExitApp

HandleMessage(p_w, p_l, p_m, p_hw) {
    global myGui
    if (p_w == 0) {
        if IsSet(myGui) && IsObject(myGui) {
            myGui.Destroy()
        }
    }
}

sortArrByKey(arr, key, sortType := "N") {
    str := ""
    for index, value in arr {
        str .= value.%key% "+" index "|"
    }
    length := arr.Length
    str := Sort(str, "D| " sortType)
    finalAr := []
    finalAr.Capacity := length
    barPos := 1

    loop length {
        plusPos := InStr(str, "+", , barPos)
        barPos := InStr(str, "|", , plusPos)

        num := SubStr(str, plusPos + 1, barPos - plusPos - 1)
        finalAr.Push(arr[num])
    }
    return finalAr
}