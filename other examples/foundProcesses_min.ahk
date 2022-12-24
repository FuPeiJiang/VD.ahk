#Include %A_LineFile%\..\..\VD.ahk

foundProcesses := ""
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
        foundProcesses .= desktopNum_ " " exe "`n"
    }
}

MsgBox % foundProcesses