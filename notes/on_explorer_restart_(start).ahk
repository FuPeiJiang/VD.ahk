#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance force
ListLines Off
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
SetBatchLines -1
#KeyHistory 0
#Persistent

; https://www.autohotkey.com/boards/viewtopic.php?t=63424#p271528
DllCall("RegisterShellHookWindow", "Ptr", A_ScriptHwnd)
MsgNum := DllCall("RegisterWindowMessage", "Str", "SHELLHOOK")
OnMessage(MsgNum, "ShellMessage")

ShellMessage(wParam, lParam) {
    if (wParam == 1) { ; HSHELL_WINDOWCREATED := 1
        bak_DetectHiddenWindows := A_DetectHiddenWindows
        DetectHiddenWindows, ON ;very important

        WinExist("ahk_id " lParam) ;last found window

        explorerRestarted:=false

        WinGetTitle, this_title
        if (this_title == "Start") {
            WinGetClass, this_class
            if (this_class == "Windows.UI.Core.CoreWindow") {
                WinGet, this_exe, ProcessName
                if (this_exe == "StartMenuExperienceHost.exe") {
                    explorerRestarted:=true
                }
            }

        }
        DetectHiddenWindows % bak_DetectHiddenWindows

        if (explorerRestarted) {
            MsgBox % "Explorer Restarted!"
        }
    }
}

return

f3::Exitapp