#Include %A_LineFile%\..\..\VD.ah2
foundProcesses := ""
DetectHiddenWindows "On"
id := WinGetList()
loop id.Length {
  hwnd := id[A_Index]
  desktopNum_ := VD.getDesktopNumOfWindow("ahk_id" hwnd)
  if (desktopNum_ > -1) {
    exe := WinGetProcessName("ahk_id" hwnd)
    foundProcesses .= desktopNum_ " " exe "`n"
  }
}
MsgBox foundProcesses
