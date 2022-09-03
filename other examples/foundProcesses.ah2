#SingleInstance force
ListLines 0
SendMode "Input"
SetWorkingDir A_ScriptDir
KeyHistory 0
#WinActivateForce
ProcessSetPriority "H"
SetWinDelay -1
SetControlDelay -1
#Include %A_LineFile%\..\..\VD.ah2
VD.init()
foundProcessesArr := []
DetectHiddenWindows "On"
id := WinGetList()
loop id.Length {
  hwnd := id[A_Index]
  desktopNum_ := VD.getDesktopNumOfWindow("ahk_id" hwnd)
  if (desktopNum_ > -1) {
    exe := WinGetProcessName("ahk_id" hwnd)
    foundProcessesArr.Push({exe: exe, desktopNum_: desktopNum_})
  }
}
foundProcessesArr := sortArrByKey(foundProcessesArr, "desktopNum_")
finalStr := "('0' for `"Show on all desktops`", '1' for Desktop 1)`n`n"
for unused, v_ in foundProcessesArr {
  finalStr .= v_.desktopNum_ " " v_.exe "`n"
}
MsgBox finalStr
f3::Exitapp
sortArrByKey(arr, key, sortType := "N") {
  global k, Sort, v
  str := ""
  for k, v in arr {
    str .= v.%key% "+" k "|"
  }
  length := arr.Length
  str := Sort(str, "D| " sortType)
  finalAr := []
  finalAr.Capacity := length
  barPos := 1
  loop length {
    plusPos := InStr(str, "+", false, barPos)
    barPos := InStr(str, "|", false, plusPos)
    num := SubStr(str, plusPos + 1, barPos - plusPos - 1)
    finalAr.Push(arr[num])
  }
  return finalAr
}
