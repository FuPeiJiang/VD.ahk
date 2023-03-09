#SingleInstance force
ListLines 0
SendMode "Input"
SetWorkingDir A_ScriptDir
KeyHistory 0
#Include %A_LineFile%\..\..\VD.ah2

VD.RegisterDesktopNotifications()
VD.DefineProp("CurrentVirtualDesktopChanged", {Call:CurrentVirtualDesktopChanged})
CurrentVirtualDesktopChanged(desktopNum_Old, desktopNum_New) {
  ToolTip desktopNum_Old ", " desktopNum_New
}

f3::Exitapp
