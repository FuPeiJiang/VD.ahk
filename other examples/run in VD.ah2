#SingleInstance force
ListLines 0
SendMode "Input"
SetWorkingDir A_ScriptDir
KeyHistory 0

#Include %A_LineFile%\..\..\VD.ah2

VD.startShellMessage()
VD.Run("`"C:\Program Files (x86)\Hourglass\Hourglass.exe`"", "", "", "", "Hourglass.exe", 2)
VD.Run("`"C:\Program Files (x86)\Hourglass\Hourglass.exe`"", "", "", "", "Hourglass.exe", 3)
VD.Run("`"C:\Program Files (x86)\Hourglass\Hourglass.exe`"", "", "", "", "Hourglass.exe", 4)

f3::Exitapp