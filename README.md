# VD.ahk: Virtual Desktop

### Requires Autohotkey alpha: [AHK download page](https://www.autohotkey.com/download/2.1/) (latest: AutoHotkey_2.1-alpha.18)<br>
* v1 branch: [here](https://github.com/FuPeiJiang/VD.ahk/tree/class_VD)<br>
* v2 branch: [here](https://github.com/FuPeiJiang/VD.ahk/tree/v2_port)<br>

### Just run the examples, everything explained inside

* <kbd>Numpad1</kbd> to go to `Desktop 1`<br>
* <kbd>Numpad2</kbd> to go to `Desktop 2`<br>
* <kbd>Numpad3</kbd> to go to `Desktop 3`<br>
```autohotkey
$numpad1::VD.goToDesktopNum(1)
$numpad2::VD.goToDesktopNum(2)
$numpad3::VD.goToDesktopNum(3)
```

* <kbd>Numpad4</kbd> to move the active window to `Desktop 1`<br>
* <kbd>Numpad5</kbd> to move the active window to `Desktop 2`<br>
* <kbd>Numpad6</kbd> to move the active window to `Desktop 3`<br>
* here, I choose to follow the window
```autohotkey
$numpad4::VD.MoveWindowToDesktopNum("A",1,true)
$numpad5::VD.MoveWindowToDesktopNum("A",2,true)
$numpad6::VD.MoveWindowToDesktopNum("A",3,true)
```
* just move window
```autohotkey
$numpad7::VD.MoveWindowToDesktopNum("A",1)
$numpad8::VD.MoveWindowToDesktopNum("A",2)
$numpad9::VD.MoveWindowToDesktopNum("A",3)
```

* Left ← → Right

```autohotkey
; wrapping / cycle back to first desktop when at the last
$^#left::VD.goToRelativeDesktopNum(-1)
$^#right::VD.goToRelativeDesktopNum(+1)

; move window to left and follow it
$#!left::VD.MoveWindowToRelativeDesktopNum("A", -1, true)
; move window to right and follow it
$#!right::VD.MoveWindowToRelativeDesktopNum("A", 1, true)
```

you can remap everything, make sure to use `$` hotkeys if you want to spam them and avoid flashing icons
> ![untitled4](https://cloud.githubusercontent.com/assets/22036272/25830049/8281af16-345a-11e7-8d48-700b252e815a.png)

When writing a script, you can use `VD.WaitDesktopSwitched()`

```ahk
#include %A_LineFile%\..\VD.ahk
; paste image to ms-teams
$f1::{
    SendInput "^c"
    old_desktopNum := VD.getCurrentDesktopNum()
    VD.WaitDesktopSwitched(
        VD.goToDesktopOfWindow("ahk_exe ms-teams.exe")
    )
    SendInput "^r"  ; focus the message composer
    Sleep 100
    SendInput "^v"
    Sleep 100 ; Sleep after ^v but Before Switch
    VD.WaitDesktopSwitched(
        VD.goToDesktopNum(old_desktopNum)
    )
}
```
___
detect when the current desktop changes:<br>
create a hotkey to return you to the **previous desktop**
```ahk
#Include %A_LineFile%\..\VD.ahk
VD.previous_desktopNum := VD.getCurrentDesktopNum()
VD.ListenersCurrentVirtualDesktopChanged[(desktopNum_old, currentDesktopNum) {
    VD.previous_desktopNum := desktopNum_old
    ToolTip "changed from " desktopNum_old " to " currentDesktopNum
}] := 1
$Numpad0::VD.goToDesktopNum(VD.previous_desktopNum)
```
___
note :
`alpha` branch  is a rewrite, if there's functionality missing, I'll add it back
___
also has:
* `createDesktop()`

* rename desktop: `VD.setNameToDesktopNum("custom Desktop Name",desktopNum)`

* "Show this window on all desktops" corresponds to `VD.PinWindow(wintitle)`

* "Show windows from this app on all desktops" corresponds to `VD.PinExe(exe_path)`

- `getCount()` ;how many virtual desktops you now have
- pretty much everything virtual desktop, or so I think!<br>
  if there's anything missing/something you want: [create an issue](https://github.com/FuPeiJiang/VD.ahk/issues/new): I want to know what you're using it for

___
from https://github.com/MScholtes/VirtualDesktop/blob/812c321e286b82a10f8050755c94d21c4b69812f/VirtualDesktop11.cs#L161-L185<br>
Windows 11 has more functionality,<br>
if you need:<br>
* MoveDesktop: `// move current desktop to desktop in index (-> index = 0..Count-1)`<br>
* SetDesktopName<br>
* SetDesktopWallpaper<br>
* GetDesktopIsPerMonitor<br>
* etc.
* or anything else that's not On this list

[create an issue](https://github.com/MScholtes/VirtualDesktop/issues/new)
saying what function(s) you need, what you'll be using it for (I'm curious)
