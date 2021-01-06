## VD.ahk : Virtual Desktop (move window, go to VD of your choice instantly, get Which VD you are in, etc.)

https://www.autohotkey.com/boards/viewtopic.php?f=6&t=83381

# Just run the examples, everything explained inside
the most useful ones:<br>
* Numpad1 to go to Desktop 1<br>
* Numpad2 to go to Desktop 2<br>
* Numpad3 to go to Desktop 3<br>
- Numpad4 to move the active window to Desktop 1<br>
- Numpad5 to move the active window to Desktop 2<br>
- Numpad6 to move the active window to Desktop 3<br>
you can remap everything
<!-- Desktop2`nPress Numpad6 to move the active window to Desktop3 and go to Desktop 3 (follow the window) -->

## cool fixes:<br>
* Switching VD does not make icons (on the taskbar) flash<br>
https://github.com/mzomparelli/zVirtualDesktop/issues/59#issue-227209226
> Sometimes when I switch desktop, the application that's focused on that desktop becomes highlighted in the task bar.
> 
> Example:
> On desktop A, I have Firefox:<br>
> ![untitled](https://cloud.githubusercontent.com/assets/22036272/25830018/467f9c3a-345a-11e7-91a0-3d2a633fae68.png)
> 
> On desktop B I have Steam (note that Firefox is pinned to my task bar, which is why it's visible on this desktop):<br>
> ![untitled2](https://cloud.githubusercontent.com/assets/22036272/25830028/563f7a3c-345a-11e7-8672-f0e43baf440f.png)
> 
> When I switch from A to B, Steam becomes highlighted:<br>
> ![untitled3](https://cloud.githubusercontent.com/assets/22036272/25830040/675eff36-345a-11e7-970b-9a689eec74b3.png)
> (It's possible to stop the blinking by explicitly selecting the blinking window with the mouse or alt+tab.)
> 
> When I switch back to A, Firefox also becomes highlighted (and Steam stays visible in the taskbar because it is highlighted):<br>
> ![untitled4](https://cloud.githubusercontent.com/assets/22036272/25830049/8281af16-345a-11e7-8d48-700b252e815a.png)
> 
> This doesn't always happen; I'm not sure if it depends on the programs or where the mouse/keyboard focus is when switching desktops. That said, it happens far too often, to the point where it defeats the point of having multiple desktops at all.
> 
> (The problem isn't limited to Firefox and Steam, the same thing also happens with other programs like Explorer.)
<br>

how ? `WinActivate` taskbar before switching and `WinMinimize` taskbar after arriving
* Switch VD reliably works in FULLSCREEN thanks to `SetTimer, pleaseSwitchDesktop, -50`



