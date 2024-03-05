; VD.getCurrentDesktopNum()
; VD.getDesktopNumOfWindow(wintitle) ;please use VD.goToDesktopOfWindow instead if you just want to go there.
; VD.getDesktopNumOfWindow(wintitle) ;returns 0 for "Show on all desktops"

; VD.getCount() ;how many virtual desktops you now have
; VD.getRelativeDesktopNum(anchor_desktopNum, relative_count)

; VD.goToDesktopNum(desktopNum)
; VD.goToDesktopOfWindow(wintitle, activateYourWindow:=true)
; VD.gotoRelativeDesktopNum(relative_count)

; VD.MoveWindowToDesktopNum(wintitle, desktopNum)
; VD.MoveWindowToCurrentDesktop(wintitle, activateYourWindow:=true)
; VD.MoveWindowToRelativeDesktopNum(wintitle, relative_count)

; VD.createDesktop(goThere:=false) ; VD.createUntil(howMany, goToLastlyCreated:=true)
; VD.removeDesktop(desktopNum, fallback_desktopNum:=false)

; "Show this window on all desktops"
; VD.IsWindowPinned(wintitle)
; VD.TogglePinWindow(wintitle)
; VD.PinWindow(wintitle)
; VD.UnPinWindow(wintitle)

; "Show windows from this app on all desktops"
; VD.IsAppPinned(wintitle)
; VD.TogglePinApp(wintitle)
; VD.PinApp(wintitle)
; VD.UnPinApp(wintitle)

; Thanks to:
; Blackholyman:
; https://www.autohotkey.com/boards/viewtopic.php?t=67642#p291160
; and
; Flipeador:
; https://www.autohotkey.com/boards/viewtopic.php?t=54202#p234192
; https://www.autohotkey.com/boards/viewtopic.php?t=54202#p234309
; and then later
; MScholtes:
; https://github.com/MScholtes/VirtualDesktop/blob/812c321e286b82a10f8050755c94d21c4b69812f/VirtualDesktop.cs
; (for all the other functions I didnt know about. and windows 11)

class VD {

    ; #Include %A_LineFile%\..\VD.ahk
    ; or
    ; #Include %A_LineFile%\..\_VD.ahk
    ; ...{startup code}
    ; VD.init()

    ; VD.ahk : calls `VD.init()` on #Include
    ; _VD.ahk : `VD.init()` when you want, like after a GUI has rendered, for startup performance reasons
    init() {
        FileGetVersion, OS_Version, % A_WinDir "\System32\twinui.pcshell.dll"
        splitByDot:=StrSplit(OS_Version, ".")
        buildNumber:=splitByDot[3]+0
        revisionNumber:=splitByDot[4]+0
        if (buildNumber < 20348) {
            ;from 17763.1
            IID_IVirtualDesktopManagerInternal_str:="{f31574d6-b682-4cdc-bd56-1827860abec6}"
            IID_IVirtualDesktop_str:="{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}"
            ;IID_IVirtualDesktopNotification_str:="{c179334c-4295-40d3-bea1-c654d965605a}"
            this.IID_IVirtualDesktopNotification_n1:=4671150449476842316
            this.IID_IVirtualDesktopNotification_n2:=6512317045282349502

            idx_MoveViewToDesktop:=4 ;DllCall(ptr_MoveViewToDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IApplicationView,"Ptr",IVirtualDesktop)
            idx_GetCurrentDesktop:=6 ;DllCall(ptr_GetCurrentDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr*",&IVirtualDesktop:=0)
            idx_GetDesktops:=7 ;DllCall(ptr_GetDesktops,"Ptr",IVirtualDesktopManagerInternal,"Ptr*",&IObjectArray:=0)
            idx_CreateDesktop:=10 ;DllCall(ptr_CreateDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr*",&IVirtualDesktop:=0)
            idx_RemoveDesktop:=11 ;DllCall(ptr_RemoveDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",IVirtualDesktop_fallback)
            this._dll_MoveViewToDesktop:=this._dll_MoveViewToDesktop_normal
            this._dll_GetCurrentDesktop:=this._dll_GetCurrentDesktop_normal
            this._dll_GetDesktops:=this._dll_GetDesktops_normal
            this._dll_CreateDesktop:=this._dll_CreateDesktop_normal
            this._dll_RemoveDesktop:=this._dll_RemoveDesktop_normal

            this.idx_GetId:=4 ;DllCall(ptr_GetId,"Ptr",IVirtualDesktop,"Ptr",guid_buf)

            this.idx_VirtualDesktopNameChanged:=-1
            this.idx_SetDesktopName:=-1
            this.idx_GetName:=-1

            this.idx_VirtualDesktopWallpaperChanged:=-1
            this.idx_SetDesktopWallpaper:=-1
            this.idx_GetWallpaper:=-1

            this.idx_VirtualDesktopCreated:=3 ;params (IVirtualDesktop)
            this.idx_VirtualDesktopDestroyBegin:=4 ;params (IVirtualDesktop, IVirtualDesktop_fallback)
            this.idx_VirtualDesktopDestroyFailed:=5 ;params (IVirtualDesktop, IVirtualDesktop_fallback)
            this.idx_VirtualDesktopDestroyed:=6 ;params (IVirtualDesktop, IVirtualDesktop_fallback)
            this.idx_ViewVirtualDesktopChanged:=7 ;params (IApplicationView)
            this.idx_CurrentVirtualDesktopChanged:=8 ;params (IVirtualDesktop_old, IVirtualDesktop_new)
            this._dll_VirtualDesktopCreated:=VD._dll_VirtualDesktopCreated_normal
            this._dll_VirtualDesktopDestroyBegin:=VD._dll_VirtualDesktopDestroyBegin_normal
            this._dll_VirtualDesktopDestroyFailed:=VD._dll_VirtualDesktopDestroyFailed_normal
            this._dll_VirtualDesktopDestroyed:=VD._dll_VirtualDesktopDestroyed_normal
            this._dll_ViewVirtualDesktopChanged:=VD._dll_ViewVirtualDesktopChanged_normal
            this._dll_CurrentVirtualDesktopChanged:=VD._dll_CurrentVirtualDesktopChanged_normal

            this.IVirtualDesktopNotification_methods_count:=9
        } else if (buildNumber < 22000) { ;22000.51 to be more precise
            ;from 20348.2227 - Windows Server 2022
            IID_IVirtualDesktopManagerInternal_str:="{094afe11-44f2-4ba0-976f-29a97e263ee0}"
            IID_IVirtualDesktop_str:="{62fdf88b-11ca-4afb-8bd8-2296dfae49e2}"
            ;IID_IVirtualDesktopNotification_str:="{f3163e11-6b04-433c-a64b-6f82c9094257}"
            this.IID_IVirtualDesktopNotification_n1:=4844864968146173457
            this.IID_IVirtualDesktopNotification_n2:=6287598790844042150

            idx_MoveViewToDesktop:=4 ;DllCall(ptr_MoveViewToDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IApplicationView,"Ptr",IVirtualDesktop)
            idx_GetCurrentDesktop:=6 ;DllCall(ptr_GetCurrentDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",HMONITOR,"Ptr*",&IVirtualDesktop:=0)
            idx_GetDesktops:=7 ;DllCall(ptr_GetDesktops,"Ptr",IVirtualDesktopManagerInternal,"Ptr",HMONITOR,"Ptr*",&IObjectArray:=0)
            idx_CreateDesktop:=10 ;DllCall(ptr_CreateDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",HMONITOR,"Ptr*",&IVirtualDesktop:=0)
            idx_RemoveDesktop:=11 ;DllCall(ptr_RemoveDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",IVirtualDesktop_fallback)
            this._dll_MoveViewToDesktop:=this._dll_MoveViewToDesktop_normal
            this._dll_GetCurrentDesktop:=this._dll_GetCurrentDesktop_HMONITOR
            this._dll_GetDesktops:=this._dll_GetDesktops_HMONITOR
            this._dll_CreateDesktop:=this._dll_CreateDesktop_HMONITOR
            this._dll_RemoveDesktop:=this._dll_RemoveDesktop_normal

            this.idx_GetId:=4 ;DllCall(ptr_GetId,"Ptr",IVirtualDesktop,"Ptr",guid_buf)

            this.idx_VirtualDesktopNameChanged:=8 ;params (IVirtualDesktop, HSTRING)
            this._dll_VirtualDesktopNameChanged:=VD._dll_VirtualDesktopNameChanged_normal
            this.idx_SetDesktopName:=14 ;DllCall(ptr_SetDesktopName,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",HSTRING)
            this.idx_GetName:=6 ;DllCall(ptr_GetName,"Ptr",IVirtualDesktop,"Ptr*",HSTRING)

            this.idx_VirtualDesktopWallpaperChanged:=-1
            this.idx_SetDesktopWallpaper:=-1
            this.idx_GetWallpaper:=-1

            this.idx_VirtualDesktopCreated:=3 ;params (IVirtualDesktop)
            this.idx_VirtualDesktopDestroyBegin:=4 ;params (IVirtualDesktop, IVirtualDesktop_fallback)
            this.idx_VirtualDesktopDestroyFailed:=5 ;params (IVirtualDesktop, IVirtualDesktop_fallback)
            this.idx_VirtualDesktopDestroyed:=6 ;params (IVirtualDesktop, IVirtualDesktop_fallback)
            this.idx_ViewVirtualDesktopChanged:=9 ;params (IApplicationView)
            this.idx_CurrentVirtualDesktopChanged:=10 ;params (IVirtualDesktop_old, IVirtualDesktop_new)
            this._dll_VirtualDesktopCreated:=VD._dll_VirtualDesktopCreated_normal
            this._dll_VirtualDesktopDestroyBegin:=VD._dll_VirtualDesktopDestroyBegin_normal
            this._dll_VirtualDesktopDestroyFailed:=VD._dll_VirtualDesktopDestroyFailed_normal
            this._dll_VirtualDesktopDestroyed:=VD._dll_VirtualDesktopDestroyed_normal
            this._dll_ViewVirtualDesktopChanged:=VD._dll_ViewVirtualDesktopChanged_normal
            this._dll_CurrentVirtualDesktopChanged:=VD._dll_CurrentVirtualDesktopChanged_normal

            this.IVirtualDesktopNotification_methods_count:=11
        } else if (buildNumber < 22483) { ;22483.1000 to be more precise
            ;from 22000.51
            IID_IVirtualDesktopManagerInternal_str:="{b2f925b9-5a0f-4d2e-9f4d-2b1507593c10}"
            IID_IVirtualDesktop_str:="{536d3495-b208-4cc9-ae26-de8111275bf8}"
            ;IID_IVirtualDesktopNotification_str:="{cd403e52-deed-4c13-b437-b98380f2b1e8}"
            this.IID_IVirtualDesktopNotification_n1:=5481970284372180562
            this.IID_IVirtualDesktopNotification_n2:=-1679294552252794956

            idx_MoveViewToDesktop:=4 ;DllCall(ptr_MoveViewToDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IApplicationView,"Ptr",IVirtualDesktop)
            idx_GetCurrentDesktop:=6 ;DllCall(ptr_GetCurrentDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",HMONITOR,"Ptr*",&IVirtualDesktop:=0)
            idx_GetDesktops:=7 ;DllCall(ptr_GetDesktops,"Ptr",IVirtualDesktopManagerInternal,"Ptr",HMONITOR,"Ptr*",&IObjectArray:=0)
            idx_CreateDesktop:=10 ;DllCall(ptr_CreateDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",HMONITOR,"Ptr*",&IVirtualDesktop:=0)
            idx_RemoveDesktop:=12 ;DllCall(ptr_RemoveDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",IVirtualDesktop_fallback)
            this._dll_MoveViewToDesktop:=this._dll_MoveViewToDesktop_normal
            this._dll_GetCurrentDesktop:=this._dll_GetCurrentDesktop_HMONITOR
            this._dll_GetDesktops:=this._dll_GetDesktops_HMONITOR
            this._dll_CreateDesktop:=this._dll_CreateDesktop_HMONITOR
            this._dll_RemoveDesktop:=this._dll_RemoveDesktop_normal

            this.idx_GetId:=4 ;DllCall(ptr_GetId,"Ptr",IVirtualDesktop,"Ptr",guid_buf)

            this.idx_VirtualDesktopNameChanged:=9 ;params (IVirtualDesktop, HSTRING)
            this._dll_VirtualDesktopNameChanged:=VD._dll_VirtualDesktopNameChanged_normal
            this.idx_SetDesktopName:=15 ;DllCall(ptr_SetDesktopName,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",HSTRING)
            this.idx_GetName:=6 ;DllCall(ptr_GetName,"Ptr",IVirtualDesktop,"Ptr*",HSTRING)

            this.idx_VirtualDesktopWallpaperChanged:=12 ;params (IVirtualDesktop, HSTRING)
            this._dll_VirtualDesktopWallpaperChanged:=VD._dll_VirtualDesktopWallpaperChanged_normal
            this.idx_SetDesktopWallpaper:=16 ;DllCall(ptr_SetDesktopWallpaper,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",HSTRING)
            this.idx_GetWallpaper:=7 ;DllCall(ptr_GetWallpaper,"Ptr",IVirtualDesktop,"Ptr*",HSTRING)

            this.idx_VirtualDesktopCreated:=3 ;params (IObjectArray, IVirtualDesktop)
            this.idx_VirtualDesktopDestroyBegin:=4 ;params (IObjectArray, IVirtualDesktop, IVirtualDesktop_fallback)
            this.idx_VirtualDesktopDestroyFailed:=5 ;params (IObjectArray, IVirtualDesktop, IVirtualDesktop_fallback)
            this.idx_VirtualDesktopDestroyed:=6 ;params (IObjectArray, IVirtualDesktop, IVirtualDesktop_fallback)
            this.idx_ViewVirtualDesktopChanged:=10 ;params (IApplicationView)
            this.idx_CurrentVirtualDesktopChanged:=11 ;params (IObjectArray, IVirtualDesktop_old, IVirtualDesktop_new)
            this._dll_VirtualDesktopCreated:=VD._dll_VirtualDesktopCreated_IObjectArray
            this._dll_VirtualDesktopDestroyBegin:=VD._dll_VirtualDesktopDestroyBegin_IObjectArray
            this._dll_VirtualDesktopDestroyFailed:=VD._dll_VirtualDesktopDestroyFailed_IObjectArray
            this._dll_VirtualDesktopDestroyed:=VD._dll_VirtualDesktopDestroyed_IObjectArray
            this._dll_ViewVirtualDesktopChanged:=VD._dll_ViewVirtualDesktopChanged_normal
            this._dll_CurrentVirtualDesktopChanged:=VD._dll_CurrentVirtualDesktopChanged_IObjectArray

            this.IVirtualDesktopNotification_methods_count:=13
        } else if (buildNumber <= 22621 && revisionNumber < 2215) {
            ;from 22483.1000
            ;from 22621.1778 (they're identical)
            ;yeah yeah, IID are the same as above, but vftable differs
            IID_IVirtualDesktopManagerInternal_str:="{b2f925b9-5a0f-4d2e-9f4d-2b1507593c10}"
            IID_IVirtualDesktop_str:="{536d3495-b208-4cc9-ae26-de8111275bf8}"
            ;IID_IVirtualDesktopNotification_str:="{cd403e52-deed-4c13-b437-b98380f2b1e8}"
            this.IID_IVirtualDesktopNotification_n1:=5481970284372180562
            this.IID_IVirtualDesktopNotification_n2:=-1679294552252794956

            idx_MoveViewToDesktop:=4 ;DllCall(ptr_MoveViewToDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IApplicationView,"Ptr",IVirtualDesktop)
            idx_GetCurrentDesktop:=6 ;DllCall(ptr_GetCurrentDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",HMONITOR,"Ptr*",&IVirtualDesktop:=0)
            idx_GetDesktops:=8 ;DllCall(ptr_GetDesktops,"Ptr",IVirtualDesktopManagerInternal,"Ptr",HMONITOR,"Ptr*",&IObjectArray:=0)
            idx_CreateDesktop:=11 ;DllCall(ptr_CreateDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",HMONITOR,"Ptr*",&IVirtualDesktop:=0)
            idx_RemoveDesktop:=13 ;DllCall(ptr_RemoveDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",IVirtualDesktop_fallback)
            this._dll_MoveViewToDesktop:=this._dll_MoveViewToDesktop_normal
            this._dll_GetCurrentDesktop:=this._dll_GetCurrentDesktop_HMONITOR
            this._dll_GetDesktops:=this._dll_GetDesktops_HMONITOR
            this._dll_CreateDesktop:=this._dll_CreateDesktop_HMONITOR
            this._dll_RemoveDesktop:=this._dll_RemoveDesktop_normal

            this.idx_GetId:=4 ;DllCall(ptr_GetId,"Ptr",IVirtualDesktop,"Ptr",guid_buf)

            this.idx_VirtualDesktopNameChanged:=9 ;params (IVirtualDesktop, HSTRING)
            this._dll_VirtualDesktopNameChanged:=VD._dll_VirtualDesktopNameChanged_normal
            this.idx_SetDesktopName:=16 ;DllCall(ptr_SetDesktopName,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",HSTRING)
            this.idx_GetName:=6 ;DllCall(ptr_GetName,"Ptr",IVirtualDesktop,"Ptr*",HSTRING)

            this.idx_VirtualDesktopWallpaperChanged:=12 ;params (IVirtualDesktop, HSTRING)
            this._dll_VirtualDesktopWallpaperChanged:=VD._dll_VirtualDesktopWallpaperChanged_normal
            this.idx_SetDesktopWallpaper:=17 ;DllCall(ptr_SetDesktopWallpaper,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",HSTRING)
            this.idx_GetWallpaper:=7 ;DllCall(ptr_GetWallpaper,"Ptr",IVirtualDesktop,"Ptr*",HSTRING)

            this.idx_VirtualDesktopCreated:=3 ;params (IObjectArray, IVirtualDesktop)
            this.idx_VirtualDesktopDestroyBegin:=4 ;params (IObjectArray, IVirtualDesktop, IVirtualDesktop_fallback)
            this.idx_VirtualDesktopDestroyFailed:=5 ;params (IObjectArray, IVirtualDesktop, IVirtualDesktop_fallback)
            this.idx_VirtualDesktopDestroyed:=6 ;params (IObjectArray, IVirtualDesktop, IVirtualDesktop_fallback)
            this.idx_ViewVirtualDesktopChanged:=10 ;params (IApplicationView)
            this.idx_CurrentVirtualDesktopChanged:=11 ;params (IObjectArray, IVirtualDesktop_old, IVirtualDesktop_new)
            this._dll_VirtualDesktopCreated:=VD._dll_VirtualDesktopCreated_IObjectArray
            this._dll_VirtualDesktopDestroyBegin:=VD._dll_VirtualDesktopDestroyBegin_IObjectArray
            this._dll_VirtualDesktopDestroyFailed:=VD._dll_VirtualDesktopDestroyFailed_IObjectArray
            this._dll_VirtualDesktopDestroyed:=VD._dll_VirtualDesktopDestroyed_IObjectArray
            this._dll_ViewVirtualDesktopChanged:=VD._dll_ViewVirtualDesktopChanged_normal
            this._dll_CurrentVirtualDesktopChanged:=VD._dll_CurrentVirtualDesktopChanged_IObjectArray

            this.IVirtualDesktopNotification_methods_count:=13
        } else if (buildNumber <= 22631 && revisionNumber < 3085) {
            ;from 22621.2215
            IID_IVirtualDesktopManagerInternal_str:="{4970ba3d-fd4e-4647-bea3-d89076ef4b9c}"
            IID_IVirtualDesktop_str:="{3f07f4be-b107-441a-af0f-39d82529072c}"
            ;IID_IVirtualDesktopNotification_str:="{b287fa1c-7771-471a-a2df-9b6b21f0d675}"
            this.IID_IVirtualDesktopNotification_n1:=5123538856297626140
            this.IID_IVirtualDesktopNotification_n2:=8491238173783613346

            idx_MoveViewToDesktop:=4 ;DllCall(ptr_MoveViewToDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IApplicationView,"Ptr",IVirtualDesktop)
            idx_GetCurrentDesktop:=6 ;DllCall(ptr_GetCurrentDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr*",&IVirtualDesktop:=0)
            idx_GetDesktops:=7 ;DllCall(ptr_GetDesktops,"Ptr",IVirtualDesktopManagerInternal,"Ptr*",&IObjectArray:=0)
            idx_CreateDesktop:=10 ;DllCall(ptr_CreateDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr*",&IVirtualDesktop:=0)
            idx_RemoveDesktop:=12 ;DllCall(ptr_RemoveDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",IVirtualDesktop_fallback)
            this._dll_MoveViewToDesktop:=this._dll_MoveViewToDesktop_normal
            this._dll_GetCurrentDesktop:=this._dll_GetCurrentDesktop_normal
            this._dll_GetDesktops:=this._dll_GetDesktops_normal
            this._dll_CreateDesktop:=this._dll_CreateDesktop_normal
            this._dll_RemoveDesktop:=this._dll_RemoveDesktop_normal

            this.idx_GetId:=4 ;DllCall(ptr_GetId,"Ptr",IVirtualDesktop,"Ptr",guid_buf)

            this.idx_VirtualDesktopNameChanged:=8 ;params (IVirtualDesktop, HSTRING)
            this._dll_VirtualDesktopNameChanged:=VD._dll_VirtualDesktopNameChanged_normal
            this.idx_SetDesktopName:=15 ;DllCall(ptr_SetDesktopName,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",HSTRING)
            this.idx_GetName:=5 ;DllCall(ptr_GetName,"Ptr",IVirtualDesktop,"Ptr*",HSTRING)

            this.idx_VirtualDesktopWallpaperChanged:=11 ;params (IVirtualDesktop, HSTRING)
            this._dll_VirtualDesktopWallpaperChanged:=VD._dll_VirtualDesktopWallpaperChanged_normal
            this.idx_SetDesktopWallpaper:=16 ;DllCall(ptr_SetDesktopWallpaper,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",HSTRING)
            this.idx_GetWallpaper:=6 ;DllCall(ptr_GetWallpaper,"Ptr",IVirtualDesktop,"Ptr*",HSTRING)

            this.idx_VirtualDesktopCreated:=3 ;params (IVirtualDesktop)
            this.idx_VirtualDesktopDestroyBegin:=4 ;params (IVirtualDesktop, IVirtualDesktop_fallback)
            this.idx_VirtualDesktopDestroyFailed:=5 ;params (IVirtualDesktop, IVirtualDesktop_fallback)
            this.idx_VirtualDesktopDestroyed:=6 ;params (IVirtualDesktop, IVirtualDesktop_fallback)
            this.idx_ViewVirtualDesktopChanged:=9 ;params (IApplicationView)
            this.idx_CurrentVirtualDesktopChanged:=10 ;params (IVirtualDesktop_old, IVirtualDesktop_new)
            this._dll_VirtualDesktopCreated:=VD._dll_VirtualDesktopCreated_normal
            this._dll_VirtualDesktopDestroyBegin:=VD._dll_VirtualDesktopDestroyBegin_normal
            this._dll_VirtualDesktopDestroyFailed:=VD._dll_VirtualDesktopDestroyFailed_normal
            this._dll_VirtualDesktopDestroyed:=VD._dll_VirtualDesktopDestroyed_normal
            this._dll_ViewVirtualDesktopChanged:=VD._dll_ViewVirtualDesktopChanged_normal
            this._dll_CurrentVirtualDesktopChanged:=VD._dll_CurrentVirtualDesktopChanged_normal

            this.IVirtualDesktopNotification_methods_count:=14
        } else {
            ;from 22631.3085
            ;the only difference with the above is IID_IVirtualDesktopNotification
            IID_IVirtualDesktopManagerInternal_str:="{4970ba3d-fd4e-4647-bea3-d89076ef4b9c}"
            IID_IVirtualDesktop_str:="{3f07f4be-b107-441a-af0f-39d82529072c}"
            ;IID_IVirtualDesktopNotification_str:="{b9e5e94d-233e-49ab-af5c-2b4541c3aade}"
            this.IID_IVirtualDesktopNotification_n1:=5308375338100058445
            this.IID_IVirtualDesktopNotification_n2:=-2401892766147978065

            idx_MoveViewToDesktop:=4 ;DllCall(ptr_MoveViewToDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IApplicationView,"Ptr",IVirtualDesktop)
            idx_GetCurrentDesktop:=6 ;DllCall(ptr_GetCurrentDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr*",&IVirtualDesktop:=0)
            idx_GetDesktops:=7 ;DllCall(ptr_GetDesktops,"Ptr",IVirtualDesktopManagerInternal,"Ptr*",&IObjectArray:=0)
            idx_CreateDesktop:=10 ;DllCall(ptr_CreateDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr*",&IVirtualDesktop:=0)
            idx_RemoveDesktop:=12 ;DllCall(ptr_RemoveDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",IVirtualDesktop_fallback)
            this._dll_MoveViewToDesktop:=this._dll_MoveViewToDesktop_normal
            this._dll_GetCurrentDesktop:=this._dll_GetCurrentDesktop_normal
            this._dll_GetDesktops:=this._dll_GetDesktops_normal
            this._dll_CreateDesktop:=this._dll_CreateDesktop_normal
            this._dll_RemoveDesktop:=this._dll_RemoveDesktop_normal

            this.idx_GetId:=4 ;DllCall(ptr_GetId,"Ptr",IVirtualDesktop,"Ptr",guid_buf)

            this.idx_VirtualDesktopNameChanged:=8 ;params (IVirtualDesktop, HSTRING)
            this._dll_VirtualDesktopNameChanged:=VD._dll_VirtualDesktopNameChanged_normal
            this.idx_SetDesktopName:=15 ;DllCall(ptr_SetDesktopName,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",HSTRING)
            this.idx_GetName:=5 ;DllCall(ptr_GetName,"Ptr",IVirtualDesktop,"Ptr*",HSTRING)

            this.idx_VirtualDesktopWallpaperChanged:=11 ;params (IVirtualDesktop, HSTRING)
            this._dll_VirtualDesktopWallpaperChanged:=VD._dll_VirtualDesktopWallpaperChanged_normal
            this.idx_SetDesktopWallpaper:=16 ;DllCall(ptr_SetDesktopWallpaper,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",HSTRING)
            this.idx_GetWallpaper:=6 ;DllCall(ptr_GetWallpaper,"Ptr",IVirtualDesktop,"Ptr*",HSTRING)

            this.idx_VirtualDesktopCreated:=3 ;params (IVirtualDesktop)
            this.idx_VirtualDesktopDestroyBegin:=4 ;params (IVirtualDesktop, IVirtualDesktop_fallback)
            this.idx_VirtualDesktopDestroyFailed:=5 ;params (IVirtualDesktop, IVirtualDesktop_fallback)
            this.idx_VirtualDesktopDestroyed:=6 ;params (IVirtualDesktop, IVirtualDesktop_fallback)
            this.idx_ViewVirtualDesktopChanged:=9 ;params (IApplicationView)
            this.idx_CurrentVirtualDesktopChanged:=10 ;params (IVirtualDesktop_old, IVirtualDesktop_new)
            this._dll_VirtualDesktopCreated:=VD._dll_VirtualDesktopCreated_normal
            this._dll_VirtualDesktopDestroyBegin:=VD._dll_VirtualDesktopDestroyBegin_normal
            this._dll_VirtualDesktopDestroyFailed:=VD._dll_VirtualDesktopDestroyFailed_normal
            this._dll_VirtualDesktopDestroyed:=VD._dll_VirtualDesktopDestroyed_normal
            this._dll_ViewVirtualDesktopChanged:=VD._dll_ViewVirtualDesktopChanged_normal
            this._dll_CurrentVirtualDesktopChanged:=VD._dll_CurrentVirtualDesktopChanged_normal

            this.IVirtualDesktopNotification_methods_count:=14
        }

        this.IVirtualDesktopManager := ComObjCreate("{aa509086-5ca9-4c25-8f95-589d3c07b48a}", "{a5cd92ff-29be-454c-8d04-d82879fb3f1b}")
        ;this._dll_GetViewForHwnd(VD_animation_gui_hwnd) always returns 0
        ;therefore, I can't replace MoveWindowToDesktop() with MoveViewToDesktop() and stop using IVirtualDesktopManager
        this.ptr_MoveWindowToDesktop := this._vtable(this.IVirtualDesktopManager, 5)

        this.CImmersiveShell_IServiceProvider := ComObjCreate("{c2f03a33-21f5-47fa-b4bb-156362a2f239}", "{6d5140c1-7436-11ce-8034-00aa006009fa}")

        this.IVirtualDesktopManagerInternal := ComObjQuery(this.CImmersiveShell_IServiceProvider, "{c5e0cdca-7b6e-41b2-9fc4-d93975cc467b}", IID_IVirtualDesktopManagerInternal_str)
        this.IVirtualDesktopPinnedApps := ComObjQuery(this.CImmersiveShell_IServiceProvider, "{b5a399e7-1c87-46b8-88e9-fc5747b171bd}", "{4ce81583-1e4c-4632-a621-07a53543148f}")
        this.IApplicationViewCollection := ComObjQuery(this.CImmersiveShell_IServiceProvider,"{1841c6d7-4f9d-42c0-af41-8747538f10e5}","{1841c6d7-4f9d-42c0-af41-8747538f10e5}" )

        this.ptr_MoveViewToDesktop := this._vtable(this.IVirtualDesktopManagerInternal, idx_MoveViewToDesktop)
        this.ptr_GetCurrentDesktop := this._vtable(this.IVirtualDesktopManagerInternal, idx_GetCurrentDesktop)
        this.ptr_GetDesktops := this._vtable(this.IVirtualDesktopManagerInternal, idx_GetDesktops)
        this.ptr_CreateDesktop := this._vtable(this.IVirtualDesktopManagerInternal, idx_CreateDesktop)
        this.ptr_RemoveDesktop := this._vtable(this.IVirtualDesktopManagerInternal, idx_RemoveDesktop)
        if (this.idx_SetDesktopName > -1) {
            this.ptr_SetDesktopName := this._vtable(this.IVirtualDesktopManagerInternal, this.idx_SetDesktopName)
        }
        if (this.idx_SetDesktopWallpaper > -1) {
            this.ptr_SetDesktopWallpaper := this._vtable(this.IVirtualDesktopManagerInternal, this.idx_SetDesktopWallpaper)
        }

        this.ptr_IsViewPinned := this._vtable(this.IVirtualDesktopPinnedApps, 6) ;DllCall(ptr_IsViewPinned,"Ptr",IVirtualDesktopPinnedApps,"Ptr",IApplicationView,"Int*",&viewIsPinned:=0)
        this.ptr_PinView := this._vtable(this.IVirtualDesktopPinnedApps, 7) ;DllCall(ptr_PinView,"Ptr",IVirtualDesktopPinnedApps,"Ptr",IApplicationView)
        this.ptr_UnpinView := this._vtable(this.IVirtualDesktopPinnedApps, 8) ;DllCall(ptr_UnpinView,"Ptr",IVirtualDesktopPinnedApps,"Ptr",IApplicationView)

        this.ptr_GetViewForHwnd := this._vtable(this.IApplicationViewCollection, 6) ;DllCall(ptr_GetViewForHwnd,"Ptr",IApplicationViewCollection,"Ptr",HWND,"Ptr*",&IApplicationView:=0)

        ;----------------------

        ; VarSetCapacity(IID_IVirtualDesktop, 16)
        ; this will never be garbage collected
        this.IID_IVirtualDesktop_ptr := DllCall("GlobalAlloc","UInt",0x00,"Uint",16,"Ptr")
        DllCall("ole32\CLSIDFromString","Str",IID_IVirtualDesktop_str,"Ptr",this.IID_IVirtualDesktop_ptr)

        ;----------------------

        this.savedLocalizedWord_Desktop:=false
    }
    ;dll methods start
    _dll_MoveViewToDesktop_normal(IApplicationView,IVirtualDesktop) {
        DllCall(this.ptr_MoveViewToDesktop,"Ptr",this.IVirtualDesktopManagerInternal,"Ptr",IApplicationView,"Ptr",IVirtualDesktop)
    }
    _dll_GetCurrentDesktop_normal() {
        DllCall(this.ptr_GetCurrentDesktop,"Ptr",this.IVirtualDesktopManagerInternal,"Ptr*",IVirtualDesktop_currentDesktop:=0)
        return IVirtualDesktop_currentDesktop
    }
    _dll_GetDesktops_normal() {
        DllCall(this.ptr_GetDesktops,"Ptr",this.IVirtualDesktopManagerInternal,"Ptr*",IObjectArray:=0)
        return IObjectArray
    }
    _dll_CreateDesktop_normal() {
        DllCall(this.ptr_CreateDesktop,"Ptr",this.IVirtualDesktopManagerInternal,"Ptr*",IVirtualDesktop_created:=0)
        return IVirtualDesktop_created
    }
    _dll_RemoveDesktop_normal(IVirtualDesktop,IVirtualDesktop_fallback) {
        DllCall(this.ptr_RemoveDesktop,"Ptr",this.IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",IVirtualDesktop_fallback)
    }
    _dll_GetCurrentDesktop_HMONITOR() {
        DllCall(this.ptr_GetCurrentDesktop,"Ptr",this.IVirtualDesktopManagerInternal,"Ptr",0,"Ptr*",IVirtualDesktop_currentDesktop:=0)
        return IVirtualDesktop_currentDesktop
    }
    _dll_GetDesktops_HMONITOR() {
        DllCall(this.ptr_GetDesktops,"Ptr",this.IVirtualDesktopManagerInternal,"Ptr",0,"Ptr*",IObjectArray:=0)
        return IObjectArray
    }
    _dll_CreateDesktop_HMONITOR() {
        DllCall(this.ptr_CreateDesktop,"Ptr",this.IVirtualDesktopManagerInternal,"Ptr",0,"Ptr*",IVirtualDesktop_created:=0)
        return IVirtualDesktop_created
    }

    _dll_IsViewPinned(IApplicationView) {
        DllCall(this.ptr_IsViewPinned,"Ptr",this.IVirtualDesktopPinnedApps,"Ptr",IApplicationView,"Int*",viewIsPinned:=0)
        return viewIsPinned
    }
    _dll_PinView(IApplicationView) {
        DllCall(this.ptr_PinView,"Ptr",this.IVirtualDesktopPinnedApps,"Ptr",IApplicationView)
    }
    _dll_UnpinView(IApplicationView) {
        DllCall(this.ptr_UnpinView,"Ptr",this.IVirtualDesktopPinnedApps,"Ptr",IApplicationView)
    }

    _dll_GetViewForHwnd(HWND) {
        DllCall(this.ptr_GetViewForHwnd,"Ptr",this.IApplicationViewCollection,"Ptr",HWND,"Ptr*",IApplicationView:=0)
        return IApplicationView
    }
    ;dll methods end

    ;actual methods start
    getCount() { ;how many virtual desktops you now have
        return this._GetDesktops_Obj().GetCount()
    }

    goToDesktopNum(desktopNum) { ; Lej77 https://github.com/Grabacr07/VirtualDesktop/pull/23#issuecomment-334918711
        firstWindowId:=this._getFirstWindowInVD(desktopNum)

        Gui VD_animation_gui:New, % "-Border -SysMenu +Owner -Caption +HwndVD_animation_gui_hwnd_tmp"
        VD_animation_gui_hwnd:=VD_animation_gui_hwnd_tmp+0
        IVirtualDesktop := this._GetDesktops_Obj().GetAt(desktopNum)
        VarSetCapacity(GUID_Desktop, 16)
        ptr_GetId:=this._vtable(IVirtualDesktop, this.idx_GetId)
        DllCall(ptr_GetId,"Ptr",IVirtualDesktop,"Ptr",&GUID_Desktop)
        ;this._dll_GetViewForHwnd(VD_animation_gui_hwnd) always returns 0
        ;therefore, I can't replace MoveWindowToDesktop() with MoveViewToDesktop() and stop using IVirtualDesktopManager
        DllCall(this.ptr_MoveWindowToDesktop, "Ptr", this.IVirtualDesktopManager, "Ptr", VD_animation_gui_hwnd, "Ptr", &GUID_Desktop)
        DllCall("ShowWindow","Ptr",VD_animation_gui_hwnd,"Int",4) ;after gui on current desktop owned by current process became active window, Show gui on different desktop owned by current process
        this.SetForegroundWindow(VD_animation_gui_hwnd)
        loop 20 {
            if (this.getCurrentDesktopNum()==desktopNum) { ; wildest hack ever..
                if (firstWindowId) {
                    DllCall("SetForegroundWindow","Ptr",firstWindowId)
                } else {
                    this._activateDesktopBackground()
                }
                break
            }
            Sleep 25
        }
        Gui VD_animation_gui:Destroy

    }

    _getLocalizedWord_Desktop() {
        if (this.savedLocalizedWord_Desktop) {
            return this.savedLocalizedWord_Desktop
        }

        hModule := DllCall("GetModuleHandle", "Str","shell32.dll", "Ptr") ;ahk always loads "shell32.dll"
        length:=DllCall("LoadString", "Uint",hModule, "Uint",21769, "Ptr*",lpBuffer, "Int",0) ;21769="Desktop"
        this.savedLocalizedWord_Desktop := StrGet(lpBuffer, length, "UTF-16")
        return this.savedLocalizedWord_Desktop
    }
    getNameFromDesktopNum(desktopNum) {
        desktopName:=""
        if (desktopNum > 0 && this.idx_GetName > -1) {
            IVirtualDesktop:=this._GetDesktops_Obj().GetAt(desktopNum)
            if (IVirtualDesktop) {
                ptr_GetName:=this._vtable(IVirtualDesktop, this.idx_GetName)
                DllCall(ptr_GetName,"Ptr",IVirtualDesktop,"Ptr*",HSTRING:=0)
                desktopName:=StrGet(DllCall("combase\WindowsGetStringRawBuffer","Ptr",HSTRING,"Uint*",length:=0,"Ptr"),"UTF-16")
                DllCall("combase\WindowsDeleteString","Ptr",HSTRING)
            }
        }
        if (!desktopName) {
            desktopName:=this._getLocalizedWord_Desktop() " " desktopNum
        }

        return desktopName
    }
    setNameToDesktopNum(desktopName,desktopNum) {
        if (desktopNum > 0 && this.idx_SetDesktopName > -1) {
            IVirtualDesktop:=this._GetDesktops_Obj().GetAt(desktopNum)
            if (IVirtualDesktop) {
                DllCall("combase\WindowsCreateString","WStr",desktopName,"Uint",StrLen(desktopName),"Ptr*",HSTRING:=0)
                DllCall(this.ptr_SetDesktopName,"Ptr",this.IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",HSTRING)
                DllCall("combase\WindowsDeleteString","Ptr",HSTRING)
            }
        }
    }
    getWallpaperFromDesktopNum(desktopNum) {
        desktopWallpaper:=""
        if (desktopNum > 0 && this.idx_GetWallpaper > -1) {
            IVirtualDesktop:=this._GetDesktops_Obj().GetAt(desktopNum)
            if (IVirtualDesktop) {
                ptr_GetWallpaper:=this._vtable(IVirtualDesktop, this.idx_GetWallpaper)
                DllCall(ptr_GetWallpaper,"Ptr",IVirtualDesktop,"Ptr*",HSTRING:=0)
                desktopWallpaper:=StrGet(DllCall("combase\WindowsGetStringRawBuffer","Ptr",HSTRING,"Uint*",length:=0,"Ptr"),"UTF-16")
                DllCall("combase\WindowsDeleteString","Ptr",HSTRING)
            }
        }
        return desktopWallpaper
    }
    setWallpaperToDesktopNum(desktopWallpaper,desktopNum) {
        if (desktopNum > 0 && this.idx_SetDesktopWallpaper > -1) {
            IVirtualDesktop:=this._GetDesktops_Obj().GetAt(desktopNum)
            if (IVirtualDesktop) {
                DllCall("combase\WindowsCreateString","WStr",desktopWallpaper,"Uint",StrLen(desktopWallpaper),"Ptr*",HSTRING:=0)
                DllCall(this.ptr_SetDesktopWallpaper,"Ptr",this.IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",HSTRING)
                DllCall("combase\WindowsDeleteString","Ptr",HSTRING)
            }
        }
    }

    getDesktopNumOfWindow(wintitle) {
        found:=this._tryGetValidWindow(wintitle)
        if (!found) {
            return -1 ;for false
        }
        thePView:=found[2]

        desktopNum_ofWindow:=this._desktopNum_from_pView(thePView)
        return desktopNum_ofWindow ; 0 for "Show on all desktops"
    }

    goToDesktopOfWindow(wintitle, activateYourWindow:=true) {
        found:=this._tryGetValidWindow(wintitle)
        if (!found) {
            return -1 ;for false
        }
        theHwnd:=found[1]
        thePView:=found[2]

        desktopNum_ofWindow:=this._desktopNum_from_pView(thePView)
        this.goToDesktopNum(desktopNum_ofWindow)

        if (activateYourWindow) {
            WinActivate, ahk_id %theHwnd%
        }
    }

    MoveWindowToDesktopNum(wintitle, desktopNum) {
        found:=this._tryGetValidWindow(wintitle)
        if (!found) {
            return -1 ;for false
        }
        theHwnd:=found[1]
        thePView:=found[2]

        needActivateWindowUnder:=false
        if (activeHwnd:=WinExist("A")) {
            if (activeHwnd==theHwnd) {
                currentDesktopNum:=this.getCurrentDesktopNum()
                if (!(currentDesktopNum==desktopNum)) {
                    needActivateWindowUnder:=true
                }
            }
        }

        IVirtualDesktop:=this._GetDesktops_Obj().GetAt(desktopNum)
        this._dll_MoveViewToDesktop(thePView,IVirtualDesktop)

        if (needActivateWindowUnder) {
            firstWindowId:=this._getFirstWindowInVD(currentDesktopNum, theHwnd)
            if (firstWindowId) {
                this.SetForegroundWindow(firstWindowId)
            } else {
                this._activateDesktopBackground()
            }
        }

    }

    getRelativeDesktopNum(anchor_desktopNum, relative_count) {
        Desktops_Obj:=this._GetDesktops_Obj()
        count_Desktops:=Desktops_Obj.GetCount()

        absolute_desktopNum:=anchor_desktopNum + relative_count
        ;// The 1-based indices wrap around on the first and last desktop.
        ;// say count_Desktops:=3
        absolute_desktopNum:=Mod(absolute_desktopNum, count_Desktops)
        ; 4 -> 1
        if (absolute_desktopNum <= 0) {
            ; 0 -> 3
            absolute_desktopNum:=absolute_desktopNum + count_Desktops
        }

        return absolute_desktopNum
    }

    MoveWindowToRelativeDesktopNum(wintitle, relative_count) {

        desktopNum_ofWindow := this.getDesktopNumOfWindow(wintitle)
        absolute_desktopNum := this.getRelativeDesktopNum(desktopNum_ofWindow, relative_count)

        this.MoveWindowToDesktopNum(wintitle, absolute_desktopNum)

        return absolute_desktopNum
    }

    gotoRelativeDesktopNum(relative_count) {
        this.goToDesktopNum(this.getRelativeDesktopNum(this.getCurrentDesktopNum(), relative_count))
    }

    MoveWindowToCurrentDesktop(wintitle, activateYourWindow:=true) {
        found:=this._tryGetValidWindow(wintitle)
        if (!found) {
            return -1 ;for false
        }
        theHwnd:=found[1]
        thePView:=found[2]

        currentDesktopNum:=this.getCurrentDesktopNum()
        IVirtualDesktop:=this._GetDesktops_Obj().GetAt(currentDesktopNum)

        this._dll_MoveViewToDesktop(thePView,IVirtualDesktop)

        if (activateYourWindow) {
            WinActivate % "ahk_id " theHwnd
        }
    }

    getCurrentDesktopNum() {
        IVirtualDesktop_ofCurrentDesktop:=this._dll_GetCurrentDesktop()

        desktopNum:=this._desktopNum_from_IVirtualDesktop(IVirtualDesktop_ofCurrentDesktop)
        return desktopNum
    }

    createDesktop(goThere:=false) {
        IVirtualDesktop_ofNewDesktop:=this._dll_CreateDesktop()

        if (goThere) {
            ;we could assume that it's the rightmost desktop:
            ; desktopNum:=this.getCount()
            ;but I'm not risking it
            desktopNum:=this._desktopNum_from_IVirtualDesktop(IVirtualDesktop_ofNewDesktop)
            this.goToDesktopNum(desktopNum)
        }
    }

    createUntil(howMany, goToLastlyCreated:=false) {
        howManyThereAlreadyAre:=this.getCount()
        if (howManyThereAlreadyAre>=howMany) {
            return
        }

        ;this will create until one less than wanted
        loop % howMany - howManyThereAlreadyAre - 1 {
            this.createDesktop(false)
        }
        this.createDesktop(goToLastlyCreated)
    }

    removeDesktop(desktopNum, fallback_desktopNum:=false) {
        ;FALLBACK IS ONLY USED IF YOU ARE CURRENTLY ON THE VD BEING DELETED
        ;but we NEED a fallback, regardless, so I'm not checking if you are currently on the vd being deleted

        Desktops_Obj:=this._GetDesktops_Obj()

        ;if no fallback,
        if (!fallback_desktopNum) {

            ;look left
            if (desktopNum > 1) {
                fallback_desktopNum:=desktopNum - 1
            }
            ;look right
            else if (desktopNum < Desktops_Obj.GetCount()) {
                fallback_desktopNum:=desktopNum + 1
            }
            ;no fallback to go to
            else {
                return false
            }
        }

        IVirtualDesktop:=Desktops_Obj.GetAt(desktopNum)
        IVirtualDesktop_fallback:=Desktops_Obj.GetAt(fallback_desktopNum)

        this._dll_RemoveDesktop(IVirtualDesktop,IVirtualDesktop_fallback)
    }

    IsWindowPinned(wintitle) {
        found:=this._tryGetValidWindow(wintitle)
        if (!found) {
            return -1 ;for false
        }
        thePView:=found[2]

        viewIsPinned:=this._dll_IsViewPinned(thePView)
        return viewIsPinned
    }
    TogglePinWindow(wintitle) {
        found:=this._tryGetValidWindow(wintitle)
        if (!found) {
            return -1 ;for false
        }
        thePView:=found[2]

        viewIsPinned:=this._dll_IsViewPinned(thePView)
        if (viewIsPinned) {
            this._dll_UnPinView(thePView)
        } else {
            this._dll_PinView(thePView)
        }

    }
    PinWindow(wintitle) {
        found:=this._tryGetValidWindow(wintitle)
        if (!found) {
            return -1 ;for false
        }
        thePView:=found[2]

        this._dll_PinView(thePView)
    }
    UnPinWindow(wintitle) {
        found:=this._tryGetValidWindow(wintitle)
        if (!found) {
            return -1 ;for false
        }
        thePView:=found[2]

        this._dll_UnPinView(thePView)
    }

    ; COM class start ;https://github.com/Ciantic/VirtualDesktopAccessor/blob/5bc1bbaab247b5d72e70abc9432a15275fd2d229/VirtualDesktopAccessor/dllmain.h#L718-L794
    _dll_QueryInterface_Same(riid, ppvObject) {
        if (!ppvObject) {
            return 0x80070057 ;E_INVALIDARG
        }

        ;IID_IUnknown_str:="{00000000-0000-0000-C000-000000000046}"
        ;IID_IUnknown_n1:=0
        ;IID_IUnknown_n2:=5044031582654955712

        if ((NumGet(riid+0,0x0,"Int64")==0 && NumGet(riid+0,0x8,"Int64")==5044031582654955712) || (NumGet(riid+0,0x0,"Int64")==VD.IID_IVirtualDesktopNotification_n1 && NumGet(riid+0,0x8,"Int64")==VD.IID_IVirtualDesktopNotification_n2)) {
            NumPut(this, ppvObject+0, 0, "Ptr")
            VD._dll_AddRef_Same.Call(this)
            return 0 ;S_OK
        }
        ; *ppvObject = NULL;
        NumPut(0, ppvObject+0, "Ptr")
        return 0x80004002 ;E_NOINTERFACE

        ; // Always set out parameter to NULL, validating it first.
        ; if (!ppvObject)
            ; return E_INVALIDARG;
        ; *ppvObject = NULL;

        ; if (riid == IID_IUnknown || riid == IID_IVirtualDesktopNotification)
        ; {
            ; // Increment the reference count and return the pointer.
            ; *ppvObject = (LPVOID)this;
            ; AddRef();
            ; return S_OK;
        ; }
        ; return E_NOINTERFACE;
    }
    _dll_AddRef_Same() {
        refCount:=NumGet(this+0, A_PtrSize, "UInt")
        refCount++
        NumPut(refCount, this+0, A_PtrSize, "UInt")
        ; NumPut(this + 4)
        ; refCount:=

        ; return InterlockedIncrement(&_referenceCount);
        return refCount
    }
    _dll_Release_Same() {
        refCount:=NumGet(this+0, A_PtrSize, "UInt")
        refCount--
        NumPut(refCount, this+0, A_PtrSize, "UInt")
        ; ULONG result = InterlockedDecrement(&_referenceCount);
        ; if (result == 0)
        ; {
            ; delete this;
        ; }
        return refCount
    }
    _dll_VirtualDesktopCreated_normal(IVirtualDesktop_created) {
        VD.VirtualDesktopCreated.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_created))
        return 0 ;S_OK
    }
    _dll_VirtualDesktopDestroyBegin_normal(IVirtualDesktop_destroyed, IVirtualDesktop_fallback) {
        VD.VirtualDesktopDestroyBegin.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_destroyed), VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_fallback))
        return 0 ;S_OK
    }
    _dll_VirtualDesktopDestroyFailed_normal(IVirtualDesktop_destroyed, IVirtualDesktop_fallback) {
        VD.VirtualDesktopDestroyFailed.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_destroyed), VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_fallback))
        return 0 ;S_OK
    }
    _dll_VirtualDesktopDestroyed_normal(IVirtualDesktop_destroyed, IVirtualDesktop_fallback) {
        VD.VirtualDesktopDestroyed.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_destroyed), VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_fallback))
        return 0 ;S_OK
    }
    _dll_ViewVirtualDesktopChanged_normal(IApplicationView) {
        VD.ViewVirtualDesktopChanged.Call(IApplicationView)
        return 0 ;S_OK
    }
    _dll_CurrentVirtualDesktopChanged_normal(IVirtualDesktop_old, IVirtualDesktop_new) {
        VD.CurrentVirtualDesktopChanged.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_old), VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_new))
        return 0 ;S_OK
    }
    _dll_VirtualDesktopNameChanged_normal(IVirtualDesktop, HSTRING) {
        desktopName:=StrGet(DllCall("combase\WindowsGetStringRawBuffer","Ptr",HSTRING,"Uint*",length:=0,"Ptr"),"UTF-16")
        DllCall("combase\WindowsDeleteString","Ptr",HSTRING)
        VD.VirtualDesktopNameChanged.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop), desktopName)
        return 0 ;S_OK
    }
    _dll_VirtualDesktopWallpaperChanged_normal(IVirtualDesktop, HSTRING) {
        desktopName:=StrGet(DllCall("combase\WindowsGetStringRawBuffer","Ptr",HSTRING,"Uint*",length:=0,"Ptr"),"UTF-16")
        DllCall("combase\WindowsDeleteString","Ptr",HSTRING)
        VD.VirtualDesktopWallpaperChanged.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop), desktopName)
        return 0 ;S_OK
    }
    _dll_VirtualDesktopCreated_IObjectArray(IObjectArray, IVirtualDesktop_created) {
        VD.VirtualDesktopCreated.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_created))
        return 0 ;S_OK
    }
    _dll_VirtualDesktopDestroyBegin_IObjectArray(IObjectArray, IVirtualDesktop_destroyed, IVirtualDesktop_fallback) {
        VD.VirtualDesktopDestroyBegin.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_destroyed), VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_fallback))
        return 0 ;S_OK
    }
    _dll_VirtualDesktopDestroyFailed_IObjectArray(IObjectArray, IVirtualDesktop_destroyed, IVirtualDesktop_fallback) {
        VD.VirtualDesktopDestroyFailed.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_destroyed), VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_fallback))
        return 0 ;S_OK
    }
    _dll_VirtualDesktopDestroyed_IObjectArray(IObjectArray, IVirtualDesktop_destroyed, IVirtualDesktop_fallback) {
        VD.VirtualDesktopDestroyed.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_destroyed), VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_fallback))
        return 0 ;S_OK
    }
    _dll_CurrentVirtualDesktopChanged_IObjectArray(IObjectArray, IVirtualDesktop_old, IVirtualDesktop_new) {
        VD.CurrentVirtualDesktopChanged.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_old), VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_new))
        return 0 ;S_OK
    }

    RegisterDesktopNotifications() { ;https://github.com/Ciantic/VirtualDesktopAccessor/blob/5bc1bbaab247b5d72e70abc9432a15275fd2d229/VirtualDesktopAccessor/dllmain.h#L718-L794
        methods_ptr:=DllCall("GlobalAlloc","Uint",0x40,"Uint",this.IVirtualDesktopNotification_methods_count*A_PtrSize) ;PLEASE DON'T GARBAGE COLLECT IT, this took me hours to debug, I was lucky ahkv2 garbage collected slowly
        NumPut(RegisterCallback(this._dll_QueryInterface_Same, "F"), methods_ptr+0*A_PtrSize, "Ptr")
        NumPut(RegisterCallback(this._dll_AddRef_Same, "F"), methods_ptr+1*A_PtrSize, "Ptr")
        NumPut(RegisterCallback(this._dll_Release_Same, "F"), methods_ptr+2*A_PtrSize, "Ptr")

        NumPut(RegisterCallback(this._dll_VirtualDesktopCreated, "F"), methods_ptr+this.idx_VirtualDesktopCreated*A_PtrSize, "Ptr")
        NumPut(RegisterCallback(this._dll_VirtualDesktopDestroyBegin, "F"), methods_ptr+this.idx_VirtualDesktopDestroyBegin*A_PtrSize, "Ptr")
        NumPut(RegisterCallback(this._dll_VirtualDesktopDestroyFailed, "F"), methods_ptr+this.idx_VirtualDesktopDestroyFailed*A_PtrSize, "Ptr")
        NumPut(RegisterCallback(this._dll_VirtualDesktopDestroyed, "F"), methods_ptr+this.idx_VirtualDesktopDestroyed*A_PtrSize, "Ptr")
        NumPut(RegisterCallback(this._dll_ViewVirtualDesktopChanged, "F"), methods_ptr+this.idx_ViewVirtualDesktopChanged*A_PtrSize, "Ptr")
        NumPut(RegisterCallback(this._dll_CurrentVirtualDesktopChanged, "F"), methods_ptr+this.idx_CurrentVirtualDesktopChanged*A_PtrSize, "Ptr")
        if (this.idx_VirtualDesktopNameChanged > -1) {
            NumPut(RegisterCallback(this._dll_VirtualDesktopNameChanged, "F"), methods_ptr+this.idx_VirtualDesktopNameChanged*A_PtrSize, "Ptr")
        }
        if (this.idx_VirtualDesktopWallpaperChanged > -1) {
            NumPut(RegisterCallback(this._dll_VirtualDesktopWallpaperChanged, "F"), methods_ptr+this.idx_VirtualDesktopWallpaperChanged*A_PtrSize, "Ptr")
        }

        ptr:=methods_ptr
        end:=methods_ptr + this.IVirtualDesktopNotification_methods_count*A_PtrSize
        callback_noop:=0
        while (ptr < end) {
            if (!NumGet(ptr+0,"Ptr")) {
                if (!callback_noop) {
                    callback_noop:=RegisterCallback("VD._No_Op", "F")
                }
                NumPut(callback_noop,ptr+0)
            }
            ptr+=A_PtrSize
        }

        this.RegisterDesktopNotifications_Same(methods_ptr)
    }
    _No_Op() {
    }
    RegisterDesktopNotifications_Same(methods_ptr) {
        obj:=DllCall("GlobalAlloc","Uint",0x00,"Uint",A_PtrSize + 4) ;PLEASE DON'T GARBAGE COLLECT IT, this took me hours to debug, I was lucky ahkv2 garbage collected slowly
        NumPut(methods_ptr, obj+0, 0, "Ptr")
        NumPut(0, obj+0, A_PtrSize, "UInt") ;refCount

        IDesktopNotificationService := ComObjQuery(this.CImmersiveShell_IServiceProvider, "{A501FDEC-4A09-464C-AE4E-1B9C21B84918}", "{0CD45E71-D927-4F15-8B0A-8FEF525337BF}")
        ptr_Register:=this._vtable(IDesktopNotificationService, 3)
        DllCall(ptr_Register,"Ptr",IDesktopNotificationService,"Ptr",obj,"Uint*",pdwCookie:=0)
        ; HRESULT hrNotificationService = pServiceProvider->QueryService(
		; CLSID_IVirtualNotificationService,
		; __uuidof(IVirtualDesktopNotificationService),
		; (PVOID*)&pDesktopNotificationService);
    }

    VirtualDesktopCreated(desktopNum:=0) {
    }
    VirtualDesktopDestroyBegin(desktopNum_Destroyed:=0, desktopNum_Fallback:=0) {
    }
    VirtualDesktopDestroyFailed(desktopNum_Destroyed:=0, desktopNum_Fallback:=0) {
    }
    VirtualDesktopDestroyed(desktopNum_Destroyed:=0, desktopNum_Fallback:=0) {
    }
    ViewVirtualDesktopChanged(pView:=0) {
    }
    CurrentVirtualDesktopChanged(desktopNum_Old:=0, desktopNum_New:=0) {
    }
    VirtualDesktopNameChanged(desktopNum:=0, desktopName:="") {
    }
    VirtualDesktopWallpaperChanged(desktopNum:=0, desktopWallpaper:="") {
    }

    ;actual methods end

    ;internal methods start

    SetForegroundWindow(hWnd) {
        if (DllCall("AllowSetForegroundWindow","Uint",DllCall("GetCurrentProcessId"))) {
            DllCall("SetForegroundWindow","Ptr",hwnd)
        } else {

            LCtrlDown:=GetKeyState("LCtrl")
            RCtrlDown:=GetKeyState("RCtrl")
            LShiftDown:=GetKeyState("LShift")
            RShiftDown:=GetKeyState("RShift")
            LWinDown:=GetKeyState("LWin")
            RWinDown:=GetKeyState("RWin")
            LAltDown:=GetKeyState("LAlt")
            RAltDown:=GetKeyState("RAlt")

            if ((LCtrlDown || RCtrlDown) && (LWinDown || RWinDown)) {
                toRelease:=""
                if (LShiftDown) {
                    toRelease.="{LShift Up}"
                }
                if (RShiftDown) {
                    toRelease.="{RShift Up}"
                }
                if (toRelease) {
                    Send % "{Blind}" toRelease
                }
            }
            Send % "{LAlt Down}{LAlt Down}" ;more stable than single on test: testf1_hotkey_hooked_Ctrl_Win_Alt.ah2
            ;Send "{LAlt Down}"
            DllCall("SetForegroundWindow","Ptr",hwnd)

            toAppend:=""
            if (!LAltDown) {
                toAppend.="{LAlt Up}"
            }
            if (RAltDown) {
                toAppend.="{RAlt Down}"
            }
            if (LCtrlDown) {
                toAppend.="{LCtrl Down}"
            }
            if (RCtrlDown) {
                toAppend.="{RCtrl Down}"
            }
            if (LShiftDown) {
                toAppend.="{LShift Down}"
            }
            if (RShiftDown) {
                toAppend.="{RShift Down}"
            }
            if (LWinDown) {
                toAppend.="{LWin Down}"
            }
            if (RWinDown) {
                toAppend.="{RWin Down}"
            }
            if (toAppend) {
                Send % "{Blind}" toAppend
            }
        }
    }

    _activateDesktopBackground() { ;this is really copying extremely long comments for short code like in AHK source code
        ; Win10:
        ; "FolderView ahk_class SysListView32 ahk_exe explorer.exe"
        ; "ahk_class SHELLDLL_DefView ahk_exe explorer.exe"
        ; "Program Manager ahk_class Progman ahk_exe explorer.exe" is the top level parent

        ; the parent parent of FolderView BECOMES "ahk_class WorkerW ahk_exe explorer.exe" after you press Win+Tab
        ; WorkerW doesn't exist before you press Win+Tab
        ; it's the same for Win11, Progman gets replaced by WorkerW, Progman still exists but isn't the parent of FolderView or top-level window that gets activated

        ; Q: if WinActivate Progman activates WorkerW(we want that) then what's the problem ?
        ; A: WinActivate will send {Alt down}{Alt up}{Alt down}{Alt up} if Progman is not activated : AHK source code: ((VK_MENU | 0x12 | ALT key)) https://github.com/AutoHotkey/AutoHotkey/blob/df84a3e902b522db0756a7366bd9884c80fa17b6/source/window.cpp#L260-L261
        ; the desktop background is correctly activated, we just don't want the extra Alt keys:
        ; if the hotkey is Ctrl+Shift+Win, and you add an Alt in there, Office 365 hotkey is triggered:
        ; https://github.com/FuPeiJiang/VD.ahk/issues/40#issuecomment-1548252485
        ; https://answers.microsoft.com/en-us/msoffice/forum/all/help-disabling-office-hotkey-of-ctrl-win-alt-shift/040ef6e5-8152-449b-849a-7494323101bb
        ; https://superuser.com/questions/1457073/how-do-i-disable-specific-windows-10-office-keyboard-shortcut-ctrlshiftwinal
        ; this is also bad because it prevents subsequent uses of the hotkey #!Right:: because {Alt up} releases Alt
        ; if (WinExist("ahk_class WorkerW ahk_exe explorer.exe")) {
        ;     WinActivate % "ahk_class WorkerW ahk_exe explorer.exe"
        ; } else {
        ;     WinActivate % "ahk_class Progman ahk_exe explorer.exe"
        ; }
        DllCall("SetForegroundWindow","Ptr",WinExist("ahk_class Progman ahk_exe explorer.exe"))
    }

    _getFirstWindowInVD(desktopNum, excludeHwnd:=0) {
        bak_DetectHiddenWindows:=A_DetectHiddenWindows
        DetectHiddenWindows, on
        returnValue:=0
        WinGet, outHwndList, List
        VarSetCapacity(GUID_Desktop, 16)
        IVirtualDesktop:=this._GetDesktops_Obj().GetAt(desktopNum)
        ptr_GetId:=this._vtable(IVirtualDesktop, this.idx_GetId)
        DllCall(ptr_GetId,"Ptr",IVirtualDesktop,"Ptr",&GUID_Desktop)
        n1:=NumGet(GUID_Desktop,0x0,"Int64")
        n2:=NumGet(GUID_Desktop,0x8,"Int64")
        loop % outHwndList {
            theHwnd:=outHwndList%A_Index%+0
            if (theHwnd==excludeHwnd) {
                continue
            }
            arr_success_pView_hWnd:=this._isValidWindow(theHwnd)
            if (arr_success_pView_hWnd[1]==0) {
                thePView:=arr_success_pView_hWnd[2]
                WinGet, OutputVar_MinMax, MinMax, % "ahk_id " theHwnd
                if (!(OutputVar_MinMax==-1)) { ;not Minimized

                    ptr_GetVirtualDesktopId:=this._vtable(thePView, 25)
                    DllCall(ptr_GetVirtualDesktopId,"Ptr",thePView,"Ptr",&GUID_Desktop)

                    if (n1==NumGet(GUID_Desktop,0x0,"Int64") && n2==NumGet(GUID_Desktop,0x8,"Int64")) {
                        ; WinActivate % "ahk_id " theHwnd
                        returnValue:=theHwnd
                        break
                    }
                }
            }
        }
        DetectHiddenWindows % bak_DetectHiddenWindows
        return returnValue
    }

    _tryGetValidWindow(wintitle) {
        bak_DetectHiddenWindows:=A_DetectHiddenWindows
        bak_TitleMatchMode:=A_TitleMatchMode
        DetectHiddenWindows, on
        SetTitleMatchMode, 2
        WinGet, outHwndList, List, % wintitle
        returnValue:=false
        loop % outHwndList {
            theHwnd:=outHwndList%A_Index%+0
            arr_success_pView_hWnd:=this._isValidWindow(theHwnd)
            pView:=arr_success_pView_hWnd[2]
            if (pView) {
                returnValue:=[arr_success_pView_hWnd[3], pView]
                break
            }
        }

        SetTitleMatchMode % bak_TitleMatchMode
        DetectHiddenWindows % bak_DetectHiddenWindows
        return returnValue
    }

    _desktopNum_from_IVirtualDesktop(IVirtualDesktop) {
        Desktops_Obj:=this._GetDesktops_Obj()
        Loop % Desktops_Obj.GetCount() {
            IVirtualDesktop_ofDesktop:=Desktops_Obj.GetAt(A_Index)

            if (IVirtualDesktop_ofDesktop == IVirtualDesktop) {
                return A_Index
            }
        }
        return 0 ;for "Show on all desktops"
    }

    _desktopNum_from_pView(thePView) {
        VarSetCapacity(GUID_Desktop, 16)
        ptr_GetVirtualDesktopId:=this._vtable(thePView, 25)
        DllCall(ptr_GetVirtualDesktopId,"Ptr",thePView,"Ptr",&GUID_Desktop)
        n1:=NumGet(GUID_Desktop,0x0,"Int64")
        n2:=NumGet(GUID_Desktop,0x8,"Int64")

        Desktops_Obj:=this._GetDesktops_Obj()
        Loop % Desktops_Obj.GetCount() {
            IVirtualDesktop:=Desktops_Obj.GetAt(A_Index)

            ptr_GetId:=this._vtable(IVirtualDesktop, this.idx_GetId)
            DllCall(ptr_GetId,"Ptr",IVirtualDesktop,"Ptr",&GUID_Desktop)

            if (n1==NumGet(GUID_Desktop,0x0,"Int64") && n2==NumGet(GUID_Desktop,0x8,"Int64")) {
                return A_Index
            }
        }
        return 0 ;for "Show on all desktops"
    }

    _GetDesktops_Obj() {
        IObjectArray:=this._dll_GetDesktops()
        return new this.IObjectArray_Wrapper(IObjectArray, this.IID_IVirtualDesktop_ptr)
    }

    ;internal methods end

    ;utility methods start

    _isValidWindow(hWnd,checkUpper:=true) { ;returns [0,pView,hWnd] if succeeded
        returnValue:=[1,0,0]
        breakToReturnFalse:
        loop 1 {
            dwStyle:=DllCall("GetWindowLongPtrW","Ptr",hWnd,"Int",-16,"Ptr")
            if (!(dwStyle & 0x10000000)) { ;0x10000000=WS_VISIBLE
                break breakToReturnFalse
            }
            dwExStyle:=DllCall("GetWindowLongPtrW","Ptr",hWnd,"Int",-20,"Ptr")
            if (!(dwExStyle&0x00040000)) { ;0x00040000=WS_EX_APPWINDOW
                if (dwExStyle&0x00000080 || dwExStyle&0x08000000) { ;0x00000080=WS_EX_TOOLWINDOW, 0x08000000=WS_EX_NOACTIVATE
                    break breakToReturnFalse
                }
                ; if any of ancestor is valid window, can't be valid window
                if (checkUpper) {
                    toCheck:=[]
                    upHwnd:=hWnd
                    while (upHwnd := DllCall("GetWindow","Ptr",upHwnd,"Uint",4)) { ;4=GW_OWNER
                        if (upHwnd==65552) {
                            break breakToReturnFalse
                        }
                        toCheck.Push(upHwnd)
                    }
                    i:=toCheck.Length() + 1
                    while (i-->1) { ;i goes to 1 (lmao)
                        arr_success_pView_hWnd:=this._isValidWindow(toCheck[i],false)
                        if (arr_success_pView_hWnd[1]==0) {
                            arr_success_pView_hWnd[1]:=2
                            returnValue:=arr_success_pView_hWnd
                            break breakToReturnFalse
                        }
                    }
                }
            }

            pView:=this._dll_GetViewForHwnd(hWnd)
            if (!pView) {
                break breakToReturnFalse
            }

            returnValue:=[0,pView,hWnd]
        }
        return returnValue
    }
    ;-------------------
    _vtable(ppv, index) {
        Return NumGet(NumGet(0+ppv)+A_PtrSize*index)
    }

    class IObjectArray_Wrapper {
        __New(IObjectArray, IID_Interface_ptr) {
            this.IObjectArray:=IObjectArray
            this.IID_Interface_ptr:=IID_Interface_ptr

            this.ptr_GetAt:=VD._vtable(IObjectArray,4)
        }
        __Delete() {
            ;IUnknown::Release
            ptr_Release:=VD._vtable(this.IObjectArray,2)
            DllCall(ptr_Release, "Ptr", this.IObjectArray)
        }
        GetAt(oneBasedIndex) {
            DllCall(this.ptr_GetAt, "Ptr", this.IObjectArray, "UInt", oneBasedIndex - 1, "Ptr", this.IID_Interface_ptr, "Ptr*", IInterface:=0)
            return IInterface
        }
        GetCount() {
            ptr_GetCount:=VD._vtable(this.IObjectArray,3)
            Count := 0
            DllCall(ptr_GetCount, "Ptr", this.IObjectArray, "UInt*", Count)
            return Count
        }

    }
    ;utility methods end

}
