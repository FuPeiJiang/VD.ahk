; must call VD_init() before any of these functions
; VD_getCurrentDesktop() ;this will return whichDesktop
; VD_getDesktopOfWindow(wintitle) ;this will return whichDesktop ;please use VD_goToDesktopOfWindow instead if you just want to go there.
; VD_getCount() ;this will return the number of virtual desktops you currently have
; VD_goToDesktop(whichDesktop)
; VD_goToDesktopOfWindow(wintitle, activate:=true)
; VD_sendToDesktop(wintitle,whichDesktop,followYourWindow:=true,activate:=true)
; VD_sendToCurrentDesktop(wintitle,activate:=true)

; VD_createDesktop(goThere:=true) ; VD_createUntil(howMany, goThere:=true)
; VD_removeDesktop(whichDesktop, fallback_which:=false)

; "Show this window on all desktops"
; VD_IsWindowPinned(wintitle)
; VD_TogglePinWindow(wintitle)
; VD_PinWindow(wintitle)
; VD_UnPinWindow(wintitle)

; "Show windows from this app on all desktops"
; VD_IsAppPinned(wintitle)
; VD_TogglePinApp(wintitle)
; VD_PinApp(wintitle)
; VD_UnPinApp(wintitle)

; internal functions
; VD_getCurrentIVirtualDesktop()
; VD_SwitchDesktop(IVirtualDesktop)
; VD_isValidWindow(hWnd)
; VD_getWintitle(hWnd)
; VD_IsWindow(hWnd){
; VD_vtable(ppv, idx)

; Thanks to:
; Blackholyman:
; https://www.autohotkey.com/boards/viewtopic.php?t=67642#p291160
; and
; Flipeador:
; https://www.autohotkey.com/boards/viewtopic.php?t=54202#p234192
; https://www.autohotkey.com/boards/viewtopic.php?t=54202#p234309

class VD {

    static dummyStatic1 := VD._init()

    init() { ;if you want to init early
        ; dummyStatic1 will be initiated and call _init()
    }

    _init()
    {
        splitByDot:=StrSplit(A_OSVersion, ".")
        buildNumber:=splitByDot[3]
        if (buildNumber < 22000)
        {
            ; Windows 10
            IID_IVirtualDesktopManagerInternal:="{F31574D6-B682-4CDC-BD56-1827860ABEC6}" ;https://github.com/MScholtes/VirtualDesktop/blob/812c321e286b82a10f8050755c94d21c4b69812f/VirtualDesktop.cs#L178-L191
            this._dll_GetDesktops:=this._dll_GetDesktops_Win10 ;conditionally assign method to method
            this._dll_SwitchDesktop:=this._dll_SwitchDesktop_Win10 ;conditionally assign method to method
        }
        else
        {
            ; Windows 11
            IID_IVirtualDesktopManagerInternal:="{B2F925B9-5A0F-4D2E-9F4D-2B1507593C10}" ;https://github.com/MScholtes/VirtualDesktop/blob/812c321e286b82a10f8050755c94d21c4b69812f/VirtualDesktop11.cs#L163-L185
            this._dll_GetDesktops:=this._dll_GetDesktops_Win11 ;conditionally assign method to method
            this._dll_SwitchDesktop:=this._dll_SwitchDesktop_Win11 ;conditionally assign method to method
        }

        VarSetCapacity(GUID_IID_IVirtualDesktop, 16)
        DllCall("Ole32.dll\CLSIDFromString", "Str", "{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}", "UPtr", &GUID_IID_IVirtualDesktop)
        this.Ptr_GUID_IID_IVirtualDesktop:=&GUID_IID_IVirtualDesktop

        this.IVirtualDesktopManager := ComObjCreate("{AA509086-5CA9-4C25-8F95-589D3C07B48A}", "{A5CD92FF-29BE-454C-8D04-D82879FB3F1B}")
        this.GetWindowDesktopId := VD_vtable(this.IVirtualDesktopManager, 4)

        IServiceProvider := ComObjCreate("{C2F03A33-21F5-47FA-B4BB-156362A2F239}", "{6D5140C1-7436-11CE-8034-00AA006009FA}")

        ;https://github.com/MScholtes/VirtualDesktop/blob/812c321e286b82a10f8050755c94d21c4b69812f/VirtualDesktop.cs#L178-L191
        this.IVirtualDesktopManagerInternal := ComObjQuery(IServiceProvider, "{C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B}", IID_IVirtualDesktopManagerInternal)
        ; this.GetCount := VD_vtable(this.IVirtualDesktopManagerInternal, 3 ; int GetCount();
        this.MoveViewToDesktop := VD_vtable(this.IVirtualDesktopManagerInternal, 4) ; void MoveViewToDesktop(object pView, IVirtualDesktop desktop);
        this.GetCurrentDesktop := VD_vtable(this.IVirtualDesktopManagerInternal, 6) ; IVirtualDesktop GetCurrentDesktop();
        this.CanViewMoveDesktops := VD_vtable(this.IVirtualDesktopManagerInternal, 5) ; bool CanViewMoveDesktops(object pView);
        this.GetDesktops := VD_vtable(this.IVirtualDesktopManagerInternal, 7) ; IObjectArray GetDesktops();
        this.GetAdjacentDesktop := VD_vtable(this.IVirtualDesktopManagerInternal, 8) ; int GetAdjacentDesktop(IVirtualDesktop from, int direction, out IVirtualDesktop desktop);
        this.SwitchDesktop := VD_vtable(this.IVirtualDesktopManagerInternal, 9) ; void SwitchDesktop(IVirtualDesktop desktop);
        this.CreateDesktop := VD_vtable(this.IVirtualDesktopManagerInternal, 10) ; IVirtualDesktop CreateDesktop();
        this.RemoveDesktop := VD_vtable(this.IVirtualDesktopManagerInternal, 11) ; void RemoveDesktop(IVirtualDesktop desktop, IVirtualDesktop fallback);
        ; this.FindDesktop := VD_vtable(this.IVirtualDesktopManagerInternal, 12) ; IVirtualDesktop FindDesktop(ref Guid desktopid);

        ;https://github.com/MScholtes/VirtualDesktop/blob/812c321e286b82a10f8050755c94d21c4b69812f/VirtualDesktop.cs#L225-L234
        this.IVirtualDesktopPinnedApps := ComObjQuery(IServiceProvider, "{B5A399E7-1C87-46B8-88E9-FC5747B171BD}", "{4CE81583-1E4C-4632-A621-07A53543148F}")
        this.IsAppIdPinned := VD_vtable(this.IVirtualDesktopPinnedApps, 3) ; bool IsAppIdPinned(string appId);
        this.PinAppID := VD_vtable(this.IVirtualDesktopPinnedApps, 4) ; void PinAppID(string appId);
        this.UnpinAppID := VD_vtable(this.IVirtualDesktopPinnedApps, 5) ; void UnpinAppID(string appId);
        this.IsViewPinned := VD_vtable(this.IVirtualDesktopPinnedApps, 6) ; bool IsViewPinned(IApplicationView applicationView);
        this.PinView := VD_vtable(this.IVirtualDesktopPinnedApps, 7) ; void PinView(IApplicationView applicationView);
        this.UnpinView := VD_vtable(this.IVirtualDesktopPinnedApps, 8) ; void UnpinView(IApplicationView applicationView);


        ImmersiveShell := ComObjCreate("{C2F03A33-21F5-47FA-B4BB-156362A2F239}", "{00000000-0000-0000-C000-000000000046}")
        ; if !(IApplicationViewCollection := ComObjQuery(ImmersiveShell,"{1841C6D7-4F9D-42C0-AF41-8747538F10E5}" ) ) ; doesn't work
        ; SAME CLSID and IID ?
        ; wait it's not CLSID:
        ; SID
        ; A service identifier in the same form as IID. When omitting this parameter, also omit the comma.
        this.IApplicationViewCollection := ComObjQuery(ImmersiveShell,"{1841C6D7-4F9D-42C0-AF41-8747538F10E5}","{1841C6D7-4F9D-42C0-AF41-8747538F10E5}" )
        if (!this.IApplicationViewCollection) ; 1607-1809
        {
            MsgBox IApplicationViewCollection interface not supported.
        }
        this.GetViewForHwnd := VD_vtable(IApplicationViewCollection, 6) ; (IntPtr hwnd, out IApplicationView view);
    }
    ;dll methods start
    _dll_GetDesktops_Win10() {
        IObjectArray := 0
        DllCall(this.GetDesktops, "UPtr", this.IVirtualDesktopManagerInternal, "UPtr*", IObjectArray)
        return IObjectArray
    }
    _dll_GetDesktops_Win11() {
        IObjectArray := 0
        DllCall(this.GetDesktops, "UPtr", this.IVirtualDesktopManagerInternal, "Ptr", 0, "UPtr*", IObjectArray)
        return IObjectArray
    }
    _dll_SwitchDesktop_Win10(IVirtualDesktop) {
        DllCall(this.SwitchDesktop, "ptr", this.IVirtualDesktopManagerInternal, "UPtr", IVirtualDesktop)
    }
    _dll_SwitchDesktop_Win11(IVirtualDesktop) {
        DllCall(this.SwitchDesktop, "ptr", this.IVirtualDesktopManagerInternal, "Ptr", 0, "UPtr", IVirtualDesktop)
    }
    ;dll methods end

    ;actual methods start
    getCount()
    {
        IObjectArray:=this._dll_GetDesktops()

        ; IObjectArray::GetCount ;https://github.com/MScholtes/VirtualDesktop/blob/812c321e286b82a10f8050755c94d21c4b69812f/VirtualDesktop.cs#L239-L243
        GetCount:=VD_vtable(IObjectArray,3)

        vd_Count := 0
        DllCall(GetCount, "UPtr", IObjectArray, "UInt*", vd_Count)
        return vd_Count
    }

    goToDesktop(desktopNum) {
        IVirtualDesktop:=this._IVirtualDesktop_from_desktopNum(desktopNum)
        this._SwitchDesktop(IVirtualDesktop)

        if (this._isWindowFullScreen("A"))
            timerFunc := ObjBindMethod(this, "_pleaseSwitchDesktop", desktopNum) ;https://www.autohotkey.com/docs/commands/SetTimer.htm#ExampleClass
            SetTimer % timerFunc, -50

    }
    ;actual methods end

    ;internal methods start
    _IVirtualDesktop_from_desktopNum(desktopNum) {
        IObjectArray:=this._dll_GetDesktops()

        GetAt:=VD_vtable(IObjectArray,4)
        DllCall(GetAt, "UPtr", IObjectArray, "UInt", desktopNum - 1, "UPtr", this.Ptr_GUID_IID_IVirtualDesktop, "UPtr*", IVirtualDesktop)
        return IVirtualDesktop
    }

    _SwitchDesktop(IVirtualDesktop) {
        ;activate taskbar before
        WinActivate, ahk_class Shell_TrayWnd
        WinWaitActive, ahk_class Shell_TrayWnd
        this._dll_SwitchDesktop(IVirtualDesktop)
        this._dll_SwitchDesktop(IVirtualDesktop)
        WinMinimize, ahk_class Shell_TrayWnd
    }

    _pleaseSwitchDesktop(desktopNum) {
        ;IVirtualDesktop should be calculated again because IVirtualDesktop could have changed
        ;what we want is the same desktopNum
        IVirtualDesktop:=this.IVirtualDesktop_from_desktopNum(desktopNum)
        this._SwitchDesktop(desktopNum)
        ;this method is goToDesktop(), but without the recursion, to prevent recursion
    }
    ;internal methods end

    ;utility methods start
    _isWindowFullScreen( winTitle ) {
        ;checks if the specified window is full screen

        winID := WinExist( winTitle )

        If ( !winID )
            Return false

        WinGet style, Style, ahk_id %WinID%
        WinGetPos ,,,winW,winH, %winTitle%
        ; 0x800000 is WS_BORDER.
        ; 0x20000000 is WS_MINIMIZE.
        ; no border and not minimized
        Return ((style & 0x20800000) or winH < A_ScreenHeight or winW < A_ScreenWidth) ? false : true
    }
    ;utility methods end



}


VD_getCurrentIVirtualDesktop()
{
    global GetCurrentDesktop, IVirtualDesktopManagerInternal
    CurrentIVirtualDesktop := 0
    DllCall(GetCurrentDesktop, "UPtr", IVirtualDesktopManagerInternal, "UPtrP", CurrentIVirtualDesktop, "UInt")
    return CurrentIVirtualDesktop
}

VD_getCurrentDesktop() ;this will return whichDesktop
{
    global
    CurrentIVirtualDesktop := 0
    DllCall(GetCurrentDesktop, "UPtr", IVirtualDesktopManagerInternal, "UPtrP", CurrentIVirtualDesktop, "UInt")

    VarSetCapacity(vd_strGUID, (38 + 1) * 2)
    VarSetCapacity(vd_GUID, 16)

    DllCall(VD_vtable(CurrentIVirtualDesktop,4), "UPtr", CurrentIVirtualDesktop, "UPtr", &vd_GUID, "UInt")

    DllCall("Ole32.dll\StringFromGUID2", "UPtr", &vd_GUID, "UPtr", &vd_strGUID, "Int", 38 + 1)
    currentDesktop_strGUID:=StrGet(&vd_strGUID, "UTF-16")

    ; IVirtualDesktopManagerInternal::GetDesktops method
    IObjectArray := 0
    DllCall(GetDesktops, "UPtr", IVirtualDesktopManagerInternal, "UPtrP", IObjectArray, "UInt")
    ; IObjectArray::GetCount method
    ; https://docs.microsoft.com/en-us/windows/desktop/api/objectarray/nf-objectarray-iobjectarray-getcount
    vd_Count := 0
    DllCall(VD_vtable(IObjectArray,3), "UPtr", IObjectArray, "UIntP", vd_Count, "UInt")

    IVirtualDesktop := 0
    Loop % (vd_Count)
    {
        ; https://github.com/nullpo-head/Windows-10-Virtual-Desktop-Switching-Shortcut/blob/master/VirtualDesktopSwitcher/VirtualDesktopSwitcher/VirtualDesktops.h
        DllCall("Ole32.dll\CLSIDFromString", "Str", "{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}", "UPtr", &vd_GUID)

        ; IObjectArray::GetAt method
        ; https://docs.microsoft.com/en-us/windows/desktop/api/objectarray/nf-objectarray-iobjectarray-getat
        DllCall(VD_vtable(IObjectArray,4), "UPtr", IObjectArray, "UInt", A_Index-1, "UPtr", &vd_GUID, "UPtrP", IVirtualDesktop, "UInt")

        ; IVirtualDesktop::GetID method
        DllCall(VD_vtable(IVirtualDesktop,4), "UPtr", IVirtualDesktop, "UPtr", &vd_GUID, "UInt")
        DllCall("Ole32.dll\StringFromGUID2", "UPtr", &vd_GUID, "UPtr", &vd_strGUID, "Int", 38 + 1)
        if (StrGet(&vd_strGUID, "UTF-16") = currentDesktop_strGUID) {
            return A_Index
        }
    }
}
VD_getDesktopOfWindow(wintitle)
{
    global
    DetectHiddenWindows, on
    WinGet, hwndsOfWinTitle, List, %wintitle%
    DetectHiddenWindows, off
    loop % hwndsOfWinTitle {
        IfEqual, False, % VD_isValidWindow(hwndsOfWinTitle%A_Index%), continue

        VarSetCapacity(vd_GUID, 16)
        vd_HRESULT := DllCall(GetWindowDesktopId, "UPtr", IVirtualDesktopManager, "Ptr", hwndsOfWinTitle%A_Index%, "UPtr", &vd_GUID, "UInt")
        if ( !vd_HRESULT ) ; OK
        {
            VarSetCapacity(vd_strGUID, (38 + 1) * 2)
            DllCall("Ole32.dll\StringFromGUID2", "UPtr", &vd_GUID, "UPtr", &vd_strGUID, "Int", 38 + 1)
            desktopOfWindow:=StrGet(&vd_strGUID, "UTF-16")
            if (desktopOfWindow and desktopOfWindow!="{00000000-0000-0000-0000-000000000000}") {
                break
            }
        }
    }
    ; IVirtualDesktopManagerInternal::GetDesktops method
    IObjectArray := 0
    DllCall(GetDesktops, "UPtr", IVirtualDesktopManagerInternal, "UPtrP", IObjectArray, "UInt")
    ; IObjectArray::GetCount method
    ; https://docs.microsoft.com/en-us/windows/desktop/api/objectarray/nf-objectarray-iobjectarray-getcount
    vd_Count := 0
    DllCall(VD_vtable(IObjectArray,3), "UPtr", IObjectArray, "UIntP", vd_Count, "UInt")

    VarSetCapacity(vd_strGUID, (38 + 1) * 2)
    VarSetCapacity(vd_GUID, 16)

    IVirtualDesktop := 0
    Loop % (vd_Count)
    {
        ; https://github.com/nullpo-head/Windows-10-Virtual-Desktop-Switching-Shortcut/blob/master/VirtualDesktopSwitcher/VirtualDesktopSwitcher/VirtualDesktops.h
        DllCall("Ole32.dll\CLSIDFromString", "Str", "{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}", "UPtr", &vd_GUID)

        ; IObjectArray::GetAt method
        ; https://docs.microsoft.com/en-us/windows/desktop/api/objectarray/nf-objectarray-iobjectarray-getat
        DllCall(VD_vtable(IObjectArray,4), "UPtr", IObjectArray, "UInt", A_Index-1, "UPtr", &vd_GUID, "UPtrP", IVirtualDesktop, "UInt")

        ; IVirtualDesktop::GetID method
        DllCall(VD_vtable(IVirtualDesktop,4), "UPtr", IVirtualDesktop, "UPtr", &vd_GUID, "UInt")
        DllCall("Ole32.dll\StringFromGUID2", "UPtr", &vd_GUID, "UPtr", &vd_strGUID, "Int", 38 + 1)
        if (StrGet(&vd_strGUID, "UTF-16") = desktopOfWindow) {
            return A_Index
        }
    }
}

VD_goToDesktopOfWindow(wintitle, activate:=true)
{
    global
    DetectHiddenWindows, on
    WinGet, hwndsOfWinTitle, List, %wintitle%
    DetectHiddenWindows, off
    loop % hwndsOfWinTitle {
        IfEqual, False, % VD_isValidWindow(hwndsOfWinTitle%A_Index%), continue

        VarSetCapacity(vd_GUID, 16)
        vd_HRESULT := DllCall(GetWindowDesktopId, "UPtr", IVirtualDesktopManager, "Ptr", hwndsOfWinTitle%A_Index%, "UPtr", &vd_GUID, "UInt")
        if ( !vd_HRESULT ) ; OK
        {
            VarSetCapacity(vd_strGUID, (38 + 1) * 2)
            DllCall("Ole32.dll\StringFromGUID2", "UPtr", &vd_GUID, "UPtr", &vd_strGUID, "Int", 38 + 1)
            desktopOfWindow:=StrGet(&vd_strGUID, "UTF-16")
            if (desktopOfWindow and desktopOfWindow!="{00000000-0000-0000-0000-000000000000}") {
                theHwnd:=hwndsOfWinTitle%A_Index%
                break
            }
        }
    }

    ; IVirtualDesktopManagerInternal::GetDesktops method
    IObjectArray := 0
    DllCall(GetDesktops, "UPtr", IVirtualDesktopManagerInternal, "UPtrP", IObjectArray, "UInt")
    ; IObjectArray::GetCount method
    ; https://docs.microsoft.com/en-us/windows/desktop/api/objectarray/nf-objectarray-iobjectarray-getcount
    vd_Count := 0
    DllCall(VD_vtable(IObjectArray,3), "UPtr", IObjectArray, "UIntP", vd_Count, "UInt")

    VarSetCapacity(vd_strGUID, (38 + 1) * 2)
    VarSetCapacity(vd_GUID, 16)

    IVirtualDesktop := 0
    Loop % (vd_Count)
    {
        ; https://github.com/nullpo-head/Windows-10-Virtual-Desktop-Switching-Shortcut/blob/master/VirtualDesktopSwitcher/VirtualDesktopSwitcher/VirtualDesktops.h
        DllCall("Ole32.dll\CLSIDFromString", "Str", "{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}", "UPtr", &vd_GUID)

        ; IObjectArray::GetAt method
        ; https://docs.microsoft.com/en-us/windows/desktop/api/objectarray/nf-objectarray-iobjectarray-getat
        DllCall(VD_vtable(IObjectArray,4), "UPtr", IObjectArray, "UInt", A_Index-1, "UPtr", &vd_GUID, "UPtrP", IVirtualDesktop, "UInt")

        ; IVirtualDesktop::GetID method
        DllCall(VD_vtable(IVirtualDesktop,4), "UPtr", IVirtualDesktop, "UPtr", &vd_GUID, "UInt")
        DllCall("Ole32.dll\StringFromGUID2", "UPtr", &vd_GUID, "UPtr", &vd_strGUID, "Int", 38 + 1)
        if (StrGet(&vd_strGUID, "UTF-16") = desktopOfWindow) {
            VD_SwitchDesktop(IVirtualDesktop)
            if (activate)
                WinActivate, ahk_id %theHwnd%
        }
    }
}

VD_sendToDesktop(wintitle,whichDesktop,followYourWindow:=true,activate:=true)
{
    global

    if (VD_ByrefpViewAndHwnd(wintitle, thePView, theHwnd)) { ;Byref
        return
    }

    if (thePView) {
        IObjectArray := 0
        DllCall(GetDesktops, "UPtr", IVirtualDesktopManagerInternal, "UPtrP", IObjectArray, "UInt")

        VarSetCapacity(vd_strGUID, (38 + 1) * 2)
        VarSetCapacity(vd_GUID, 16)

        IVirtualDesktop := 0

        ; https://github.com/nullpo-head/Windows-10-Virtual-Desktop-Switching-Shortcut/blob/master/VirtualDesktopSwitcher/VirtualDesktopSwitcher/VirtualDesktops.h
        DllCall("Ole32.dll\CLSIDFromString", "Str", "{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}", "UPtr", &vd_GUID)

        ; IObjectArray::GetAt method
        ; https://docs.microsoft.com/en-us/windows/desktop/api/objectarray/nf-objectarray-iobjectarray-getat
        DllCall(VD_vtable(IObjectArray,4), "UPtr", IObjectArray, "UInt", whichDesktop - 1, "UPtr", &vd_GUID, "UPtrP", IVirtualDesktop, "UInt")

        DllCall(MoveViewToDesktop, "ptr", IVirtualDesktopManagerInternal, "Ptr", thePView, "UPtr", IVirtualDesktop, "UInt")

        if (followYourWindow) {
            VD_SwitchDesktop(IVirtualDesktop)
            WinActivate, ahk_id %theHwnd%
        }
    }
}

VD_sendToCurrentDesktop(wintitle,activate:=true)
{
    global

    if (VD_ByrefpViewAndHwnd(wintitle, thePView, theHwnd)) { ;Byref
        return
    }

    CurrentIVirtualDesktop := 0
    DllCall(GetCurrentDesktop, "UPtr", IVirtualDesktopManagerInternal, "UPtrP", CurrentIVirtualDesktop, "UInt")

    DllCall(MoveViewToDesktop, "ptr", IVirtualDesktopManagerInternal, "Ptr", thePView, "UPtr", CurrentIVirtualDesktop, "UInt")
    if (activate)
        WinActivate, ahk_id %theHwnd%

}

; VD_toggleShowOnAllDesktops(wintitle) {
; VD_toggleShowOnAllDesktops(wintitle) {
; WinGet, WinExStyle, ExStyle, % wintitle
; If (WinExStyle & 0x00000080)
; WinSet, ExStyle, -0x00000080, % wintitle
; else
; WinSet, ExStyle, +0x00000080, % wintitle
; }

; VD_showOnAllDesktops(wintitle, enableOrDisable:=true) {
; WinSet, ExStyle, % (enableOrDisable ? "+" : "-") "0x00000080", % wintitle
; }

VD_createDesktop(goThere:=true) {
    global CreateDesktop, IVirtualDesktopManagerInternal

    DllCall(CreateDesktop, "UPtr", IVirtualDesktopManagerInternal, "Ptr*", newlyCreatedDesktop)

    if (goThere) {
        VD_goToDesktop(VD_getCount())
    }
}

VD_createUntil(howMany) {

    loop % howMany - VD_getCount() {
        VD_createDesktop(false)
    }

    if (goThere) {
        VD_goToDesktop(VD_getCount())
    }
}

VD_removeDesktop(whichDesktop, fallback_which:=false) {
    ;FALLBACK IS ONLY USED IF YOU ARE CURRENTLY ON THAT VD
    VD_internal_removeDesktop(VD_IVirtualDesktop_from_whichDesktop(whichDesktop), VD_IVirtualDesktop_from_whichDesktop(fallback_which))
}

VD_internal_removeDesktop(IVirtualDesktop, fallback:=false) {
    global RemoveDesktop, GetAdjacentDesktop, IVirtualDesktopManagerInternal

    if (fallback==false) {
        ;look left
        DllCall(GetAdjacentDesktop, "UPtr", IVirtualDesktopManagerInternal, "Ptr", IVirtualDesktop, "Uint", 3, "Ptr*", fallback) ; 3 = LeftDirection
        if (fallback==0) {
            ;look right
            DllCall(GetAdjacentDesktop, "UPtr", IVirtualDesktopManagerInternal, "Ptr", IVirtualDesktop, "Uint", 4, "Ptr*", fallback) ; 4 = RightDirection
            if (fallback==0) {
                return false
            }
        }
    }

    ;FALLBACK IS ONLY USED IF YOU ARE CURRENTLY ON THAT VD
    DllCall(RemoveDesktop, "UPtr", IVirtualDesktopManagerInternal, "Ptr", IVirtualDesktop, "Ptr", fallback)

}

VD_IsWindowPinned(wintitle) {
    global IsViewPinned, IVirtualDesktopPinnedApps

    if (VD_ByrefpViewAndHwnd(wintitle, thePView, theHwnd)) { ;Byref
        return
    }
    ; https://github.com/Ciantic/VirtualDesktopAccessor/blob/5bc1bbaab247b5d72e70abc9432a15275fd2d229/VirtualDesktopAccessor/dllmain.h#L377
    ; pinnedApps->IsViewPinned(pView, &isPinned);
    viewIsPinned:=0
    DllCall(IsViewPinned, "UPtr", IVirtualDesktopPinnedApps, "Ptr", thePView, "Int*",viewIsPinned)
    return viewIsPinned
}
VD_TogglePinWindow(wintitle) {
    global IsViewPinned, PinView, UnPinView, IVirtualDesktopPinnedApps

    if (VD_ByrefpViewAndHwnd(wintitle, thePView, theHwnd)) { ;Byref
        return
    }

    viewIsPinned:=0
    DllCall(IsViewPinned, "UPtr", IVirtualDesktopPinnedApps, "Ptr", thePView, "Int*",viewIsPinned)
    if (viewIsPinned) {
        DllCall(UnPinView, "UPtr", IVirtualDesktopPinnedApps, "Ptr", thePView)
    } else {
        DllCall(PinView, "UPtr", IVirtualDesktopPinnedApps, "Ptr", thePView)
    }

}
VD_PinWindow(wintitle) {
    global PinView, IVirtualDesktopPinnedApps

    if (VD_ByrefpViewAndHwnd(wintitle, thePView, theHwnd)) { ;Byref
        return
    }

    DllCall(PinView, "UPtr", IVirtualDesktopPinnedApps, "Ptr", thePView)
}
VD_UnPinWindow(wintitle) {
    global UnPinView, IVirtualDesktopPinnedApps

    if (VD_ByrefpViewAndHwnd(wintitle, thePView, theHwnd)) { ;Byref
        return
    }

    DllCall(UnPinView, "UPtr", IVirtualDesktopPinnedApps, "Ptr", thePView)
}

VD_IsAppPinned(wintitle) {
    global IsAppIdPinned, IVirtualDesktopPinnedApps

    if (VD_ByrefpViewAndHwnd(wintitle, thePView, theHwnd)) { ;Byref
        return
    }

    appId:=VD_AppIdFromView(thePView)

    appIsPinned:=0
    DllCall(IsAppIdPinned, "UPtr", IVirtualDesktopPinnedApps, "Ptr", appId, "Int*",appIsPinned)
    return appIsPinned
}
VD_TogglePinApp(wintitle) {
    global IsAppIdPinned, PinAppID, UnpinAppID, IVirtualDesktopPinnedApps

    if (VD_ByrefpViewAndHwnd(wintitle, thePView, theHwnd)) { ;Byref
        return
    }

    appId:=VD_AppIdFromView(thePView)

    DllCall(IsAppIdPinned, "UPtr", IVirtualDesktopPinnedApps, "Ptr", appId, "Int*",appIsPinned)
    if (appIsPinned) {
        DllCall(UnpinAppID, "UPtr", IVirtualDesktopPinnedApps, "Ptr", appId)
    } else {
        DllCall(PinAppID, "UPtr", IVirtualDesktopPinnedApps, "Ptr", appId)
    }

}
VD_PinApp(wintitle) {
    global PinAppID, IVirtualDesktopPinnedApps

    if (VD_ByrefpViewAndHwnd(wintitle, thePView, theHwnd)) { ;Byref
        return
    }

    appId:=VD_AppIdFromView(thePView)

    DllCall(PinAppID, "UPtr", IVirtualDesktopPinnedApps, "Ptr", appId)
}
VD_UnPinApp(wintitle) {
    global UnpinAppID, IVirtualDesktopPinnedApps

    if (VD_ByrefpViewAndHwnd(wintitle, thePView, theHwnd)) { ;Byref
        return
    }

    appId:=VD_AppIdFromView(thePView)

    DllCall(UnpinAppID, "UPtr", IVirtualDesktopPinnedApps, "Ptr", appId)
}

;start of internal functions

VD_isValidWindow(hWnd)
{
    DetectHiddenWindows, on
    return (VD_getWintitle(hWnd) and VD_IsWindow(hWnd))
    DetectHiddenWindows, off
}
VD_getWintitle(hWnd) {
    WinGetTitle, title, ahk_id %hWnd%
    return title
}
VD_IsWindow(hWnd){
    ; DetectHiddenWindows, on
    WinGet, dwStyle, Style, ahk_id %hWnd%
    if ((dwStyle&0x08000000) || !(dwStyle&0x10000000)) {
        return false
    }
    WinGet, dwExStyle, ExStyle, ahk_id %hWnd%
    if (dwExStyle & 0x00000080) {
        return false
    }
    WinGetClass, szClass, ahk_id %hWnd%
    if (szClass = "TApplication") {
        return false
    }
    ; DetectHiddenWindows, off
    return true
}

VD_vtable(ppv, idx)
{
    Return NumGet(NumGet(0+ppv)+A_PtrSize*idx)
}

VD_AppIdFromView(view) {
    ; revelation, view IS the object, I was looking everywhere for CLSID of IApplicationView

    GetAppUserModelId:=VD_vtable(view, 17)

    ; p(GetAppUserModelId)

    ; appId:=""
    ; VarSetCapacity(appId, 3000)
    ; DllCall(GetAppUserModelId, "UPtr", view, "Ptr", &appId)
    DllCall(GetAppUserModelId, "UPtr", view, "Ptr*", appId)

    ; Ptr* passes the address to Receive the string
    ; &RECT passes the address to, well input the RECT, not for output
    ; VarSetCapacity(Rect, 16)  ; A RECT is a struct consisting of four 32-bit integers (i.e. 4*4=16).
    ; DllCall("GetWindowRect", "Ptr", WinExist(), "Ptr", &Rect)  ; WinExist() returns an HWND.
    ; MsgBox % "Left " . NumGet(Rect, 0, "Int") . " Top " . NumGet(Rect, 4, "Int")
    ; . " Right " . NumGet(Rect, 8, "Int") . " Bottom " . NumGet(Rect, 12, "Int")
    ; wait what?, now I'm confused
    ; it is because the function says [out]. not in. [in] write to YOUR struct
    ; , [out] CREATES the struct and returns the pointer to you

    ; DllCall(GetAppUserModelId, "UPtr", view, "Str", appId)
    ; p(appId)
    ; DllCall(GetAppUserModelId, "UPtr", view, "Ptr", &appId)
    ; p(StrGet(appId, 3000, "UTF-16"))

    ; https://stackoverflow.com/questions/27977474/how-to-get-appusermodelid-for-any-app-in-windows-7-8-using-vc#27977668
    ; https://docs.microsoft.com/en-us/windows/win32/api/shobjidl_core/nf-shobjidl_core-getcurrentprocessexplicitappusermodelid?redirectedfrom=MSDN
    ;   [out] PWSTR *AppID
    ; A pointer that receives the address of the AppUserModelID assigned to the process. The caller is responsible for freeing this string with CoTaskMemFree when it is no longer needed. (lol)

    ; p(appId)
    ; p(StrGet(appId, "UTF-16"))

    ; File := FileOpen(A_LineFile "\..\notes\dump", "w")
    ; File.RawWrite(0+appId, 3000)
    ; File.Close()

    return appId
}

VD_ViewFromHwnd(theHwnd) {
    global GetViewForHwnd, IApplicationViewCollection
    pView := 0
    DllCall(GetViewForHwnd, "UPtr", IApplicationViewCollection, "Ptr", theHwnd, "Ptr*", pView, "UInt")
    return pView
}

VD_ByrefpViewAndHwnd(wintitle, Byref pView,Byref theHwnd) {
    ;false if found, true if notFound
    global CanViewMoveDesktops, IVirtualDesktopManagerInternal

    DetectHiddenWindows, on
    WinGet, outHwndList, List, % wintitle
    DetectHiddenWindows, off
    loop % outHwndList {
        if (!VD_isValidWindow(outHwndList%A_Index%)) {
            continue
        }

        pView:=VD_ViewFromHwnd(outHwndList%A_Index%)

        pfCanViewMoveDesktops := 0
        DllCall(CanViewMoveDesktops, "ptr", IVirtualDesktopManagerInternal, "Ptr", pView, "int*", pfCanViewMoveDesktops, "UInt") ; return value BOOL
        if (pfCanViewMoveDesktops)
        {
            theHwnd:=outHwndList%A_Index%
            thePView:=pView
            return false
        }
    }
    return True
}
