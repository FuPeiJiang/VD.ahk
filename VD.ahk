; must call VD_init() before any of these functions
; VD_getCurrentDesktop() ;this will return whichDesktop
; VD_getDesktopOfWindow(wintitle) ;this will return whichDesktop ;please use VD_goToDesktopOfWindow instead if you just want to go there.
; VD_getCount() ;this will return the number of virtual desktops you currently have
; VD_goToDesktop(whichDesktop)
; VD_goToDesktopOfWindow(wintitle, activate:=true)
; VD_sendToDesktop(wintitle,whichDesktop,followYourWindow:=false,activate:=true)
; VD_sendToCurrentDesktop(wintitle,activate:=true)

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

VD_init()
{
    global
    IVirtualDesktopManager := ComObjCreate("{AA509086-5CA9-4C25-8F95-589D3C07B48A}", "{A5CD92FF-29BE-454C-8D04-D82879FB3F1B}")
    GetWindowDesktopId := VD_vtable(IVirtualDesktopManager, 4)

    IServiceProvider := ComObjCreate("{C2F03A33-21F5-47FA-B4BB-156362A2F239}", "{6D5140C1-7436-11CE-8034-00AA006009FA}")

    IVirtualDesktopManagerInternal := ComObjQuery(IServiceProvider, "{C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B}", "{F31574D6-B682-4CDC-BD56-1827860ABEC6}")
    MoveViewToDesktop := VD_vtable(IVirtualDesktopManagerInternal, 4) ; void MoveViewToDesktop(object pView, IVirtualDesktop desktop);
    GetCurrentDesktop := VD_vtable(IVirtualDesktopManagerInternal, 6) ; IVirtualDesktop GetCurrentDesktop();
    CanViewMoveDesktops := VD_vtable(IVirtualDesktopManagerInternal, 5) ; bool CanViewMoveDesktops(object pView);
    GetDesktops := VD_vtable(IVirtualDesktopManagerInternal, 7) ; IObjectArray GetDesktops();
    SwitchDesktop := VD_vtable(IVirtualDesktopManagerInternal, 9) ; void SwitchDesktop(IVirtualDesktop desktop);

    ;// https://github.com/MScholtes/VirtualDesktop/blob/812c321e286b82a10f8050755c94d21c4b69812f/VirtualDesktop1803.cs#L180-L188
    IVirtualDesktopPinnedApps := ComObjQuery(IServiceProvider, "{B5A399E7-1C87-46B8-88E9-FC5747B171BD}", "{4CE81583-1E4C-4632-A621-07A53543148F}")
    IsAppIdPinned := VD_vtable(IVirtualDesktopPinnedApps, 3) ; bool IsAppIdPinned(string appId);
    PinAppID := VD_vtable(IVirtualDesktopPinnedApps, 4) ; void PinAppID(string appId);
    UnpinAppID := VD_vtable(IVirtualDesktopPinnedApps, 5) ; void UnpinAppID(string appId);
    IsViewPinned := VD_vtable(IVirtualDesktopPinnedApps, 6) ; bool IsViewPinned(IApplicationView applicationView);
    PinView := VD_vtable(IVirtualDesktopPinnedApps, 7) ; void PinView(IApplicationView applicationView);
    UnpinView := VD_vtable(IVirtualDesktopPinnedApps, 8) ; void UnpinView(IApplicationView applicationView);

    IApplicationView := ComObjQuery(IServiceProvider, "{B5A399E7-1C87-46B8-88E9-FC5747B171BD}", "{F31574D6-B682-4CDC-BD56-1827860ABEC6}")
    ; EXTERN_C const IID IID_IInspectable; : 6 methods
    ; https://github.com/MScholtes/VirtualDesktop/blob/812c321e286b82a10f8050755c94d21c4b69812f/VirtualDesktop1803.cs#L73
    ; internal interface IApplicationView
    ; {
    ;6  int SetFocus();
    ;7  int SwitchTo();
    ;8  int TryInvokeBack(IntPtr /* IAsyncCallback* */ callback);
    ;9  int GetThumbnailWindow(out IntPtr hwnd);
    ;10 int GetMonitor(out IntPtr /* IImmersiveMonitor */ immersiveMonitor);
    ;11 int GetVisibility(out int visibility);
    ;12 int SetCloak(APPLICATION_VIEW_CLOAK_TYPE cloakType, int unknown);
    ;13 int GetPosition(ref Guid guid /* GUID for IApplicationViewPosition */, out IntPtr /* IApplicationViewPosition** */ position);
    ;14 int SetPosition(ref IntPtr /* IApplicationViewPosition* */ position);
    ;15 int InsertAfterWindow(IntPtr hwnd);
    ;16 int GetExtendedFramePosition(out Rect rect);
    ;17 int GetAppUserModelId([MarshalAs(UnmanagedType.LPWStr)] out string id);
    ;18 int SetAppUserModelId(string id);
    ;19 int IsEqualByAppUserModelId(string id, out int result);
    ;20 int GetViewState(out uint state);
    ;21 int SetViewState(uint state);
    ;22 int GetNeediness(out int neediness);
    ;23 int GetLastActivationTimestamp(out ulong timestamp);
    ;24 int SetLastActivationTimestamp(ulong timestamp);
    ;25 int GetVirtualDesktopId(out Guid guid);
    ;26 int SetVirtualDesktopId(ref Guid guid);
    ;27 int GetShowInSwitchers(out int flag);
    ;28 int SetShowInSwitchers(int flag);
    ;29 int GetScaleFactor(out int factor);
    ;30 int CanReceiveInput(out bool canReceiveInput);
    ;31 int GetCompatibilityPolicyType(out APPLICATION_VIEW_COMPATIBILITY_POLICY flags);
    ;32 int SetCompatibilityPolicyType(APPLICATION_VIEW_COMPATIBILITY_POLICY flags);
    ;33 int GetSizeConstraints(IntPtr /* IImmersiveMonitor* */ monitor, out Size size1, out Size size2);
    ;34 int GetSizeConstraintsForDpi(uint uint1, out Size size1, out Size size2);
    ;35 int SetSizeConstraintsForDpi(ref uint uint1, ref Size size1, ref Size size2);
    ;36 int OnMinSizePreferencesUpdated(IntPtr hwnd);
    ;37 int ApplyOperation(IntPtr /* IApplicationViewOperation* */ operation);
    ;38 int IsTray(out bool isTray);
    ;39 int IsInHighZOrderBand(out bool isInHighZOrderBand);
    ;40 int IsSplashScreenPresented(out bool isSplashScreenPresented);
    ;41 int Flash();
    ;42 int GetRootSwitchableOwner(out IApplicationView rootSwitchableOwner);
    ;43 int EnumerateOwnershipTree(out IObjectArray ownershipTree);
    ;44 int GetEnterpriseId([MarshalAs(UnmanagedType.LPWStr)] out string enterpriseId);
    ;45 int IsMirrored(out bool isMirrored);
    ;46 int Unknown1(out int unknown);
    ;47 int Unknown2(out int unknown);
    ;48 int Unknown3(out int unknown);
    ;49 int Unknown4(out int unknown);
    ; }

    ImmersiveShell := ComObjCreate("{C2F03A33-21F5-47FA-B4BB-156362A2F239}", "{00000000-0000-0000-C000-000000000046}") 
    ; if !(IApplicationViewCollection := ComObjQuery(ImmersiveShell,"{1841C6D7-4F9D-42C0-AF41-8747538F10E5}" ) ) ; doesn't work
    ; SAME CLSID and IID ?
    ; wait it's not CLSID:
    ; SID
    ; A service identifier in the same form as IID. When omitting this parameter, also omit the comma.
    if !(IApplicationViewCollection := ComObjQuery(ImmersiveShell,"{1841C6D7-4F9D-42C0-AF41-8747538F10E5}","{1841C6D7-4F9D-42C0-AF41-8747538F10E5}" ) ) ; 1607-1809
    {
        MsgBox IApplicationViewCollection interface not supported.
    }
    GetViewForHwnd := VD_vtable(IApplicationViewCollection, 6) ; (IntPtr hwnd, out IApplicationView view);
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
VD_getCount()
{
    global
    IObjectArray := 0
    DllCall(GetDesktops, "UPtr", IVirtualDesktopManagerInternal, "UPtrP", IObjectArray, "UInt")

    ; IObjectArray::GetCount method
    ; https://docs.microsoft.com/en-us/windows/desktop/api/objectarray/nf-objectarray-iobjectarray-getcount
    vd_Count := 0
    DllCall(VD_vtable(IObjectArray,3), "UPtr", IObjectArray, "UIntP", vd_Count, "UInt")
    return vd_Count
}
VD_goToDesktop(whichDesktop)
{
    global
    IObjectArray := 0
    DllCall(GetDesktops, "UPtr", IVirtualDesktopManagerInternal, "UPtrP", IObjectArray, "UInt")

    VarSetCapacity(vd_strGUID, (38 + 1) * 2)
    VarSetCapacity(vd_GUID, 16)

    IVirtualDesktop := 0

    ; https://github.com/nullpo-head/Windows-10-Virtual-Desktop-Switching-Shortcut/blob/master/VirtualDesktopSwitcher/VirtualDesktopSwitcher/VirtualDesktops.h
    DllCall("Ole32.dll\CLSIDFromString", "Str", "{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}", "UPtr", &vd_GUID)

    ; IObjectArray::GetAt method
    ; https://docs.microsoft.com/en-us/windows/desktop/api/objectarray/nf-objectarray-iobjectarray-getat
    DllCall(VD_vtable(IObjectArray,4), "UPtr", IObjectArray, "UInt", whichDesktop -1, "UPtr", &vd_GUID, "UPtrP", IVirtualDesktop, "UInt")

    VD_SwitchDesktop(IVirtualDesktop)
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
VD_sendToDesktop(wintitle,whichDesktop,followYourWindow:=false,activate:=true)
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
        DllCall(VD_vtable(IObjectArray,4), "UPtr", IObjectArray, "UInt", whichDesktop -1, "UPtr", &vd_GUID, "UPtrP", IVirtualDesktop, "UInt")

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
VD_SwitchDesktop(IVirtualDesktop)
{
    global SwitchDesktop, IVirtualDesktopManagerInternal
    ;activate taskbar before
    WinActivate, ahk_class Shell_TrayWnd
    WinWaitActive, ahk_class Shell_TrayWnd
    DllCall(SwitchDesktop, "ptr", IVirtualDesktopManagerInternal, "UPtr", IVirtualDesktop, "UInt")
    DllCall(SwitchDesktop, "ptr", IVirtualDesktopManagerInternal, "UPtr", IVirtualDesktop, "UInt")
    WinMinimize, ahk_class Shell_TrayWnd
}
VD_isWindowFullScreen( winTitle ) {
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
