; must call VD_init() before any of these functions
; VD_getCurrentDesktop() ;this will return whichDesktop
; VD_getDesktopOfWindow(wintitle) ;this will return whichDesktop ;please use VD_goToDesktopOfWindow instead if you just want to go there.
; VD_getCount() ;this will return the number of virtual desktops you currently have
; VD_goToDesktop(whichDesktop)
; VD_goToDesktopOfWindow(wintitle, activate:=true)
; VD_sendToDesktop(wintitle,whichDesktop,followYourWindow:=false,activate:=true)
; VD_sendToCurrentDesktop(wintitle,activate:=true)

; internal functions
; VD_getCurrentIVirtualDesktop()
; VD_SwitchDesktop(IVirtualDesktop)
; VD_isValidWindow(hWnd)
; VD_getWintitle(hWnd)
; VD_AltTab(){
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

    ImmersiveShell := ComObjCreate("{C2F03A33-21F5-47FA-B4BB-156362A2F239}", "{00000000-0000-0000-C000-000000000046}") 
    if !(IApplicationViewCollection := ComObjQuery(ImmersiveShell,"{1841C6D7-4F9D-42C0-AF41-8747538F10E5}","{1841C6D7-4F9D-42C0-AF41-8747538F10E5}" ) ) ; 1607-1809
    {
        MsgBox IApplicationViewCollection interface not supported.
    }
    GetViewForHwnd := VD_vtable(IApplicationViewCollection, 6) ; (IntPtr hwnd, out IApplicationView view);
}

VD_getCurrentIVirtualDesktop()
{
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
    thePView:=0

    DetectHiddenWindows, on
    WinGet, outHwndList, List, %wintitle%
    DetectHiddenWindows, off
    loop % outHwndList {
        IfEqual, False, % VD_isValidWindow(outHwndList%A_Index%), continue

        pView := 0
        DllCall(GetViewForHwnd, "UPtr", IApplicationViewCollection, "Ptr", outHwndList%A_Index%, "Ptr*", pView, "UInt")

        pfCanViewMoveDesktops := 0
        DllCall(CanViewMoveDesktops, "ptr", IVirtualDesktopManagerInternal, "Ptr", pView, "int*", pfCanViewMoveDesktops, "UInt") ; return value BOOL
        if (pfCanViewMoveDesktops)
        {
            theHwnd:=outHwndList%A_Index%
            thePView:=pView
            break
        }
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

    DetectHiddenWindows, on
    WinGet, outHwndList, List, %wintitle%
    DetectHiddenWindows, off

    pfCanViewMoveDesktops:=false
    thePView:=0
    loop % outHwndList {
        IfEqual, False, % VD_isValidWindow(outHwndList%A_Index%), continue

        pView := 0
        DllCall(GetViewForHwnd, "UPtr", IApplicationViewCollection, "Ptr", outHwndList%A_Index%, "Ptr*", pView, "UInt")

        pfCanViewMoveDesktops := 0
        DllCall(CanViewMoveDesktops, "ptr", IVirtualDesktopManagerInternal, "Ptr", pView, "int*", pfCanViewMoveDesktops, "UInt") ; return value BOOL
        if (pfCanViewMoveDesktops)
        {
            theHwnd:=outHwndList%A_Index%
            thePView:=pView
            break
        }
    }
    if (pfCanViewMoveDesktops) {
        CurrentIVirtualDesktop := 0
        DllCall(GetCurrentDesktop, "UPtr", IVirtualDesktopManagerInternal, "UPtrP", CurrentIVirtualDesktop, "UInt")

        DllCall(MoveViewToDesktop, "ptr", IVirtualDesktopManagerInternal, "Ptr", thePView, "UPtr", CurrentIVirtualDesktop, "UInt")
        if (activate)
            winactivate, ahk_id %theHwnd%
    }

}
;start of internal functions
VD_SwitchDesktop(IVirtualDesktop)
{
    global
    winactivate, ahk_class Shell_TrayWnd
    WinWaitActive, ahk_class Shell_TrayWnd
    DllCall(SwitchDesktop, "ptr", IVirtualDesktopManagerInternal, "UPtr", IVirtualDesktop, "UInt")
    DllCall(SwitchDesktop, "ptr", IVirtualDesktopManagerInternal, "UPtr", IVirtualDesktop, "UInt")
    WinMinimize, ahk_class Shell_TrayWnd
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
    Return NumGet(NumGet(1*ppv)+A_PtrSize*idx)
}