; must call VD_init() before any of these functions
; VD_getCurrentVirtualDesktop() ;this will return whichDesktop
; VD_getDesktopOfWindow() ;this will return whichDesktop ;please use VD_goToDesktopOfWindow instead if you just want to go there.
; VD_getCount() ;this will return the number of virtual desktops you currently have
; VD_goToDesktop(whichDesktop)
; VD_goToDesktopOfWindow(wintitle, activate:=true)
; VD_sendToDesktop(wintitle,whichDesktop,followYourWindow:=false,activate:=true)
; VD_sendToCurrentDesktop(wintitle,activate:=true)

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

VD_getCurrentVirtualDesktop() ;this will return whichDesktop
{
    global
    CurrentIVirtualDesktop := 0
    GetCurrentDesktop_return_value := DllCall(GetCurrentDesktop, "UPtr", IVirtualDesktopManagerInternal, "UPtrP", CurrentIVirtualDesktop, "UInt")

    VarSetCapacity(strGUID, (38 + 1) * 2)
    VarSetCapacity(GUID, 16)

    DllCall(VD_vtable(CurrentIVirtualDesktop,4), "UPtr", CurrentIVirtualDesktop, "UPtr", &GUID, "UInt")
    firstPointer:=&GUID
    DllCall("Ole32.dll\StringFromGUID2", "UPtr", &GUID, "UPtr", &strGUID, "Int", 38 + 1)
    strGUID_currentDesktop:=StrGet(&strGUID, "UTF-16")

    ; IVirtualDesktopManagerInternal::GetDesktops method
    IObjectArray := 0
    DllCall(GetDesktops, "UPtr", IVirtualDesktopManagerInternal, "UPtrP", IObjectArray, "UInt")
    ; IObjectArray::GetCount method
    ; https://docs.microsoft.com/en-us/windows/desktop/api/objectarray/nf-objectarray-iobjectarray-getcount
    Count := 0
    DllCall(VD_vtable(IObjectArray,3), "UPtr", IObjectArray, "UIntP", Count, "UInt")

    IVirtualDesktop := 0
    Loop % (Count)
    {
        ; https://github.com/nullpo-head/Windows-10-Virtual-Desktop-Switching-Shortcut/blob/master/VirtualDesktopSwitcher/VirtualDesktopSwitcher/VirtualDesktops.h
        DllCall("Ole32.dll\CLSIDFromString", "Str", "{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}", "UPtr", &GUID)

        ; IObjectArray::GetAt method
        ; https://docs.microsoft.com/en-us/windows/desktop/api/objectarray/nf-objectarray-iobjectarray-getat
        DllCall(VD_vtable(IObjectArray,4), "UPtr", IObjectArray, "UInt", A_Index-1, "UPtr", &GUID, "UPtrP", IVirtualDesktop, "UInt")

        ; IVirtualDesktop::GetID method
        DllCall(VD_vtable(IVirtualDesktop,4), "UPtr", IVirtualDesktop, "UPtr", &GUID, "UInt")
        DllCall("Ole32.dll\StringFromGUID2", "UPtr", &GUID, "UPtr", &strGUID, "Int", 38 + 1)
        if (StrGet(&strGUID, "UTF-16") = strGUID_currentDesktop) {
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
        VarSetCapacity(GUID, 16)
        R := DllCall(GetWindowDesktopId, "UPtr", IVirtualDesktopManager, "Ptr", hwndsOfWinTitle%A_Index%, "UPtr", &GUID, "UInt")
        if ( !R ) ; OK
        {
            VarSetCapacity(strGUID, (38 + 1) * 2)
            DllCall("Ole32.dll\StringFromGUID2", "UPtr", &GUID, "UPtr", &strGUID, "Int", 38 + 1)
            desktopOfWindow:=StrGet(&strGUID, "UTF-16")
            if (desktopOfWindow and desktopOfWindow!="{00000000-0000-0000-0000-000000000000}") {
                hwndOfWindow:=hwndsOfWinTitle%A_Index%
                break
            }
        }
    }

    ; IVirtualDesktopManagerInternal::GetDesktops method
    IObjectArray := 0
    DllCall(GetDesktops, "UPtr", IVirtualDesktopManagerInternal, "UPtrP", IObjectArray, "UInt")
    ; IObjectArray::GetCount method
    ; https://docs.microsoft.com/en-us/windows/desktop/api/objectarray/nf-objectarray-iobjectarray-getcount
    Count := 0
    DllCall(VD_vtable(IObjectArray,3), "UPtr", IObjectArray, "UIntP", Count, "UInt")

    VarSetCapacity(strGUID, (38 + 1) * 2)
    VarSetCapacity(GUID, 16)

    IVirtualDesktop := 0
    Loop % (Count)
    {
        ; https://github.com/nullpo-head/Windows-10-Virtual-Desktop-Switching-Shortcut/blob/master/VirtualDesktopSwitcher/VirtualDesktopSwitcher/VirtualDesktops.h
        DllCall("Ole32.dll\CLSIDFromString", "Str", "{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}", "UPtr", &GUID)

        ; IObjectArray::GetAt method
        ; https://docs.microsoft.com/en-us/windows/desktop/api/objectarray/nf-objectarray-iobjectarray-getat
        DllCall(VD_vtable(IObjectArray,4), "UPtr", IObjectArray, "UInt", A_Index-1, "UPtr", &GUID, "UPtrP", IVirtualDesktop, "UInt")

        ; IVirtualDesktop::GetID method
        DllCall(VD_vtable(IVirtualDesktop,4), "UPtr", IVirtualDesktop, "UPtr", &GUID, "UInt")
        DllCall("Ole32.dll\StringFromGUID2", "UPtr", &GUID, "UPtr", &strGUID, "Int", 38 + 1)
        if (StrGet(&strGUID, "UTF-16") = desktopOfWindow) {
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
    Count := 0
    DllCall(VD_vtable(IObjectArray,3), "UPtr", IObjectArray, "UIntP", Count, "UInt")
    return Count
}
VD_goToDesktop(whichDesktop)
{
    global
    IObjectArray := 0
    DllCall(GetDesktops, "UPtr", IVirtualDesktopManagerInternal, "UPtrP", IObjectArray, "UInt")

    VarSetCapacity(strGUID, (38 + 1) * 2)
    VarSetCapacity(GUID, 16)

    IVirtualDesktop := 0

    ; https://github.com/nullpo-head/Windows-10-Virtual-Desktop-Switching-Shortcut/blob/master/VirtualDesktopSwitcher/VirtualDesktopSwitcher/VirtualDesktops.h
    DllCall("Ole32.dll\CLSIDFromString", "Str", "{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}", "UPtr", &GUID)

    ; IObjectArray::GetAt method
    ; https://docs.microsoft.com/en-us/windows/desktop/api/objectarray/nf-objectarray-iobjectarray-getat
    DllCall(VD_vtable(IObjectArray,4), "UPtr", IObjectArray, "UInt", whichDesktop -1, "UPtr", &GUID, "UPtrP", IVirtualDesktop, "UInt")

    DllCall(SwitchDesktop, "ptr", IVirtualDesktopManagerInternal, "UPtr", IVirtualDesktop, "UInt")
}

VD_goToDesktopOfWindow(wintitle, activate:=true)
{
    global
    DetectHiddenWindows, on
    WinGet, hwndsOfWinTitle, List, %wintitle%
    DetectHiddenWindows, off
    loop % hwndsOfWinTitle {
        VarSetCapacity(GUID, 16)
        R := DllCall(GetWindowDesktopId, "UPtr", IVirtualDesktopManager, "Ptr", hwndsOfWinTitle%A_Index%, "UPtr", &GUID, "UInt")
        if ( !R ) ; OK
        {
            VarSetCapacity(strGUID, (38 + 1) * 2)
            DllCall("Ole32.dll\StringFromGUID2", "UPtr", &GUID, "UPtr", &strGUID, "Int", 38 + 1)
            desktopOfWindow:=StrGet(&strGUID, "UTF-16")
            if (desktopOfWindow and desktopOfWindow!="{00000000-0000-0000-0000-000000000000}") {
                hwndOfWindow:=hwndsOfWinTitle%A_Index%
                break
            }
        }
    }

    ; IVirtualDesktopManagerInternal::GetDesktops method
    IObjectArray := 0
    DllCall(GetDesktops, "UPtr", IVirtualDesktopManagerInternal, "UPtrP", IObjectArray, "UInt")
    ; IObjectArray::GetCount method
    ; https://docs.microsoft.com/en-us/windows/desktop/api/objectarray/nf-objectarray-iobjectarray-getcount
    Count := 0
    DllCall(VD_vtable(IObjectArray,3), "UPtr", IObjectArray, "UIntP", Count, "UInt")

    VarSetCapacity(strGUID, (38 + 1) * 2)
    VarSetCapacity(GUID, 16)

    IVirtualDesktop := 0
    Loop % (Count)
    {
        ; https://github.com/nullpo-head/Windows-10-Virtual-Desktop-Switching-Shortcut/blob/master/VirtualDesktopSwitcher/VirtualDesktopSwitcher/VirtualDesktops.h
        DllCall("Ole32.dll\CLSIDFromString", "Str", "{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}", "UPtr", &GUID)

        ; IObjectArray::GetAt method
        ; https://docs.microsoft.com/en-us/windows/desktop/api/objectarray/nf-objectarray-iobjectarray-getat
        DllCall(VD_vtable(IObjectArray,4), "UPtr", IObjectArray, "UInt", A_Index-1, "UPtr", &GUID, "UPtrP", IVirtualDesktop, "UInt")

        ; IVirtualDesktop::GetID method
        DllCall(VD_vtable(IVirtualDesktop,4), "UPtr", IVirtualDesktop, "UPtr", &GUID, "UInt")
        DllCall("Ole32.dll\StringFromGUID2", "UPtr", &GUID, "UPtr", &strGUID, "Int", 38 + 1)
        if (StrGet(&strGUID, "UTF-16") = desktopOfWindow) {
            DllCall(SwitchDesktop, "ptr", IVirtualDesktopManagerInternal, "UPtr", IVirtualDesktop, "UInt")
            if (activate)
            WinActivate, ahk_id %hwndOfWindow%
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
        ; IObjectArray::GetCount method
        ; https://docs.microsoft.com/en-us/windows/desktop/api/objectarray/nf-objectarray-iobjectarray-getcount
        Count := 0
        DllCall(VD_vtable(IObjectArray,3), "UPtr", IObjectArray, "UIntP", Count, "UInt")

        VarSetCapacity(strGUID, (38 + 1) * 2)
        VarSetCapacity(GUID, 16)

        IVirtualDesktop := 0

        ; https://github.com/nullpo-head/Windows-10-Virtual-Desktop-Switching-Shortcut/blob/master/VirtualDesktopSwitcher/VirtualDesktopSwitcher/VirtualDesktops.h
        DllCall("Ole32.dll\CLSIDFromString", "Str", "{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}", "UPtr", &GUID)

        ; IObjectArray::GetAt method
        ; https://docs.microsoft.com/en-us/windows/desktop/api/objectarray/nf-objectarray-iobjectarray-getat
        DllCall(VD_vtable(IObjectArray,4), "UPtr", IObjectArray, "UInt", whichDesktop -1, "UPtr", &GUID, "UPtrP", IVirtualDesktop, "UInt")

        DllCall(MoveViewToDesktop, "ptr", IVirtualDesktopManagerInternal, "Ptr", thePView, "UPtr", IVirtualDesktop, "UInt")

        if (followYourWindow) {
            DllCall(SwitchDesktop, "ptr", IVirtualDesktopManagerInternal, "UPtr", IVirtualDesktop, "UInt")
            WinActivate, ahk_id %hwndOfWindow%
        }
    }
}

VD_sendToCurrentDesktop(wintitle,activate:=true)
{
    global
    DetectHiddenWindows, on
    WinGet, outHwndList, List, %wintitle%
    DetectHiddenWindows, off

    thePView:=0
    loop % outHwndList {
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
        CurrentIVirtualDesktop := 0
        GetCurrentDesktop_return_value := DllCall(GetCurrentDesktop, "UPtr", IVirtualDesktopManagerInternal, "UPtrP", CurrentIVirtualDesktop, "UInt")

        DllCall(MoveViewToDesktop, "ptr", IVirtualDesktopManagerInternal, "Ptr", thePView, "UPtr", CurrentIVirtualDesktop, "UInt")
        if (activate)
        winactivate, ahk_id %theHwnd%
    }
}


VD_vtable(ppv, idx)
{
    Return NumGet(NumGet(1*ppv)+A_PtrSize*idx)
}