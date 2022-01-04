; VD.getCurrentDesktopNum()
; VD.getDesktopNumOfWindow(wintitle) ;please use VD.goToDesktopOfWindow instead if you just want to go there.

; VD.getCount() ;how many virtual desktops you now have

; VD.goToDesktopNum(desktopNum)
; VD.goToDesktopOfWindow(wintitle, activateYourWindow:=true)

; VD.sendWindowToDesktopNum(wintitle,desktopNum,followYourWindow:=true,activateYourWindow:=true)
; VD.sendWindowToCurrentDesktop(wintitle,activateYourWindow:=true)

; optional: call VD.init() to initialize internal variables before using methods, or else variables will be initialized when you first use the class(I think)

; VD.createDesktop(goThere:=true) ; VD.createUntil(howMany, goToLastlyCreated:=true)
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
            IID_IVirtualDesktopManagerInternal_:="{F31574D6-B682-4CDC-BD56-1827860ABEC6}" ;https://github.com/MScholtes/VirtualDesktop/blob/812c321e286b82a10f8050755c94d21c4b69812f/VirtualDesktop.cs#L178-L191
            this._dll_GetDesktops:=this._dll_GetDesktops_Win10 ;conditionally assign method to method
            this._dll_SwitchDesktop:=this._dll_SwitchDesktop_Win10 ;conditionally assign method to method
            IID_IVirtualDesktop_:="{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}" ;https://github.com/MScholtes/VirtualDesktop/blob/812c321e286b82a10f8050755c94d21c4b69812f/VirtualDesktop.cs#L149-L150
        }
        else
        {
            ; Windows 11
            IID_IVirtualDesktopManagerInternal_:="{B2F925B9-5A0F-4D2E-9F4D-2B1507593C10}" ;https://github.com/MScholtes/VirtualDesktop/blob/812c321e286b82a10f8050755c94d21c4b69812f/VirtualDesktop11.cs#L163-L185
            this._dll_GetDesktops:=this._dll_GetDesktops_Win11 ;conditionally assign method to method
            this._dll_SwitchDesktop:=this._dll_SwitchDesktop_Win11 ;conditionally assign method to method
            IID_IVirtualDesktop_:="{536D3495-B208-4CC9-AE26-DE8111275BF8}" ;https://github.com/MScholtes/VirtualDesktop/blob/812c321e286b82a10f8050755c94d21c4b69812f/VirtualDesktop11.cs#L149-L150
        }
        ;----------------------
        this.IVirtualDesktopManager := ComObjCreate("{AA509086-5CA9-4C25-8F95-589D3C07B48A}", "{A5CD92FF-29BE-454C-8D04-D82879FB3F1B}")
        this.GetWindowDesktopId := this._vtable(this.IVirtualDesktopManager, 4)

        IServiceProvider := ComObjCreate("{C2F03A33-21F5-47FA-B4BB-156362A2F239}", "{6D5140C1-7436-11CE-8034-00AA006009FA}")

        ;https://github.com/MScholtes/VirtualDesktop/blob/812c321e286b82a10f8050755c94d21c4b69812f/VirtualDesktop.cs#L178-L191
        this.IVirtualDesktopManagerInternal := ComObjQuery(IServiceProvider, "{C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B}", IID_IVirtualDesktopManagerInternal_)
        ; this.GetCount := this._vtable(this.IVirtualDesktopManagerInternal, 3 ; int GetCount();
        this.MoveViewToDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 4) ; void MoveViewToDesktop(object pView, IVirtualDesktop desktop);
        this.GetCurrentDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 6) ; IVirtualDesktop GetCurrentDesktop();
        this.CanViewMoveDesktops := this._vtable(this.IVirtualDesktopManagerInternal, 5) ; bool CanViewMoveDesktops(object pView);
        this.GetDesktops := this._vtable(this.IVirtualDesktopManagerInternal, 7) ; IObjectArray GetDesktops();
        this.GetAdjacentDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 8) ; int GetAdjacentDesktop(IVirtualDesktop from, int direction, out IVirtualDesktop desktop);
        this.SwitchDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 9) ; void SwitchDesktop(IVirtualDesktop desktop);
        this.Ptr_CreateDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 10) ; IVirtualDesktop CreateDesktop();
        this.Ptr_RemoveDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 11) ; void RemoveDesktop(IVirtualDesktop desktop, IVirtualDesktop fallback);
        this.FindDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 12) ; IVirtualDesktop FindDesktop(ref Guid desktopid);

        ;https://github.com/MScholtes/VirtualDesktop/blob/812c321e286b82a10f8050755c94d21c4b69812f/VirtualDesktop.cs#L225-L234
        this.IVirtualDesktopPinnedApps := ComObjQuery(IServiceProvider, "{B5A399E7-1C87-46B8-88E9-FC5747B171BD}", "{4CE81583-1E4C-4632-A621-07A53543148F}")
        this.IsAppIdPinned := this._vtable(this.IVirtualDesktopPinnedApps, 3) ; bool IsAppIdPinned(string appId);
        this.PinAppID := this._vtable(this.IVirtualDesktopPinnedApps, 4) ; void PinAppID(string appId);
        this.UnpinAppID := this._vtable(this.IVirtualDesktopPinnedApps, 5) ; void UnpinAppID(string appId);
        this.IsViewPinned := this._vtable(this.IVirtualDesktopPinnedApps, 6) ; bool IsViewPinned(IApplicationView applicationView);
        this.PinView := this._vtable(this.IVirtualDesktopPinnedApps, 7) ; void PinView(IApplicationView applicationView);
        this.UnpinView := this._vtable(this.IVirtualDesktopPinnedApps, 8) ; void UnpinView(IApplicationView applicationView);


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
        this.GetViewForHwnd := this._vtable(this.IApplicationViewCollection, 6) ; (IntPtr hwnd, out IApplicationView view);


        ;----------------------

        ; VarSetCapacity(IID_IVirtualDesktop, 16)
        ; this will never be garbage collected
        this.Ptr_IID_IVirtualDesktop := DllCall( "GlobalAlloc", "UInt",0x40, "UInt", 16, "Ptr")
        DllCall("Ole32.dll\CLSIDFromString", "Str", IID_IVirtualDesktop_, "Ptr", this.Ptr_IID_IVirtualDesktop)

    }
    ;dll methods start
    _dll_GetDesktops_Win10() {
        IObjectArray := 0
        DllCall(this.GetDesktops, "UPtr", this.IVirtualDesktopManagerInternal, "Ptr*", IObjectArray)
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
    getCount() { ;how many virtual desktops you now have
        return this._GetDesktops_Obj().GetCount()
    }

    goToDesktopNum(desktopNum) {
        IVirtualDesktop:=this._GetDesktops_Obj().GetAt(desktopNum)
        this._SwitchDesktop(IVirtualDesktop)

        if (this._isWindowFullScreen("A")) {
            timerFunc := ObjBindMethod(this, "_pleaseSwitchDesktop", desktopNum) ;https://www.autohotkey.com/docs/commands/SetTimer.htm#ExampleClass
            SetTimer % timerFunc, -50
        }

    }

    getDesktopNumOfWindow(wintitle)
    {
        found:=this._getFirstValidWindow(wintitle)
        if (!found) {
            return -1 ;for false
        }
        theHwnd:=found[1]

        IVirtualDesktop_ofWindow:=this._IVirtualDesktop_from_Hwnd(theHwnd)

        desktopNum:=this._desktopNum_from_IVirtualDesktop(IVirtualDesktop_ofWindow)
        return desktopNum
    }

    sendWindowToDesktopNum(wintitle,desktopNum,followYourWindow:=true,activateYourWindow:=true)
    {
        found:=this._getFirstValidWindow(wintitle)
        if (!found) {
            return -1 ;for false
        }
        theHwnd:=found[1]
        thePView:=found[2]

        IVirtualDesktop:=this._GetDesktops_Obj().GetAt(desktopNum)

        DllCall(this.MoveViewToDesktop, "ptr", this.IVirtualDesktopManagerInternal, "Ptr", thePView, "Ptr", IVirtualDesktop)

        if (followYourWindow) {
            this.goToDesktopNum(desktopNum)
        }
        if (activateYourWindow) {
            WinActivate, ahk_id %theHwnd%
        }
    }

    goToDesktopOfWindow(wintitle, activateYourWindow:=true) {
        desktopNum:=this.getDesktopNumOfWindow(wintitle)
        this.goToDesktopNum(desktopNum)

        if (activateYourWindow) {
            WinActivate, ahk_id %theHwnd%
        }
    }

    getCurrentDesktopNum()
    {
        IVirtualDesktop_ofCurrentDesktop := 0
        DllCall(this.GetCurrentDesktop, "UPtr", this.IVirtualDesktopManagerInternal, "Ptr*", IVirtualDesktop_ofCurrentDesktop)

        desktopNum:=this._desktopNum_from_IVirtualDesktop(IVirtualDesktop_ofCurrentDesktop)
        return desktopNum
    }

    sendWindowToCurrentDesktop(wintitle,activateYourWindow:=true)
    {
        desktopNum:=this.getCurrentDesktopNum()
        this.sendWindowToDesktopNum(wintitle, desktopNum, false, activateYourWindow)
    }

    createDesktop(goThere:=true) {
        IVirtualDesktop_ofNewDesktop:=0
        DllCall(this.Ptr_CreateDesktop, "UPtr", this.IVirtualDesktopManagerInternal, "Ptr*", IVirtualDesktop_ofNewDesktop)

        if (goThere) {
            ;we could assume that it's the rightmost desktop:
            ; desktopNum:=this.getCount()
            ;but I'm not risking it
            desktopNum:=this._desktopNum_from_IVirtualDesktop(IVirtualDesktop_ofNewDesktop)
            this.goToDesktopNum(desktopNum)
        }
    }

    createUntil(howMany, goToLastlyCreated:=true) {
        howManyThereAlreadyAre:=Vd.getCount()
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

        DllCall(this.Ptr_RemoveDesktop, "UPtr", this.IVirtualDesktopManagerInternal, "Ptr", IVirtualDesktop, "Ptr", IVirtualDesktop_fallback)
    }

    IsWindowPinned(wintitle) {
        found:=this._getFirstValidWindow(wintitle)
        if (!found) {
            return -1 ;for false
        }
        thePView:=found[2]

        viewIsPinned:=this._IsViewPinned(thePView)
        return viewIsPinned
    }
    TogglePinWindow(wintitle) {
        found:=this._getFirstValidWindow(wintitle)
        if (!found) {
            return -1 ;for false
        }
        thePView:=found[2]

        viewIsPinned:=this._IsViewPinned(thePView)
        if (viewIsPinned) {
            DllCall(this.UnPinView, "UPtr", this.IVirtualDesktopPinnedApps, "Ptr", thePView)
        } else {
            DllCall(this.PinView, "UPtr", this.IVirtualDesktopPinnedApps, "Ptr", thePView)
        }

    }
    PinWindow(wintitle) {
        found:=this._getFirstValidWindow(wintitle)
        if (!found) {
            return -1 ;for false
        }
        thePView:=found[2]

        DllCall(this.PinView, "UPtr", this.IVirtualDesktopPinnedApps, "Ptr", thePView)
    }
    UnPinWindow(wintitle) {
        found:=this._getFirstValidWindow(wintitle)
        if (!found) {
            return -1 ;for false
        }
        thePView:=found[2]

        DllCall(this.UnPinView, "UPtr", this.IVirtualDesktopPinnedApps, "Ptr", thePView)
    }

    ;actual methods end

    ;internal methods start
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
        IVirtualDesktop:=this._GetDesktops_Obj().GetAt(desktopNum)
        this._SwitchDesktop(IVirtualDesktop)
        ;this method is goToDesktopNum(), but without the recursion, to prevent recursion
    }

    _getFirstValidWindow(wintitle) {
        global CanViewMoveDesktops, IVirtualDesktopManagerInternal

        DetectHiddenWindows, on
        WinGet, outHwndList, List, % wintitle
        DetectHiddenWindows, off
        loop % outHwndList {
            if (!this._isValidWindow(outHwndList%A_Index%)) {
                continue
            }

            pView:=this._view_from_Hwnd(outHwndList%A_Index%)
            pfCanViewMoveDesktops := 0
            DllCall(this.CanViewMoveDesktops, "UPtr", this.IVirtualDesktopManagerInternal, "Ptr", pView, "int*", pfCanViewMoveDesktops) ; return value BOOL
            if (!pfCanViewMoveDesktops) {
                continue
            }
            ;we can finally return
            return [outHwndList%A_Index%, pView]
        }
        ;none found
        return false
    }

    _view_from_Hwnd(theHwnd) {
        pView := 0
        DllCall(this.GetViewForHwnd, "UPtr", this.IApplicationViewCollection, "Ptr", theHwnd, "Ptr*", pView)
        return pView
    }

    _desktopGUID_from_Hwnd(theHwnd) {
        VarSetCapacity(GUID_Desktop, 16)
        HRESULT := DllCall(this.GetWindowDesktopId, "UPtr", this.IVirtualDesktopManager, "Ptr", theHwnd, "UPtr", &GUID_Desktop)
        if (!(HRESULT==0)) {
            return false
        }

        desktopGUID:=this._string_from_GUID(GUID_Desktop)
        if (!desktopGUID) {
            return false
        }
        if (desktopGUID=="{00000000-0000-0000-0000-000000000000}") {
            return false
        }

        return desktopGUID
    }

    _IVirtualDesktop_from_Hwnd(theHwnd) {
        VarSetCapacity(GUID_Desktop, 16)
        HRESULT := DllCall(this.GetWindowDesktopId, "UPtr", this.IVirtualDesktopManager, "Ptr", theHwnd, "Ptr", &GUID_Desktop)
        if (!(HRESULT==0)) {
            return false
        }

        IVirtualDesktop_ofWindow:=0
        DllCall(this.FindDesktop, "UPtr", this.IVirtualDesktopManagerInternal, "Ptr", &GUID_Desktop, "Ptr*", IVirtualDesktop_ofWindow)

        return IVirtualDesktop_ofWindow
    }

    _desktopNum_from_IVirtualDesktop(IVirtualDesktop) {
        Desktops_Obj:=this._GetDesktops_Obj()
        Loop % Desktops_Obj.GetCount()
        {
            IVirtualDesktop_ofDesktop:=Desktops_Obj.GetAt(A_Index)

            if (IVirtualDesktop_ofDesktop == IVirtualDesktop) {
                return A_Index
            }
        }
        return -1 ;for false
    }

    _GetDesktops_Obj() {
        IObjectArray:=this._dll_GetDesktops()
        return new this.IObjectArray_Wrapper(IObjectArray, this.Ptr_IID_IVirtualDesktop)
    }

    _IsViewPinned(thePView) {
        viewIsPinned:=0
        DllCall(this.IsViewPinned, "UPtr", this.IVirtualDesktopPinnedApps, "Ptr", thePView, "Int*",viewIsPinned)
        return viewIsPinned
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

    _isValidWindow(hWnd)
    {
        DetectHiddenWindows, on

        breakToReturnFalse:
        loop 1 {
            WinGetTitle, title, ahk_id %hWnd%
            if (!title) {
                break breakToReturnFalse
            }

            WinGet, dwStyle, Style, ahk_id %hWnd%
            if ((dwStyle&0x08000000) || !(dwStyle&0x10000000)) {
                break breakToReturnFalse
            }
            WinGet, dwExStyle, ExStyle, ahk_id %hWnd%
            if (dwExStyle & 0x00000080) {
                break breakToReturnFalse
            }
            WinGetClass, szClass, ahk_id %hWnd%
            if (szClass = "TApplication") {
                break breakToReturnFalse
            }

            DetectHiddenWindows, off
            return true

        }
        DetectHiddenWindows, off
        return false
    }
    ;-------------------
    _vtable(ppv, index) {
        Return NumGet(NumGet(0+ppv)+A_PtrSize*index)
    }
    _string_from_GUID(Byref byref_GUID) {
        VarSetCapacity(strGUID, 38 * 2) ;38 is StrLen("{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}")
        DllCall("Ole32.dll\StringFromGUID2", "UPtr", &byref_GUID, "UPtr", &strGUID, "Int", 38 + 1)
        return StrGet(&strGUID, "UTF-16")
    }

    class IObjectArray_Wrapper {
        __New(IObjectArray, Ptr_IID_Interface) {
            this.IObjectArray:=IObjectArray
            this.Ptr_IID_Interface:=Ptr_IID_Interface

            this.Ptr_GetAt:=VD._vtable(IObjectArray,4)
            this.Ptr_GetCount:=VD._vtable(IObjectArray,3)
        }
        GetAt(oneBasedIndex) {
            Ptr_Interface:=0
            DllCall(this.Ptr_GetAt, "UPtr", this.IObjectArray, "UInt", oneBasedIndex - 1, "Ptr", this.Ptr_IID_Interface, "Ptr*", Ptr_Interface)
            return Ptr_Interface
        }
        GetCount() {
            Count := 0
            DllCall(this.Ptr_GetCount, "UPtr", this.IObjectArray, "UInt*", Count)
            return Count
        }

    }
    ;utility methods end

}
