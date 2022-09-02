; VD.getCurrentDesktopNum()
; VD.getDesktopNumOfWindow(wintitle) ;please use VD.goToDesktopOfWindow instead if you just want to go there.
; VD.getDesktopNumOfWindow(wintitle) ;returns 0 for "Show on all desktops"

; VD.getCount() ;how many virtual desktops you now have

; VD.goToDesktopNum(desktopNum)
; VD.goToDesktopOfWindow(wintitle, activateYourWindow:=true)

; VD.MoveWindowToDesktopNum(wintitle, desktopNum)
; VD.MoveWindowToCurrentDesktop(wintitle, activateYourWindow:=true)
; VD.MoveWindowToRelativeDesktopNum(wintitle, relative_desktopNum)

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

    ; #Include VD.ahk
    ; or
    ; #Include _VD.ahk
    ; ...{startup code}
    ; VD.init()

    ; VD.ahk : calls `VD.init()` on #Include
    ; _VD.ahk : `VD.init()` when you want, like after a GUI has rendered, for startup performance reasons
    init()
    {
        splitByDot:=StrSplit(A_OSVersion, ".")
        buildNumber:=splitByDot[3]
        if (buildNumber < 22000)
        {
            ; Windows 10
            IID_IVirtualDesktopManagerInternal_:="{F31574D6-B682-4CDC-BD56-1827860ABEC6}" ;https://github.com/MScholtes/VirtualDesktop/blob/812c321e286b82a10f8050755c94d21c4b69812f/VirtualDesktop.cs#L177-L191
            IID_IVirtualDesktop_:="{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}" ;https://github.com/MScholtes/VirtualDesktop/blob/812c321e286b82a10f8050755c94d21c4b69812f/VirtualDesktop.cs#L149-L150
            ;conditionally assign method to method
            this._dll_GetCurrentDesktop:=this._dll_GetCurrentDesktop_Win10
            this._dll_GetDesktops:=this._dll_GetDesktops_Win10
            this._dll_SwitchDesktop:=this._dll_SwitchDesktop_Win10
            this._dll_CreateDesktop:=this._dll_CreateDesktop_Win10
            this._dll_GetName:=this._dll_GetName_Win10
            this.RegisterDesktopNotifications:=this.RegisterDesktopNotifications_Win10
        }
        else
        {
            ; Windows 11
            IID_IVirtualDesktopManagerInternal_:="{B2F925B9-5A0F-4D2E-9F4D-2B1507593C10}" ;https://github.com/MScholtes/VirtualDesktop/blob/812c321e286b82a10f8050755c94d21c4b69812f/VirtualDesktop11.cs#L163-L185
            IID_IVirtualDesktop_:="{536D3495-B208-4CC9-AE26-DE8111275BF8}" ;https://github.com/MScholtes/VirtualDesktop/blob/812c321e286b82a10f8050755c94d21c4b69812f/VirtualDesktop11.cs#L149-L150
            ;conditionally assign method to method
            this._dll_GetCurrentDesktop:=this._dll_GetCurrentDesktop_Win11
            this._dll_GetDesktops:=this._dll_GetDesktops_Win11
            this._dll_SwitchDesktop:=this._dll_SwitchDesktop_Win11
            this._dll_GetName:=this._dll_GetName_Win11
            this.RegisterDesktopNotifications:=this.RegisterDesktopNotifications_Win11
        }
        ;----------------------
        this.IVirtualDesktopManager := ComObjCreate("{AA509086-5CA9-4C25-8F95-589D3C07B48A}", "{A5CD92FF-29BE-454C-8D04-D82879FB3F1B}")
        this.GetWindowDesktopId := this._vtable(this.IVirtualDesktopManager, 4)

        this.IServiceProvider := ComObjCreate("{C2F03A33-21F5-47FA-B4BB-156362A2F239}", "{6D5140C1-7436-11CE-8034-00AA006009FA}")

        this.IVirtualDesktopManagerInternal := ComObjQuery(this.IServiceProvider, "{C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B}", IID_IVirtualDesktopManagerInternal_)

        ; this.GetCount := this._vtable(this.IVirtualDesktopManagerInternal, 3 ; int GetCount();
        this.MoveViewToDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 4) ; void MoveViewToDesktop(object pView, IVirtualDesktop desktop);
        this.CanViewMoveDesktops := this._vtable(this.IVirtualDesktopManagerInternal, 5) ; bool CanViewMoveDesktops(IApplicationView view);
        this.GetCurrentDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 6) ; IVirtualDesktop GetCurrentDesktop(); || IVirtualDesktop GetCurrentDesktop(IntPtr hWndOrMon);
        if (buildNumber < 22000) {
            ;Windows 10 https://github.com/MScholtes/VirtualDesktop/blob/812c321e286b82a10f8050755c94d21c4b69812f/VirtualDesktop.cs#L177-L191

            this.GetDesktops := this._vtable(this.IVirtualDesktopManagerInternal, 7) ; void GetDesktops(out IObjectArray desktops);
            ; this.GetAdjacentDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 8) ; int GetAdjacentDesktop(IVirtualDesktop from, int direction, out IVirtualDesktop desktop);
            this.SwitchDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 9) ; void SwitchDesktop(IVirtualDesktop desktop);
            this.Ptr_CreateDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 10) ; IVirtualDesktop CreateDesktop();
            this.Ptr_RemoveDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 11) ; void RemoveDesktop(IVirtualDesktop desktop, IVirtualDesktop fallback);
            this.FindDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 12) ; IVirtualDesktop FindDesktop(ref Guid desktopid);
        } else if (buildNumber < 22489) {
            ;Windows 11 https://github.com/MScholtes/VirtualDesktop/blob/812c321e286b82a10f8050755c94d21c4b69812f/VirtualDesktop11.cs#L163-L185

            this.GetDesktops := this._vtable(this.IVirtualDesktopManagerInternal, 7) ; void GetDesktops(IntPtr hWndOrMon, out IObjectArray desktops);
            ; this.GetAdjacentDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 8) ; int GetAdjacentDesktop(IVirtualDesktop from, int direction, out IVirtualDesktop desktop);
            this.SwitchDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 9) ; void SwitchDesktop(IntPtr hWndOrMon, IVirtualDesktop desktop);
            this.Ptr_CreateDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 10) ; IVirtualDesktop CreateDesktop(IntPtr hWndOrMon);
            ; this.MoveDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 11) ; void MoveDesktop(IVirtualDesktop desktop, IntPtr hWndOrMon, int nIndex);
            this.Ptr_RemoveDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 12) ; void RemoveDesktop(IVirtualDesktop desktop, IVirtualDesktop fallback);
            this.FindDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 13) ; IVirtualDesktop FindDesktop(ref Guid desktopid);
        } else {
            ;Windows 11 Insider build 22489 https://github.com/MScholtes/VirtualDesktop/blob/9f3872e1275408a0802bdbe46df499bb7645dc87/VirtualDesktop11Insider.cs#L163-L186

            ; this.GetAllCurrentDesktops := this._vtable(this.IVirtualDesktopManagerInternal, 7) ; IObjectArray GetAllCurrentDesktops();
            this.GetDesktops := this._vtable(this.IVirtualDesktopManagerInternal, 8) ; void GetDesktops(IntPtr hWndOrMon, out IObjectArray desktops);
            ; this.GetAdjacentDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 9) ; int GetAdjacentDesktop(IVirtualDesktop from, int direction, out IVirtualDesktop desktop);
            this.SwitchDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 10) ; void SwitchDesktop(IntPtr hWndOrMon, IVirtualDesktop desktop);
            this.Ptr_CreateDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 11) ; IVirtualDesktop CreateDesktop(IntPtr hWndOrMon);
            ; this.MoveDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 12) ; void MoveDesktop(IVirtualDesktop desktop, IntPtr hWndOrMon, int nIndex);
            this.Ptr_RemoveDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 13) ; void RemoveDesktop(IVirtualDesktop desktop, IVirtualDesktop fallback);
            this.FindDesktop := this._vtable(this.IVirtualDesktopManagerInternal, 14) ; IVirtualDesktop FindDesktop(ref Guid desktopid);
        }

        ;https://github.com/MScholtes/VirtualDesktop/blob/812c321e286b82a10f8050755c94d21c4b69812f/VirtualDesktop.cs#L225-L234
        this.IVirtualDesktopPinnedApps := ComObjQuery(this.IServiceProvider, "{B5A399E7-1C87-46B8-88E9-FC5747B171BD}", "{4CE81583-1E4C-4632-A621-07A53543148F}")

        ; this.IsAppIdPinned := this._vtable(this.IVirtualDesktopPinnedApps, 3) ; bool IsAppIdPinned(string appId);
        ; this.PinAppID := this._vtable(this.IVirtualDesktopPinnedApps, 4) ; void PinAppID(string appId);
        ; this.UnpinAppID := this._vtable(this.IVirtualDesktopPinnedApps, 5) ; void UnpinAppID(string appId);
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
        this.Ptr_IID_IVirtualDesktop := DllCall( "GlobalAlloc", "UInt",0x00, "UInt", 16, "Ptr")
        DllCall("Ole32.dll\CLSIDFromString", "Str", IID_IVirtualDesktop_, "Ptr", this.Ptr_IID_IVirtualDesktop)

        ;----------------------

    }
    ;dll methods start
    _dll_GetCurrentDesktop_Win10() {
        IVirtualDesktop_ofCurrentDesktop := 0
        DllCall(this.GetCurrentDesktop, "UPtr", this.IVirtualDesktopManagerInternal, "Ptr*", IVirtualDesktop_ofCurrentDesktop)
        return IVirtualDesktop_ofCurrentDesktop
    }
    _dll_GetCurrentDesktop_Win11() {
        IVirtualDesktop_ofCurrentDesktop := 0
        DllCall(this.GetCurrentDesktop, "UPtr", this.IVirtualDesktopManagerInternal, "Ptr", 0, "Ptr*", IVirtualDesktop_ofCurrentDesktop)
        return IVirtualDesktop_ofCurrentDesktop
    }
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
    _dll_CreateDesktop_Win10() {
        IVirtualDesktop_ofNewDesktop:=0
        DllCall(this.Ptr_CreateDesktop, "UPtr", this.IVirtualDesktopManagerInternal, "Ptr*", IVirtualDesktop_ofNewDesktop)
        return IVirtualDesktop_ofNewDesktop
    }
    _dll_CreateDesktop_Win11() {
        IVirtualDesktop_ofNewDesktop:=0
        DllCall(this.Ptr_CreateDesktop, "UPtr", this.IVirtualDesktopManagerInternal, "Ptr", 0, "Ptr*", IVirtualDesktop_ofNewDesktop)
        return IVirtualDesktop_ofNewDesktop
    }
    _dll_GetName_Win10(IVirtualDesktop) {
        QueryInterface:=this._vtable(IVirtualDesktop, 0)
        VarSetCapacity(CLSID, 16)
        DllCall("Ole32.dll\CLSIDFromString", "Str","{31EBDE3F-6EC3-4CBD-B9FB-0EF6D09B41F4}", "Ptr",&CLSID)
        DllCall(QueryInterface, "UPtr",IVirtualDesktop, "Ptr",&CLSID, "Ptr*", IVirtualDesktop2)

        GetName:=this._vtable(IVirtualDesktop2,5)
        DllCall(GetName, "UPtr", IVirtualDesktop2, "Ptr*", Handle_DesktopName)
        if (Handle_DesktopName==0) {
            return "" ;you can't have empty desktopName so this can represent error
        }
        Ptr_DesktopName:=DllCall("combase\WindowsGetStringRawBuffer", "Ptr",Handle_DesktopName, "UInt*",length, "Ptr")
        desktopName:=StrGet(Ptr_DesktopName+0,"UTF-16")
        return desktopName
    }
    /* _dll_GetName_Win10(IVirtualDesktop) {
        GetId := this._vtable(IVirtualDesktop, 4)
        VarSetCapacity(GUID_Desktop, 16)
        DllCall(GetId, "UPtr",IVirtualDesktop, "Ptr",&GUID_Desktop)

        strGUID:=this._string_from_GUID(GUID_Desktop)
        KeyName:="HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops\Desktops\" strGUID
        ;RegRead gives "" on key not found, convenient
        RegRead, desktopName, % KeyName, % "Name"
        ; I don't like this because registry can be edited, then it wouldn't reflect the desktopName
        return desktopName
    }
    */
    _dll_GetName_Win11(IVirtualDesktop) {
        GetName:=this._vtable(IVirtualDesktop,6)
        DllCall(GetName, "UPtr", IVirtualDesktop, "Ptr*", Handle_DesktopName)
        if (Handle_DesktopName==0) {
            return "" ;you can't have empty desktopName so this can represent error
        }
        Ptr_DesktopName:=DllCall("combase\WindowsGetStringRawBuffer", "Ptr",Handle_DesktopName, "UInt*",length, "Ptr")
        desktopName:=StrGet(Ptr_DesktopName+0,"UTF-16")
        return desktopName
    }
    ;dll methods end

    ;actual methods start
    getCount() { ;how many virtual desktops you now have
        return this._GetDesktops_Obj().GetCount()
    }

    goToDesktopNum(desktopNum) {
        IVirtualDesktop:=this._GetDesktops_Obj().GetAt(desktopNum)
        this._SwitchIVirtualDesktop(IVirtualDesktop)

        if (this._isWindowFullScreen("A")) {
            timerFunc := ObjBindMethod(this, "_pleaseSwitchDesktop", desktopNum) ;https://www.autohotkey.com/docs/commands/SetTimer.htm#ExampleClass
            SetTimer % timerFunc, -50
        }

    }

    getNameFromDesktopNum(desktopNum) {
        IVirtualDesktop:=this._GetDesktops_Obj().GetAt(desktopNum)
        _name:=this._dll_GetName(IVirtualDesktop)
        return _name
        ; if (_name=="") {
            ; return "Desktop " _name
        ; }
    }

    getDesktopNumOfWindow(wintitle) {
        found:=this._getFirstValidWindow(wintitle)
        if (!found) {
            return -1 ;for false
        }
        theHwnd:=found[1]

        desktopNum_ofWindow:=this._desktopNum_from_Hwnd(theHwnd)
        return desktopNum_ofWindow ; 0 for "Show on all desktops"
    }

    goToDesktopOfWindow(wintitle, activateYourWindow:=true) {
        found:=this._getFirstValidWindow(wintitle)
        if (!found) {
            return -1 ;for false
        }
        theHwnd:=found[1]

        desktopNum_ofWindow:=this._desktopNum_from_Hwnd(theHwnd)
        this.goToDesktopNum(desktopNum_ofWindow)

        if (activateYourWindow) {
            WinActivate, ahk_id %theHwnd%
        }
    }

    MoveWindowToDesktopNum(wintitle, desktopNum)
    {
        found:=this._getFirstValidWindow(wintitle)
        if (!found) {
            return -1 ;for false
        }
        theHwnd:=found[1]
        thePView:=found[2]

        IVirtualDesktop:=this._GetDesktops_Obj().GetAt(desktopNum)

        this._MoveView_to_IVirtualDesktop(thePView, IVirtualDesktop)
    }

    MoveWindowToRelativeDesktopNum(wintitle, relative_desktopNum)
    {
        found:=this._getFirstValidWindow(wintitle)
        if (!found) {
            return -1 ;for false
        }
        theHwnd:=found[1]
        thePView:=found[2]

        desktopNum_ofWindow:=this._desktopNum_from_Hwnd(theHwnd)

        Desktops_Obj:=this._GetDesktops_Obj()
        count_Desktops:=Desktops_Obj.GetCount()

        absolute_desktopNum:=desktopNum_ofWindow + relative_desktopNum
        ;// The 1-based indices wrap around on the first and last desktop.
        ;// say count_Desktops:=3
        absolute_desktopNum:=Mod(absolute_desktopNum, count_Desktops)
        ; 4 -> 1
        if (absolute_desktopNum <= 0) {
            ; 0 -> 3
            absolute_desktopNum:=absolute_desktopNum + count_Desktops
        }

        IVirtualDesktop:=Desktops_Obj.GetAt(absolute_desktopNum)

        this._MoveView_to_IVirtualDesktop(thePView, IVirtualDesktop)
        return absolute_desktopNum
    }

    MoveWindowToCurrentDesktop(wintitle, activateYourWindow:=true)
    {
        found:=this._getFirstValidWindow(wintitle)
        if (!found) {
            return -1 ;for false
        }
        theHwnd:=found[1]
        thePView:=found[2]

        currentDesktopNum:=this.getCurrentDesktopNum()
        IVirtualDesktop:=this._GetDesktops_Obj().GetAt(currentDesktopNum)

        this._MoveView_to_IVirtualDesktop(thePView, IVirtualDesktop)

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

    ; COM class start ;https://github.com/Ciantic/VirtualDesktopAccessor/blob/5bc1bbaab247b5d72e70abc9432a15275fd2d229/VirtualDesktopAccessor/dllmain.h#L718-L794
    _QueryInterface_Win10(riid, ppvObject) {
        if (!ppvObject) {
            return 0x80070057 ;E_INVALIDARG
        }

        str_IID_IUnknown:="{00000000-0000-0000-C000-000000000046}"
        str_IID_IVirtualDesktopNotification:="{C179334C-4295-40D3-BEA1-C654D965605A}"

        VarSetCapacity(someStr, 40)
        DllCall("Ole32\StringFromGUID2", "Ptr", riid, "Ptr",&someStr, "Ptr",40)
        str_riid:=StrGet(&someStr)

        if (str_riid==str_IID_IUnknown || str_riid==str_IID_IVirtualDesktopNotification) {
            NumPut(this, ppvObject+0, 0, "Ptr")
            VD._AddRef_Same.Call(this)
            return 0 ;S_OK
        }
        ; *ppvObject = NULL;
        NumPut(0, ppvObject+0, 0, "Ptr")
        return 0x80004002 ;E_NOINTERFACE

        ; // Always set out parameter to NULL, validating it first.
        ; if (!ppvObject)
            ; return E_INVALIDARG;
        ; *ppvObject = NULL;
;
        ; if (riid == IID_IUnknown || riid == IID_IVirtualDesktopNotification)
        ; {
            ; // Increment the reference count and return the pointer.
            ; *ppvObject = (LPVOID)this;
            ; AddRef();
            ; return S_OK;
        ; }
        ; return E_NOINTERFACE;
    }
    _AddRef_Same() {
        refCount:=NumGet(this+0, A_PtrSize, "UInt")
        refCount++
        NumPut(refCount, this+0, A_PtrSize, "UInt")
        ; NumPut(this + 4)
        ; refCount:=

        ; return InterlockedIncrement(&_referenceCount);
        return refCount
    }
    _Release_Same() {
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
    _VirtualDesktopCreated_Win10(pDesktop) {
        ; Tooltip % 11111
        VD.VirtualDesktopCreated.Call(VD._desktopNum_from_IVirtualDesktop(pDesktop))
        return 0 ;S_OK
    }
    _VirtualDesktopDestroyBegin_Win10(pDesktopDestroyed, pDesktopFallback) {
        ; Tooltip % 22222
        VD.VirtualDesktopDestroyBegin.Call(VD._desktopNum_from_IVirtualDesktop(pDesktopDestroyed), VD._desktopNum_from_IVirtualDesktop(pDesktopFallback))
        return 0 ;S_OK
    }
    _VirtualDesktopDestroyFailed_Win10(pDesktopDestroyed, pDesktopFallback) {
        ; Tooltip % 33333
        VD.VirtualDesktopDestroyFailed.Call(VD._desktopNum_from_IVirtualDesktop(pDesktopDestroyed), VD._desktopNum_from_IVirtualDesktop(pDesktopFallback))
        return 0 ;S_OK
    }
    _VirtualDesktopDestroyed_Win10(pDesktopDestroyed, pDesktopFallback) {
        ; Tooltip % 44444
        VD.VirtualDesktopDestroyed.Call(VD._desktopNum_from_IVirtualDesktop(pDesktopDestroyed), VD._desktopNum_from_IVirtualDesktop(pDesktopFallback))
        return 0 ;S_OK
    }
    _ViewVirtualDesktopChanged_Win10(pView) {
        ; Tooltip % 55555
        VD.ViewVirtualDesktopChanged.Call(pView)
        return 0 ;S_OK
    }
    _CurrentVirtualDesktopChanged_Win10(pDesktopOld, pDesktopNew) {
        VD.CurrentVirtualDesktopChanged.Call(VD._desktopNum_from_IVirtualDesktop(pDesktopOld), VD._desktopNum_from_IVirtualDesktop(pDesktopNew))
        return 0 ;S_OK
    }
    RegisterDesktopNotifications_Win10() { ;https://github.com/Ciantic/VirtualDesktopAccessor/blob/5bc1bbaab247b5d72e70abc9432a15275fd2d229/VirtualDesktopAccessor/dllmain.h#L718-L794
        methods:=DllCall("GlobalAlloc", "Uint",0x00, "Uint",9*A_PtrSize) ;PLEASE DON'T GARBAGE COLLECT IT, this took me hours to debug, I was lucky ahkv2 garbage collected slowly
        NumPut(RegisterCallback("VD._QueryInterface_Win10", "F"), methods+0, 0*A_PtrSize, "Ptr")
        NumPut(RegisterCallback("VD._AddRef_Same", "F"), methods+0, 1*A_PtrSize, "Ptr")
        NumPut(RegisterCallback("VD._Release_Same", "F"), methods+0, 2*A_PtrSize, "Ptr")
        NumPut(RegisterCallback("VD._VirtualDesktopCreated_Win10", "F"), methods+0, 3*A_PtrSize, "Ptr")
        NumPut(RegisterCallback("VD._VirtualDesktopDestroyBegin_Win10", "F"), methods+0, 4*A_PtrSize, "Ptr")
        NumPut(RegisterCallback("VD._VirtualDesktopDestroyFailed_Win10", "F"), methods+0, 5*A_PtrSize, "Ptr")
        NumPut(RegisterCallback("VD._VirtualDesktopDestroyed_Win10", "F"), methods+0, 6*A_PtrSize, "Ptr")
        NumPut(RegisterCallback("VD._ViewVirtualDesktopChanged_Win10", "F"), methods+0, 7*A_PtrSize, "Ptr")
        NumPut(RegisterCallback("VD._CurrentVirtualDesktopChanged_Win10", "F"), methods+0, 8*A_PtrSize, "Ptr")

        this.RegisterDesktopNotifications_Same(methods)
    }
    ;COM class for Win11
    _QueryInterface_Win11(riid, ppvObject) {
        if (!ppvObject) {
            return 0x80070057 ;E_INVALIDARG
        }

        str_IID_IUnknown:="{00000000-0000-0000-C000-000000000046}"
        str_IID_IVirtualDesktopNotification:="{CD403E52-DEED-4C13-B437-B98380F2B1E8}"

        VarSetCapacity(someStr, 40)
        DllCall("Ole32\StringFromGUID2", "Ptr", riid, "Ptr",&someStr, "Ptr",40)
        str_riid:=StrGet(&someStr)

        if (str_riid==str_IID_IUnknown || str_riid==str_IID_IVirtualDesktopNotification) {
            NumPut(this, ppvObject+0, 0, "Ptr")
            VD._AddRef_Same.Call(this)
            return 0 ;S_OK
        }
        ; *ppvObject = NULL;
        NumPut(0, ppvObject+0, 0, "Ptr")
        return 0x80004002 ;E_NOINTERFACE
    }
    _VirtualDesktopCreated_Win11(p0, pDesktop) {
        ; Tooltip % 11111
        VD.VirtualDesktopCreated.Call(VD._desktopNum_from_IVirtualDesktop(pDesktop))
        return 0 ;S_OK
    }
    _VirtualDesktopDestroyBegin_Win11(p0, pDesktopDestroyed, pDesktopFallback) {
        ; Tooltip % 22222
        VD.VirtualDesktopDestroyBegin.Call(VD._desktopNum_from_IVirtualDesktop(pDesktopDestroyed), VD._desktopNum_from_IVirtualDesktop(pDesktopFallback))
        return 0 ;S_OK
    }
    _VirtualDesktopDestroyFailed_Win11(p0, pDesktopDestroyed, pDesktopFallback) {
        ; Tooltip % 33333
        VD.VirtualDesktopDestroyFailed.Call(VD._desktopNum_from_IVirtualDesktop(pDesktopDestroyed), VD._desktopNum_from_IVirtualDesktop(pDesktopFallback))
        return 0 ;S_OK
    }
    _VirtualDesktopDestroyed_Win11(p0, pDesktopDestroyed, pDesktopFallback) {
        ; Tooltip % 44444
        VD.VirtualDesktopDestroyed.Call(VD._desktopNum_from_IVirtualDesktop(pDesktopDestroyed), VD._desktopNum_from_IVirtualDesktop(pDesktopFallback))
        return 0 ;S_OK
    }
    _ViewVirtualDesktopChanged_Win11(pView) {
        ; Tooltip % 55555
        VD.ViewVirtualDesktopChanged.Call(pView)
        return 0 ;S_OK
    }
    _CurrentVirtualDesktopChanged_Win11(p0, pDesktopOld, pDesktopNew) {
        VD.CurrentVirtualDesktopChanged.Call(VD._desktopNum_from_IVirtualDesktop(pDesktopOld), VD._desktopNum_from_IVirtualDesktop(pDesktopNew))
        return 0 ;S_OK
    }
    _No_Op() {
    }
    RegisterDesktopNotifications_Win11() {
        methods:=DllCall("GlobalAlloc", "Uint",0x00, "Uint",13*A_PtrSize) ;PLEASE DON'T GARBAGE COLLECT IT, this took me hours to debug, I was lucky ahkv2 garbage collected slowly
        ; Thanks to mntone for IID and signatures https://mntone.hateblo.jp/entry/2021/05/23/121028#IVirtualDesktopNotification-3
        ; Thanks to NyaMisty for explanation https://github.com/mntone/VirtualDesktop/pull/1#issuecomment-922269079
        NumPut(RegisterCallback("VD._QueryInterface_Win11", "F"), methods+0, 0*A_PtrSize, "Ptr")
        NumPut(RegisterCallback("VD._AddRef_Same", "F"), methods+0, 1*A_PtrSize, "Ptr")
        NumPut(RegisterCallback("VD._Release_Same", "F"), methods+0, 2*A_PtrSize, "Ptr")
        NumPut(RegisterCallback("VD._VirtualDesktopCreated_Win11", "F"), methods+0, 3*A_PtrSize, "Ptr")
        NumPut(RegisterCallback("VD._VirtualDesktopDestroyBegin_Win11", "F"), methods+0, 4*A_PtrSize, "Ptr")
        NumPut(RegisterCallback("VD._VirtualDesktopDestroyFailed_Win11", "F"), methods+0, 5*A_PtrSize, "Ptr")
        NumPut(RegisterCallback("VD._VirtualDesktopDestroyed_Win11", "F"), methods+0, 6*A_PtrSize, "Ptr")
        ; NumPut(RegisterCallback("VD._VirtualDesktopIsPerMonitorChanged_Win11", "F"), methods+0, 7*A_PtrSize, "Ptr")
        NumPut(RegisterCallback("VD._No_Op", "F"), methods+0, 7*A_PtrSize, "Ptr")
        ; NumPut(RegisterCallback("VD._VirtualDesktopMoved_Win11", "F"), methods+0, 8*A_PtrSize, "Ptr")
        NumPut(RegisterCallback("VD._No_Op", "F"), methods+0, 8*A_PtrSize, "Ptr")
        ; NumPut(RegisterCallback("VD._VirtualDesktopNameChanged_Win11", "F"), methods+0, 9*A_PtrSize, "Ptr")
        NumPut(RegisterCallback("VD._No_Op", "F"), methods+0, 9*A_PtrSize, "Ptr")
        NumPut(RegisterCallback("VD._ViewVirtualDesktopChanged_Win11", "F"), methods+0, 10*A_PtrSize, "Ptr")
        NumPut(RegisterCallback("VD._CurrentVirtualDesktopChanged_Win11", "F"), methods+0, 11*A_PtrSize, "Ptr")
        ; NumPut(RegisterCallback("VD._VirtualDesktopWallpaperChanged_Win11", "F"), methods+0, 12*A_PtrSize, "Ptr")
        NumPut(RegisterCallback("VD._No_Op", "F"), methods+0, 12*A_PtrSize, "Ptr")

        this.RegisterDesktopNotifications_Same(methods)
    }
    RegisterDesktopNotifications_Same(methods) {
        obj:=DllCall("GlobalAlloc", "Uint",0x00, "Uint",A_PtrSize + 4) ;PLEASE DON'T GARBAGE COLLECT IT, this took me hours to debug, I was lucky ahkv2 garbage collected slowly
        NumPut(methods, obj+0, 0, "Ptr")
        NumPut(0, obj+0, A_PtrSize, "UInt") ;refCount

        pDesktopNotificationService := ComObjQuery(this.IServiceProvider, "{A501FDEC-4A09-464C-AE4E-1B9C21B84918}", "{0CD45E71-D927-4F15-8B0A-8FEF525337BF}")
        Register:=this._vtable(pDesktopNotificationService, 3)
        HRESULT:=DllCall(Register,"UPtr",pDesktopNotificationService, "Ptr",obj, "Uint*",pdwCookie:=0)
        ; ok1:=ErrorLevel
        ; ok2:=A_LastError
        ; ok:=0

        ; HRESULT hrNotificationService = pServiceProvider->QueryService(
		; CLSID_IVirtualNotificationService,
		; __uuidof(IVirtualDesktopNotificationService),
		; (PVOID*)&pDesktopNotificationService);
    }

    VirtualDesktopCreated(desktopNum) {
    }
    VirtualDesktopDestroyBegin(desktopNum_Destroyed, desktopNum_Fallback) {
    }
    VirtualDesktopDestroyFailed(desktopNum_Destroyed, desktopNum_Fallback) {
    }
    VirtualDesktopDestroyed(desktopNum_Destroyed, desktopNum_Fallback) {
    }
    ViewVirtualDesktopChanged(pView) {
    }
    CurrentVirtualDesktopChanged(desktopNum_Old, desktopNum_New) {
    }

    ; <Run in VD
    startShellMessage() {
        ; https://www.autohotkey.com/boards/viewtopic.php?t=63424#p271528
        DllCall("RegisterShellHookWindow", "Ptr", A_ScriptHwnd)
        MsgNum := DllCall("RegisterWindowMessage", "Str", "SHELLHOOK")
        OnMessage(MsgNum, func("VD.ShellMessage").bind(this))

        ; this.map_title_class:={"":{"":{"Hourglass.exe":func("VD.callback").bind(this, 2)}}}
        this.map_title_class:={}
    }

    Run(Target, WorkingDir, this_titleName, this_class, this_processName, desktopNum) {
        this.addToWaitNewWindow(this_titleName, this_class, this_processName, func("VD.callback_MoveWindow").bind(this, desktopNum))

        Run % Target, % WorkingDir
    }

    Run_lock_VD(Target, WorkingDir, this_titleName, this_class, this_processName, window_desktopNum, your_desktopNum) {
        this.addToWaitNewWindow(this_titleName, this_class, this_processName, func("VD.callback_MoveWindow_lockVD").bind(this, [window_desktopNum, your_desktopNum]))

        Run % Target, % WorkingDir
    }

    addToWaitNewWindow(this_titleName, this_class, this_processName, callback) {
        map_class_processName:=this.map_title_class.HasKey(this_titleName) ? this.map_title_class[this_titleName] : this.map_title_class[this_titleName]:={}
        map_processName_data:=map_class_processName.HasKey(this_class) ? map_class_processName[this_class] : map_class_processName[this_class]:={}
        arrOfCallback:=map_processName_data.HasKey(this_processName) ? map_processName_data[this_processName] : map_processName_data[this_processName]:=[]
        arrOfCallback.Push(callback)
    }

    callback_MoveWindow(desktopNum, hwnd) {
        WinActivate % "ahk_id " hwnd
        this.MoveWindowToDesktopNum("ahk_id " hwnd,desktopNum)
    }

    callback_MoveWindow_lockVD(tuple, hwnd) {
        window_desktopNum:=tuple[1]
        your_desktopNum:=tuple[2]
        WinActivate % "ahk_id " hwnd
        this.goToDesktopNum(your_desktopNum)
        WinActivate % "ahk_id " hwnd
        this.MoveWindowToDesktopNum("ahk_id " hwnd,window_desktopNum)
        WinActivate % "ahk_id " hwnd
    }

    ShellMessage(nCode, wParam, lParam) {
        Critical ;this is what makes many callbacks AT THE SAME TIME possible
        Sleep 100 ;necessary

        if (nCode < 0) {
            MsgBox % nCode " oh no"
        }

        hwnd:=false
        switch nCode {
            case 1:
                hwnd:=wParam
            case 12:
                MsgBox % "HSHELL_APPCOMMAND 12 oh no"
        }

        if (hwnd) { ; HSHELL_WINDOWCREATED := 1, HSHELL_MONITORCHANGED := 16
            bak_DetectHiddenWindows := A_DetectHiddenWindows
            DetectHiddenWindows, ON ;very important

            arrOfCallback:=false

            WinGetTitle, this_title, % "ahk_id " hwnd
            for subString_title, map_class_processName in this.map_title_class {
                if (InStr(this_title, subString_title, true)) {
                    WinGetClass, this_class, % "ahk_id " hwnd
                    for subString_class, map_processName_data in map_class_processName {
                        if (InStr(this_class, subString_class, true)) {
                            WinGet, this_processName, ProcessName, % "ahk_id " hwnd
                            for subString_processName, possibly_arrOfCallback in map_processName_data {
                                if (InStr(this_processName, subString_processName, true)) {
                                    arrOfCallback:=possibly_arrOfCallback
                                    break
                                }
                            }
                            break
                        }
                    }
                    break
                }
            }

            DetectHiddenWindows % bak_DetectHiddenWindows

            if (arrOfCallback) {
                callback:=arrOfCallback[1]
                callback.Call(hwnd)

                if (arrOfCallback.Length() > 1) {
                    arrOfCallback.RemoveAt(1)
                } else if (map_processName_data.Count() > 1) {
                    map_processName_data.Delete(subString_processName)
                } else if (map_class_processName.Count() > 1) {
                    map_class_processName.Delete(subString_class)
                } else {
                    this.map_title_class.Delete(subString_title)
                }

            }
        }
    }
    ; Run in VD>

    ;actual methods end

    ;internal methods start
    _MoveView_to_IVirtualDesktop(thePView, IVirtualDesktop) {
        DllCall(this.MoveViewToDesktop, "ptr", this.IVirtualDesktopManagerInternal, "Ptr", thePView, "Ptr", IVirtualDesktop)
        DllCall(this.MoveViewToDesktop, "ptr", this.IVirtualDesktopManagerInternal, "Ptr", thePView, "Ptr", IVirtualDesktop)
        this._activateWindowUnder()
    }
    _SwitchIVirtualDesktop(IVirtualDesktop) {
        ;activate taskbar before ;https://github.com/mzomparelli/zVirtualDesktop/issues/59#issuecomment-317613971
        WinActivate, ahk_class Shell_TrayWnd
        WinWaitActive, ahk_class Shell_TrayWnd

        this._dll_SwitchDesktop(IVirtualDesktop)
        this._dll_SwitchDesktop(IVirtualDesktop)
        this._activateWindowUnder()
    }

    _pleaseSwitchDesktop(desktopNum) {
        ;IVirtualDesktop should be calculated again because IVirtualDesktop could have changed
        ;what we want is the same desktopNum
        IVirtualDesktop:=this._GetDesktops_Obj().GetAt(desktopNum)
        this._SwitchIVirtualDesktop(IVirtualDesktop)
        ;this method is goToDesktopNum(), but without the recursion, to prevent recursion
    }

    _activateWindowUnder() {
        ; if this doesn't work
        ;try https://www.autohotkey.com/boards/viewtopic.php?t=28760#p326541
        found:=this._getFirstWindowUnder()
        if (!found) {
            return -1 ;for false
        }
        theHwnd:=found[1]
        ; MsgBox % theHwnd

        WinGet, OutputVar_MinMax, MinMax, % "ahk_id " theHwnd
        if (!(OutputVar_MinMax==-1)) {
            WinActivate, ahk_id %theHwnd%
        }
    }

    _getFirstWindowUnder() {
        returnValue:=false
        WinGet, outHwndList, List
        loop % outHwndList {
            if (pView:=this._isValidWindow(outHwndList%A_Index%)) {
                returnValue:=[outHwndList%A_Index%, pView]
                break
            }
        }
        return returnValue
    }

    _getFirstValidWindow(wintitle) {

        bak_DetectHiddenWindows:=A_DetectHiddenWindows
        bak_TitleMatchMode:=A_TitleMatchMode
        DetectHiddenWindows, on
        SetTitleMatchMode, 2
        WinGet, outHwndList, List, % wintitle

        returnValue:=false
        loop % outHwndList {
            if (pView:=this._isValidWindow(outHwndList%A_Index%)) {
                returnValue:=[outHwndList%A_Index%, pView]
                break
            }
        }

        SetTitleMatchMode % bak_TitleMatchMode
        DetectHiddenWindows % bak_DetectHiddenWindows
        return returnValue
    }

    _view_from_Hwnd(theHwnd) {
        pView := 0
        DllCall(this.GetViewForHwnd, "UPtr", this.IApplicationViewCollection, "Ptr", theHwnd, "Ptr*", pView)
        return pView
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
        return 0 ;for "Show on all desktops"
    }

    _desktopNum_from_Hwnd(theHwnd) {
        IVirtualDesktop:=this._IVirtualDesktop_from_Hwnd(theHwnd)
        desktopNum:=this._desktopNum_from_IVirtualDesktop(IVirtualDesktop)
        return desktopNum
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

    _isValidWindow(hWnd) ;returns pView if succeeded
    {
        ; DetectHiddenWindows, on ;this is needed, but for optimization the caller will do it

        returnValue:=false
        breakToReturnFalse:
        loop 1 {
            ; WinGetTitle, title, ahk_id %hWnd%
            WinGetTitle, title, % "ahk_id " hwnd
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

            pView:=this._view_from_Hwnd(hWnd)
            if (!pView) {
                break breakToReturnFalse
            }
            pfCanViewMoveDesktops := 0
            DllCall(this.CanViewMoveDesktops, "UPtr", this.IVirtualDesktopManagerInternal, "Ptr", pView, "int*", pfCanViewMoveDesktops) ; return value BOOL
            if (!pfCanViewMoveDesktops) {
                break breakToReturnFalse
            }

            returnValue:=pView
        }
        ; DetectHiddenWindows, off ;this is needed, but for optimization the caller will do it
        return returnValue
    }
    ;-------------------
    _vtable(ppv, index) {
        Return NumGet(NumGet(0+ppv)+A_PtrSize*index)
    }
    ; _string_from_GUID(Byref byref_GUID) {
        ; VarSetCapacity(strGUID, 38 * 2) ;38 is StrLen("{FF72FFDD-BE7E-43FC-9C03-AD81681E88E4}")
        ; DllCall("Ole32.dll\StringFromGUID2", "UPtr", &byref_GUID, "UPtr", &strGUID, "Int", 38 + 1)
        ; return StrGet(&strGUID, "UTF-16")
    ; }

    class IObjectArray_Wrapper {
        __New(IObjectArray, Ptr_IID_Interface) {
            this.IObjectArray:=IObjectArray
            this.Ptr_IID_Interface:=Ptr_IID_Interface

            this.Ptr_GetAt:=VD._vtable(IObjectArray,4)
        }
        __Delete() {
            ;IUnknown::Release
            Ptr_Release:=VD._vtable(this.IObjectArray,2)
            DllCall(Ptr_Release, "UPtr", this.IObjectArray)
        }
        GetAt(oneBasedIndex) {
            Ptr_Interface:=0
            DllCall(this.Ptr_GetAt, "UPtr", this.IObjectArray, "UInt", oneBasedIndex - 1, "Ptr", this.Ptr_IID_Interface, "Ptr*", Ptr_Interface)
            return Ptr_Interface
        }
        GetCount() {
            Ptr_GetCount:=VD._vtable(this.IObjectArray,3)
            Count := 0
            DllCall(Ptr_GetCount, "UPtr", this.IObjectArray, "UInt*", Count)
            return Count
        }

    }
    ;utility methods end

}
