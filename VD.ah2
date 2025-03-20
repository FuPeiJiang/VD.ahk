class VD {
  static dummyStatic1 := VD._init()
  static animation_on := false
  static _init() {
    OS_Version := FileGetVersion(A_WinDir "\System32\twinui.pcshell.dll")
    splitByDot := StrSplit(OS_Version, ".")
    buildNumber := splitByDot[3] + 0
    revisionNumber := splitByDot[4] + 0
    if (buildNumber < 20348) {
      IID_IVirtualDesktopManagerInternal_str := "{f31574d6-b682-4cdc-bd56-1827860abec6}"
      IID_IVirtualDesktop_str := "{ff72ffdd-be7e-43fc-9c03-ad81681e88e4}"
      this.IID_IVirtualDesktopNotification_n1 := 4671150449476842316
      this.IID_IVirtualDesktopNotification_n2 := 6512317045282349502
      idx_MoveViewToDesktop := 4
      idx_GetCurrentDesktop := 6
      idx_GetDesktops := 7
      idx_SwitchDesktop := 9
      idx_CreateDesktop := 10
      idx_RemoveDesktop := 11
      this.idx_SwitchDesktopWithAnimation := -1
      this._dll_MoveViewToDesktop := this._dll_MoveViewToDesktop_normal
      this._dll_GetCurrentDesktop := this._dll_GetCurrentDesktop_normal
      this._dll_GetDesktops := this._dll_GetDesktops_normal
      this._dll_SwitchDesktop := this._dll_SwitchDesktop_normal
      this._dll_CreateDesktop := this._dll_CreateDesktop_normal
      this._dll_RemoveDesktop := this._dll_RemoveDesktop_normal
      this.idx_GetId := 4
      this.idx_VirtualDesktopNameChanged := -1
      this.idx_SetDesktopName := -1
      this.idx_GetName := -1
      this.idx_VirtualDesktopWallpaperChanged := -1
      this.idx_SetDesktopWallpaper := -1
      this.idx_GetWallpaper := -1
      this.idx_VirtualDesktopCreated := 3
      this.idx_VirtualDesktopDestroyBegin := 4
      this.idx_VirtualDesktopDestroyFailed := 5
      this.idx_VirtualDesktopDestroyed := 6
      this.idx_ViewVirtualDesktopChanged := 7
      this.idx_CurrentVirtualDesktopChanged := 8
      this._dll_VirtualDesktopCreated := VD._dll_VirtualDesktopCreated_normal
      this._dll_VirtualDesktopDestroyBegin := VD._dll_VirtualDesktopDestroyBegin_normal
      this._dll_VirtualDesktopDestroyFailed := VD._dll_VirtualDesktopDestroyFailed_normal
      this._dll_VirtualDesktopDestroyed := VD._dll_VirtualDesktopDestroyed_normal
      this._dll_ViewVirtualDesktopChanged := VD._dll_ViewVirtualDesktopChanged_normal
      this._dll_CurrentVirtualDesktopChanged := VD._dll_CurrentVirtualDesktopChanged_normal
      this.IVirtualDesktopNotification_methods_count := 9
    } else if (buildNumber < 22000) {
      IID_IVirtualDesktopManagerInternal_str := "{094afe11-44f2-4ba0-976f-29a97e263ee0}"
      IID_IVirtualDesktop_str := "{62fdf88b-11ca-4afb-8bd8-2296dfae49e2}"
      this.IID_IVirtualDesktopNotification_n1 := 4844864968146173457
      this.IID_IVirtualDesktopNotification_n2 := 6287598790844042150
      idx_MoveViewToDesktop := 4
      idx_GetCurrentDesktop := 6
      idx_GetDesktops := 7
      idx_SwitchDesktop := 9
      idx_CreateDesktop := 10
      idx_RemoveDesktop := 11
      this.idx_SwitchDesktopWithAnimation := -1
      this._dll_MoveViewToDesktop := this._dll_MoveViewToDesktop_normal
      this._dll_GetCurrentDesktop := this._dll_GetCurrentDesktop_HMONITOR
      this._dll_GetDesktops := this._dll_GetDesktops_HMONITOR
      this._dll_SwitchDesktop := this._dll_SwitchDesktop_HMONITOR
      this._dll_CreateDesktop := this._dll_CreateDesktop_HMONITOR
      this._dll_RemoveDesktop := this._dll_RemoveDesktop_normal
      this.idx_GetId := 4
      this.idx_VirtualDesktopNameChanged := 8
      this._dll_VirtualDesktopNameChanged := VD._dll_VirtualDesktopNameChanged_normal
      this.idx_SetDesktopName := 14
      this.idx_GetName := 6
      this.idx_VirtualDesktopWallpaperChanged := -1
      this.idx_SetDesktopWallpaper := -1
      this.idx_GetWallpaper := -1
      this.idx_VirtualDesktopCreated := 3
      this.idx_VirtualDesktopDestroyBegin := 4
      this.idx_VirtualDesktopDestroyFailed := 5
      this.idx_VirtualDesktopDestroyed := 6
      this.idx_ViewVirtualDesktopChanged := 9
      this.idx_CurrentVirtualDesktopChanged := 10
      this._dll_VirtualDesktopCreated := VD._dll_VirtualDesktopCreated_normal
      this._dll_VirtualDesktopDestroyBegin := VD._dll_VirtualDesktopDestroyBegin_normal
      this._dll_VirtualDesktopDestroyFailed := VD._dll_VirtualDesktopDestroyFailed_normal
      this._dll_VirtualDesktopDestroyed := VD._dll_VirtualDesktopDestroyed_normal
      this._dll_ViewVirtualDesktopChanged := VD._dll_ViewVirtualDesktopChanged_normal
      this._dll_CurrentVirtualDesktopChanged := VD._dll_CurrentVirtualDesktopChanged_normal
      this.IVirtualDesktopNotification_methods_count := 11
    } else if (buildNumber < 22483) {
      IID_IVirtualDesktopManagerInternal_str := "{b2f925b9-5a0f-4d2e-9f4d-2b1507593c10}"
      IID_IVirtualDesktop_str := "{536d3495-b208-4cc9-ae26-de8111275bf8}"
      this.IID_IVirtualDesktopNotification_n1 := 5481970284372180562
      this.IID_IVirtualDesktopNotification_n2 := -1679294552252794956
      idx_MoveViewToDesktop := 4
      idx_GetCurrentDesktop := 6
      idx_GetDesktops := 7
      idx_SwitchDesktop := 9
      idx_CreateDesktop := 10
      idx_RemoveDesktop := 12
      this.idx_SwitchDesktopWithAnimation := -1
      this._dll_MoveViewToDesktop := this._dll_MoveViewToDesktop_normal
      this._dll_GetCurrentDesktop := this._dll_GetCurrentDesktop_HMONITOR
      this._dll_GetDesktops := this._dll_GetDesktops_HMONITOR
      this._dll_SwitchDesktop := this._dll_SwitchDesktop_HMONITOR
      this._dll_CreateDesktop := this._dll_CreateDesktop_HMONITOR
      this._dll_RemoveDesktop := this._dll_RemoveDesktop_normal
      this.idx_GetId := 4
      this.idx_VirtualDesktopNameChanged := 9
      this._dll_VirtualDesktopNameChanged := VD._dll_VirtualDesktopNameChanged_normal
      this.idx_SetDesktopName := 15
      this.idx_GetName := 6
      this.idx_VirtualDesktopWallpaperChanged := 12
      this._dll_VirtualDesktopWallpaperChanged := VD._dll_VirtualDesktopWallpaperChanged_normal
      this.idx_SetDesktopWallpaper := 16
      this.idx_GetWallpaper := 7
      this.idx_VirtualDesktopCreated := 3
      this.idx_VirtualDesktopDestroyBegin := 4
      this.idx_VirtualDesktopDestroyFailed := 5
      this.idx_VirtualDesktopDestroyed := 6
      this.idx_ViewVirtualDesktopChanged := 10
      this.idx_CurrentVirtualDesktopChanged := 11
      this._dll_VirtualDesktopCreated := VD._dll_VirtualDesktopCreated_IObjectArray
      this._dll_VirtualDesktopDestroyBegin := VD._dll_VirtualDesktopDestroyBegin_IObjectArray
      this._dll_VirtualDesktopDestroyFailed := VD._dll_VirtualDesktopDestroyFailed_IObjectArray
      this._dll_VirtualDesktopDestroyed := VD._dll_VirtualDesktopDestroyed_IObjectArray
      this._dll_ViewVirtualDesktopChanged := VD._dll_ViewVirtualDesktopChanged_normal
      this._dll_CurrentVirtualDesktopChanged := VD._dll_CurrentVirtualDesktopChanged_IObjectArray
      this.IVirtualDesktopNotification_methods_count := 13
    } else if (buildNumber <= 22621 && revisionNumber < 2215) {
      IID_IVirtualDesktopManagerInternal_str := "{b2f925b9-5a0f-4d2e-9f4d-2b1507593c10}"
      IID_IVirtualDesktop_str := "{536d3495-b208-4cc9-ae26-de8111275bf8}"
      this.IID_IVirtualDesktopNotification_n1 := 5481970284372180562
      this.IID_IVirtualDesktopNotification_n2 := -1679294552252794956
      idx_MoveViewToDesktop := 4
      idx_GetCurrentDesktop := 6
      idx_GetDesktops := 8
      idx_SwitchDesktop := 10
      idx_CreateDesktop := 11
      idx_RemoveDesktop := 13
      this.idx_SwitchDesktopWithAnimation := -1
      this._dll_MoveViewToDesktop := this._dll_MoveViewToDesktop_normal
      this._dll_GetCurrentDesktop := this._dll_GetCurrentDesktop_HMONITOR
      this._dll_GetDesktops := this._dll_GetDesktops_HMONITOR
      this._dll_SwitchDesktop := this._dll_SwitchDesktop_HMONITOR
      this._dll_CreateDesktop := this._dll_CreateDesktop_HMONITOR
      this._dll_RemoveDesktop := this._dll_RemoveDesktop_normal
      this.idx_GetId := 4
      this.idx_VirtualDesktopNameChanged := 9
      this._dll_VirtualDesktopNameChanged := VD._dll_VirtualDesktopNameChanged_normal
      this.idx_SetDesktopName := 16
      this.idx_GetName := 6
      this.idx_VirtualDesktopWallpaperChanged := 12
      this._dll_VirtualDesktopWallpaperChanged := VD._dll_VirtualDesktopWallpaperChanged_normal
      this.idx_SetDesktopWallpaper := 17
      this.idx_GetWallpaper := 7
      this.idx_VirtualDesktopCreated := 3
      this.idx_VirtualDesktopDestroyBegin := 4
      this.idx_VirtualDesktopDestroyFailed := 5
      this.idx_VirtualDesktopDestroyed := 6
      this.idx_ViewVirtualDesktopChanged := 10
      this.idx_CurrentVirtualDesktopChanged := 11
      this._dll_VirtualDesktopCreated := VD._dll_VirtualDesktopCreated_IObjectArray
      this._dll_VirtualDesktopDestroyBegin := VD._dll_VirtualDesktopDestroyBegin_IObjectArray
      this._dll_VirtualDesktopDestroyFailed := VD._dll_VirtualDesktopDestroyFailed_IObjectArray
      this._dll_VirtualDesktopDestroyed := VD._dll_VirtualDesktopDestroyed_IObjectArray
      this._dll_ViewVirtualDesktopChanged := VD._dll_ViewVirtualDesktopChanged_normal
      this._dll_CurrentVirtualDesktopChanged := VD._dll_CurrentVirtualDesktopChanged_IObjectArray
      this.IVirtualDesktopNotification_methods_count := 13
    } else if (buildNumber <= 22631 && revisionNumber < 3085) {
      IID_IVirtualDesktopManagerInternal_str := "{a3175f2d-239c-4bd2-8aa0-eeba8b0b138e}"
      IID_IVirtualDesktop_str := "{3f07f4be-b107-441a-af0f-39d82529072c}"
      this.IID_IVirtualDesktopNotification_n1 := 5123538856297626140
      this.IID_IVirtualDesktopNotification_n2 := 8491238173783613346
      idx_MoveViewToDesktop := 4
      idx_GetCurrentDesktop := 6
      idx_GetDesktops := 7
      idx_SwitchDesktop := 9
      idx_CreateDesktop := 10
      idx_RemoveDesktop := 12
      this.idx_SwitchDesktopWithAnimation := 21
      this._dll_MoveViewToDesktop := this._dll_MoveViewToDesktop_normal
      this._dll_GetCurrentDesktop := this._dll_GetCurrentDesktop_normal
      this._dll_GetDesktops := this._dll_GetDesktops_normal
      this._dll_SwitchDesktop := this._dll_SwitchDesktop_normal
      this._dll_CreateDesktop := this._dll_CreateDesktop_normal
      this._dll_RemoveDesktop := this._dll_RemoveDesktop_normal
      this.idx_GetId := 4
      this.idx_VirtualDesktopNameChanged := 8
      this._dll_VirtualDesktopNameChanged := VD._dll_VirtualDesktopNameChanged_normal
      this.idx_SetDesktopName := 15
      this.idx_GetName := 5
      this.idx_VirtualDesktopWallpaperChanged := 11
      this._dll_VirtualDesktopWallpaperChanged := VD._dll_VirtualDesktopWallpaperChanged_normal
      this.idx_SetDesktopWallpaper := 16
      this.idx_GetWallpaper := 6
      this.idx_VirtualDesktopCreated := 3
      this.idx_VirtualDesktopDestroyBegin := 4
      this.idx_VirtualDesktopDestroyFailed := 5
      this.idx_VirtualDesktopDestroyed := 6
      this.idx_ViewVirtualDesktopChanged := 9
      this.idx_CurrentVirtualDesktopChanged := 10
      this._dll_VirtualDesktopCreated := VD._dll_VirtualDesktopCreated_normal
      this._dll_VirtualDesktopDestroyBegin := VD._dll_VirtualDesktopDestroyBegin_normal
      this._dll_VirtualDesktopDestroyFailed := VD._dll_VirtualDesktopDestroyFailed_normal
      this._dll_VirtualDesktopDestroyed := VD._dll_VirtualDesktopDestroyed_normal
      this._dll_ViewVirtualDesktopChanged := VD._dll_ViewVirtualDesktopChanged_normal
      this._dll_CurrentVirtualDesktopChanged := VD._dll_CurrentVirtualDesktopChanged_normal
      this.IVirtualDesktopNotification_methods_count := 14
    } else if (buildNumber < 26100) {
      IID_IVirtualDesktopManagerInternal_str := "{53f5ca0b-158f-4124-900c-057158060b27}"
      IID_IVirtualDesktop_str := "{3f07f4be-b107-441a-af0f-39d82529072c}"
      this.IID_IVirtualDesktopNotification_n1 := 5308375338100058445
      this.IID_IVirtualDesktopNotification_n2 := -2401892766147978065
      idx_MoveViewToDesktop := 4
      idx_GetCurrentDesktop := 6
      idx_GetDesktops := 7
      idx_SwitchDesktop := 9
      idx_CreateDesktop := 10
      idx_RemoveDesktop := 12
      this.idx_SwitchDesktopWithAnimation := 21
      this._dll_MoveViewToDesktop := this._dll_MoveViewToDesktop_normal
      this._dll_GetCurrentDesktop := this._dll_GetCurrentDesktop_normal
      this._dll_GetDesktops := this._dll_GetDesktops_normal
      this._dll_SwitchDesktop := this._dll_SwitchDesktop_normal
      this._dll_CreateDesktop := this._dll_CreateDesktop_normal
      this._dll_RemoveDesktop := this._dll_RemoveDesktop_normal
      this.idx_GetId := 4
      this.idx_VirtualDesktopNameChanged := 8
      this._dll_VirtualDesktopNameChanged := VD._dll_VirtualDesktopNameChanged_normal
      this.idx_SetDesktopName := 15
      this.idx_GetName := 5
      this.idx_VirtualDesktopWallpaperChanged := 11
      this._dll_VirtualDesktopWallpaperChanged := VD._dll_VirtualDesktopWallpaperChanged_normal
      this.idx_SetDesktopWallpaper := 16
      this.idx_GetWallpaper := 6
      this.idx_VirtualDesktopCreated := 3
      this.idx_VirtualDesktopDestroyBegin := 4
      this.idx_VirtualDesktopDestroyFailed := 5
      this.idx_VirtualDesktopDestroyed := 6
      this.idx_ViewVirtualDesktopChanged := 9
      this.idx_CurrentVirtualDesktopChanged := 10
      this._dll_VirtualDesktopCreated := VD._dll_VirtualDesktopCreated_normal
      this._dll_VirtualDesktopDestroyBegin := VD._dll_VirtualDesktopDestroyBegin_normal
      this._dll_VirtualDesktopDestroyFailed := VD._dll_VirtualDesktopDestroyFailed_normal
      this._dll_VirtualDesktopDestroyed := VD._dll_VirtualDesktopDestroyed_normal
      this._dll_ViewVirtualDesktopChanged := VD._dll_ViewVirtualDesktopChanged_normal
      this._dll_CurrentVirtualDesktopChanged := VD._dll_CurrentVirtualDesktopChanged_normal
      this.IVirtualDesktopNotification_methods_count := 14
    } else {
      IID_IVirtualDesktopManagerInternal_str := "{53f5ca0b-158f-4124-900c-057158060b27}"
      IID_IVirtualDesktop_str := "{3f07f4be-b107-441a-af0f-39d82529072c}"
      this.IID_IVirtualDesktopNotification_n1 := 5308375338100058445
      this.IID_IVirtualDesktopNotification_n2 := -2401892766147978065
      idx_MoveViewToDesktop := 4
      idx_GetCurrentDesktop := 6
      idx_GetDesktops := 7
      idx_SwitchDesktop := 9
      idx_CreateDesktop := 11
      idx_RemoveDesktop := 13
      this.idx_SwitchDesktopWithAnimation := 22
      this._dll_MoveViewToDesktop := this._dll_MoveViewToDesktop_normal
      this._dll_GetCurrentDesktop := this._dll_GetCurrentDesktop_normal
      this._dll_GetDesktops := this._dll_GetDesktops_normal
      this._dll_SwitchDesktop := this._dll_SwitchDesktop_normal
      this._dll_CreateDesktop := this._dll_CreateDesktop_normal
      this._dll_RemoveDesktop := this._dll_RemoveDesktop_normal
      this.idx_GetId := 4
      this.idx_VirtualDesktopNameChanged := 8
      this._dll_VirtualDesktopNameChanged := VD._dll_VirtualDesktopNameChanged_normal
      this.idx_SetDesktopName := 16
      this.idx_GetName := 5
      this.idx_VirtualDesktopWallpaperChanged := 11
      this._dll_VirtualDesktopWallpaperChanged := VD._dll_VirtualDesktopWallpaperChanged_normal
      this.idx_SetDesktopWallpaper := 17
      this.idx_GetWallpaper := 6
      this.idx_VirtualDesktopCreated := 3
      this.idx_VirtualDesktopDestroyBegin := 4
      this.idx_VirtualDesktopDestroyFailed := 5
      this.idx_VirtualDesktopDestroyed := 6
      this.idx_ViewVirtualDesktopChanged := 9
      this.idx_CurrentVirtualDesktopChanged := 10
      this._dll_VirtualDesktopCreated := VD._dll_VirtualDesktopCreated_normal
      this._dll_VirtualDesktopDestroyBegin := VD._dll_VirtualDesktopDestroyBegin_normal
      this._dll_VirtualDesktopDestroyFailed := VD._dll_VirtualDesktopDestroyFailed_normal
      this._dll_VirtualDesktopDestroyed := VD._dll_VirtualDesktopDestroyed_normal
      this._dll_ViewVirtualDesktopChanged := VD._dll_ViewVirtualDesktopChanged_normal
      this._dll_CurrentVirtualDesktopChanged := VD._dll_CurrentVirtualDesktopChanged_normal
      this.IVirtualDesktopNotification_methods_count := 14
    }
    this.IVirtualDesktopManager := ComObject("{aa509086-5ca9-4c25-8f95-589d3c07b48a}", "{a5cd92ff-29be-454c-8d04-d82879fb3f1b}")
    this.ptr_MoveWindowToDesktop := this._vtable(this.IVirtualDesktopManager.Ptr, 5)
    this.CImmersiveShell_IServiceProvider := ComObject("{c2f03a33-21f5-47fa-b4bb-156362a2f239}", "{6d5140c1-7436-11ce-8034-00aa006009fa}")
    this.IVirtualDesktopManagerInternal := ComObjQuery(this.CImmersiveShell_IServiceProvider.Ptr, "{c5e0cdca-7b6e-41b2-9fc4-d93975cc467b}", IID_IVirtualDesktopManagerInternal_str)
    this.IVirtualDesktopPinnedApps := ComObjQuery(this.CImmersiveShell_IServiceProvider.Ptr, "{b5a399e7-1c87-46b8-88e9-fc5747b171bd}", "{4ce81583-1e4c-4632-a621-07a53543148f}")
    this.IApplicationViewCollection := ComObjQuery(this.CImmersiveShell_IServiceProvider.Ptr, "{1841c6d7-4f9d-42c0-af41-8747538f10e5}", "{1841c6d7-4f9d-42c0-af41-8747538f10e5}")
    this.ptr_MoveViewToDesktop := this._vtable(this.IVirtualDesktopManagerInternal.Ptr, idx_MoveViewToDesktop)
    this.ptr_GetCurrentDesktop := this._vtable(this.IVirtualDesktopManagerInternal.Ptr, idx_GetCurrentDesktop)
    this.ptr_GetDesktops := this._vtable(this.IVirtualDesktopManagerInternal.Ptr, idx_GetDesktops)
    this.ptr_SwitchDesktop := this._vtable(this.IVirtualDesktopManagerInternal.Ptr, idx_SwitchDesktop)
    this.ptr_CreateDesktop := this._vtable(this.IVirtualDesktopManagerInternal.Ptr, idx_CreateDesktop)
    this.ptr_RemoveDesktop := this._vtable(this.IVirtualDesktopManagerInternal.Ptr, idx_RemoveDesktop)
    if (this.idx_SwitchDesktopWithAnimation > -1) {
      this.ptr_SwitchDesktopWithAnimation := this._vtable(this.IVirtualDesktopManagerInternal.Ptr, this.idx_SwitchDesktopWithAnimation)
    }
    if (this.idx_SetDesktopName > -1) {
      this.ptr_SetDesktopName := this._vtable(this.IVirtualDesktopManagerInternal.Ptr, this.idx_SetDesktopName)
    }
    if (this.idx_SetDesktopWallpaper > -1) {
      this.ptr_SetDesktopWallpaper := this._vtable(this.IVirtualDesktopManagerInternal.Ptr, this.idx_SetDesktopWallpaper)
    }
    this.ptr_IsAppIdPinned := this._vtable(this.IVirtualDesktopPinnedApps.Ptr, 3)
    this.ptr_PinAppID := this._vtable(this.IVirtualDesktopPinnedApps.Ptr, 4)
    this.ptr_UnpinAppID := this._vtable(this.IVirtualDesktopPinnedApps.Ptr, 5)
    this.ptr_IsViewPinned := this._vtable(this.IVirtualDesktopPinnedApps.Ptr, 6)
    this.ptr_PinView := this._vtable(this.IVirtualDesktopPinnedApps.Ptr, 7)
    this.ptr_UnpinView := this._vtable(this.IVirtualDesktopPinnedApps.Ptr, 8)
    this.ptr_GetViewForHwnd := this._vtable(this.IApplicationViewCollection.Ptr, 6)
    this.IID_IVirtualDesktop_ptr := DllCall("GlobalAlloc", "UInt", 0x00, "Uint", 16, "Ptr")
    DllCall("ole32\CLSIDFromString", "Str", IID_IVirtualDesktop_str, "Ptr", this.IID_IVirtualDesktop_ptr)
    OnMessage(DllCall("RegisterWindowMessageW", "WStr", "TaskbarCreated", "Uint"), VD._ExplorerRestarted)
  }
  static _ExplorerRestarted(lParam, msg, hwnd) {
    VD._init()
  }
  static _dll_MoveViewToDesktop_normal(IApplicationView, IVirtualDesktop) {
    DllCall(this.ptr_MoveViewToDesktop, "Ptr", this.IVirtualDesktopManagerInternal.Ptr, "Ptr", IApplicationView, "Ptr", IVirtualDesktop)
  }
  static _dll_GetCurrentDesktop_normal() {
    DllCall(this.ptr_GetCurrentDesktop, "Ptr", this.IVirtualDesktopManagerInternal.Ptr, "Ptr*", &IVirtualDesktop_currentDesktop := 0)
    return IVirtualDesktop_currentDesktop
  }
  static _dll_GetDesktops_normal() {
    DllCall(this.ptr_GetDesktops, "Ptr", this.IVirtualDesktopManagerInternal.Ptr, "Ptr*", &IObjectArray := 0)
    return IObjectArray
  }
  static _dll_SwitchDesktop_normal(IVirtualDesktop) {
    DllCall(this.ptr_SwitchDesktop, "Ptr", this.IVirtualDesktopManagerInternal.Ptr, "Ptr", IVirtualDesktop)
  }
  static _dll_CreateDesktop_normal() {
    DllCall(this.ptr_CreateDesktop, "Ptr", this.IVirtualDesktopManagerInternal.Ptr, "Ptr*", &IVirtualDesktop_created := 0)
    return IVirtualDesktop_created
  }
  static _dll_RemoveDesktop_normal(IVirtualDesktop, IVirtualDesktop_fallback) {
    DllCall(this.ptr_RemoveDesktop, "Ptr", this.IVirtualDesktopManagerInternal.Ptr, "Ptr", IVirtualDesktop, "Ptr", IVirtualDesktop_fallback)
  }
  static _dll_GetCurrentDesktop_HMONITOR() {
    DllCall(this.ptr_GetCurrentDesktop, "Ptr", this.IVirtualDesktopManagerInternal.Ptr, "Ptr", 0, "Ptr*", &IVirtualDesktop_currentDesktop := 0)
    return IVirtualDesktop_currentDesktop
  }
  static _dll_GetDesktops_HMONITOR() {
    DllCall(this.ptr_GetDesktops, "Ptr", this.IVirtualDesktopManagerInternal.Ptr, "Ptr", 0, "Ptr*", &IObjectArray := 0)
    return IObjectArray
  }
  static _dll_SwitchDesktop_HMONITOR(IVirtualDesktop) {
    DllCall(this.ptr_SwitchDesktop, "Ptr", this.IVirtualDesktopManagerInternal.Ptr, "Ptr", 0, "Ptr", IVirtualDesktop)
  }
  static _dll_CreateDesktop_HMONITOR() {
    DllCall(this.ptr_CreateDesktop, "Ptr", this.IVirtualDesktopManagerInternal.Ptr, "Ptr", 0, "Ptr*", &IVirtualDesktop_created := 0)
    return IVirtualDesktop_created
  }
  static _dll_IsAppIdPinned(exe_path) {
    DllCall(this.ptr_IsAppIdPinned, "Ptr", this.IVirtualDesktopPinnedApps.Ptr, "WStr", exe_path, "Int*", &appIsPinned := 0)
    return appIsPinned
  }
  static _dll_PinAppID(exe_path) {
    DllCall(this.ptr_PinAppID, "Ptr", this.IVirtualDesktopPinnedApps.Ptr, "WStr", exe_path)
  }
  static _dll_UnpinAppID(exe_path) {
    DllCall(this.ptr_UnpinAppID, "Ptr", this.IVirtualDesktopPinnedApps.Ptr, "WStr", exe_path)
  }
  static _dll_IsViewPinned(IApplicationView) {
    DllCall(this.ptr_IsViewPinned, "Ptr", this.IVirtualDesktopPinnedApps.Ptr, "Ptr", IApplicationView, "Int*", &viewIsPinned := 0)
    return viewIsPinned
  }
  static _dll_PinView(IApplicationView) {
    DllCall(this.ptr_PinView, "Ptr", this.IVirtualDesktopPinnedApps.Ptr, "Ptr", IApplicationView)
  }
  static _dll_UnpinView(IApplicationView) {
    DllCall(this.ptr_UnpinView, "Ptr", this.IVirtualDesktopPinnedApps.Ptr, "Ptr", IApplicationView)
  }
  static _dll_GetViewForHwnd(HWND) {
    DllCall(this.ptr_GetViewForHwnd, "Ptr", this.IApplicationViewCollection.Ptr, "Ptr", HWND, "Ptr*", &IApplicationView := 0)
    return IApplicationView
  }
  static getCount() {
    return this._GetDesktops_Obj().GetCount()
  }
  static goToDesktopNum(desktopNum, shouldActivateUponArrival := true) {
    if (shouldActivateUponArrival && this._shouldActivateUponArrival()) {
      firstWindowId := this._getFirstWindowInVD(desktopNum)
      if (!firstWindowId) {
        firstWindowId := WinExist("ahk_class Progman ahk_exe explorer.exe")
      }
    } else {
      firstWindowId := 0
    }
    IVirtualDesktop := this._GetDesktops_Obj().GetAt(desktopNum)
    if (VD.animation_on && firstWindowId) {
      if (this.idx_SwitchDesktopWithAnimation > -1) {
        DllCall("SetForegroundWindow", "Ptr", firstWindowId)
        DllCall(this.ptr_SwitchDesktopWithAnimation, "Ptr", this.IVirtualDesktopManagerInternal.Ptr, "Ptr", IVirtualDesktop)
        this._waitForCurrentDesktopArrived(desktopNum, firstWindowId)
      } else {
        VD_animation_gui := Gui("-Border -SysMenu +Owner -Caption")
        VD_animation_gui_hwnd := VD_animation_gui.Hwnd + 0
        GUID_Desktop := Buffer(16)
        ptr_GetId := this._vtable(IVirtualDesktop, this.idx_GetId)
        DllCall(ptr_GetId, "Ptr", IVirtualDesktop, "Ptr", GUID_Desktop)
        DllCall(this.ptr_MoveWindowToDesktop, "Ptr", this.IVirtualDesktopManager.Ptr, "Ptr", VD_animation_gui_hwnd, "Ptr", GUID_Desktop)
        DllCall("ShowWindow", "Ptr", VD_animation_gui_hwnd, "Int", 4)
        this.SetForegroundWindow(VD_animation_gui_hwnd)
        this._waitForCurrentDesktopArrived(desktopNum, firstWindowId)
        VD_animation_gui.Destroy()
      }
    } else {
      this._dll_SwitchDesktop(IVirtualDesktop)
      DllCall("SetForegroundWindow", "Ptr", firstWindowId)
      this._waitForCurrentDesktopArrived(desktopNum, firstWindowId)
    }
  }
  static _shouldActivateUponArrival() {
    if (WinActive(this._getLocalizedWord_TaskView() " ahk_exe explorer.exe")) {
      return false
    }
    return true
  }
  static _waitForCurrentDesktopArrived(desktopNum, firstWindowId) {
    loop 20 {
      if (this.getCurrentDesktopNum() == desktopNum) {
        DllCall("SetForegroundWindow", "Ptr", firstWindowId)
        break
      }
      Sleep 25
    }
  }
  static _onceLocalizedWord_TaskView() {
    hModule := (hModule := DllCall("GetModuleHandle", "Str", "twinui.pcshell.dll", "Ptr")) ? hModule : DllCall("LoadLibrary", "Str", "twinui.pcshell.dll", "Ptr")
    length := DllCall("LoadString", "Uint", hModule, "Uint", 1512, "Ptr*", &lpBuffer := 0, "Int", 0)
    savedLocalizedWord_TaskView := StrGet(lpBuffer, length, "UTF-16")
    return savedLocalizedWord_TaskView
  }
  static _onceLocalizedWord_Desktop() {
    hModule := DllCall("GetModuleHandle", "Str", "shell32.dll", "Ptr")
    length := DllCall("LoadString", "Uint", hModule, "Uint", 21769, "Ptr*", &lpBuffer := 0, "Int", 0)
    savedLocalizedWord_Desktop := StrGet(lpBuffer, length, "UTF-16")
    return savedLocalizedWord_Desktop
  }
  static _getLocalizedWord_TaskView() {
    static savedLocalizedWord_TaskView := VD._onceLocalizedWord_TaskView()
    return savedLocalizedWord_TaskView
  }
  static _getLocalizedWord_Desktop() {
    static savedLocalizedWord_Desktop := VD._onceLocalizedWord_Desktop()
    return savedLocalizedWord_Desktop
  }
  static getNameFromDesktopNum(desktopNum) {
    desktopName := ""
    if (desktopNum > 0 && this.idx_GetName > -1) {
      IVirtualDesktop := this._GetDesktops_Obj().GetAt(desktopNum)
      if (IVirtualDesktop) {
        ptr_GetName := this._vtable(IVirtualDesktop, this.idx_GetName)
        DllCall(ptr_GetName, "Ptr", IVirtualDesktop, "Ptr*", &HSTRING := 0)
        desktopName := StrGet(DllCall("combase\WindowsGetStringRawBuffer", "Ptr", HSTRING, "Uint*", &length := 0, "Ptr"), "UTF-16")
        DllCall("combase\WindowsDeleteString", "Ptr", HSTRING)
      }
    }
    if (!desktopName) {
      desktopName := this._getLocalizedWord_Desktop() " " desktopNum
    }
    return desktopName
  }
  static setNameToDesktopNum(desktopName, desktopNum) {
    if (desktopNum > 0 && this.idx_SetDesktopName > -1) {
      IVirtualDesktop := this._GetDesktops_Obj().GetAt(desktopNum)
      if (IVirtualDesktop) {
        DllCall("combase\WindowsCreateString", "WStr", desktopName, "Uint", StrLen(desktopName), "Ptr*", &HSTRING := 0)
        DllCall(this.ptr_SetDesktopName, "Ptr", this.IVirtualDesktopManagerInternal.Ptr, "Ptr", IVirtualDesktop, "Ptr", HSTRING)
        DllCall("combase\WindowsDeleteString", "Ptr", HSTRING)
      }
    }
  }
  static getWallpaperFromDesktopNum(desktopNum) {
    desktopWallpaper := ""
    if (desktopNum > 0 && this.idx_GetWallpaper > -1) {
      IVirtualDesktop := this._GetDesktops_Obj().GetAt(desktopNum)
      if (IVirtualDesktop) {
        ptr_GetWallpaper := this._vtable(IVirtualDesktop, this.idx_GetWallpaper)
        DllCall(ptr_GetWallpaper, "Ptr", IVirtualDesktop, "Ptr*", &HSTRING := 0)
        desktopWallpaper := StrGet(DllCall("combase\WindowsGetStringRawBuffer", "Ptr", HSTRING, "Uint*", &length := 0, "Ptr"), "UTF-16")
        DllCall("combase\WindowsDeleteString", "Ptr", HSTRING)
      }
    }
    return desktopWallpaper
  }
  static setWallpaperToDesktopNum(desktopWallpaper, desktopNum) {
    if (desktopNum > 0 && this.idx_SetDesktopWallpaper > -1) {
      IVirtualDesktop := this._GetDesktops_Obj().GetAt(desktopNum)
      if (IVirtualDesktop) {
        DllCall("combase\WindowsCreateString", "WStr", desktopWallpaper, "Uint", StrLen(desktopWallpaper), "Ptr*", &HSTRING := 0)
        DllCall(this.ptr_SetDesktopWallpaper, "Ptr", this.IVirtualDesktopManagerInternal.Ptr, "Ptr", IVirtualDesktop, "Ptr", HSTRING)
        DllCall("combase\WindowsDeleteString", "Ptr", HSTRING)
      }
    }
  }
  static getDesktopNumOfWindow(wintitle) {
    found := this._tryGetValidWindow(wintitle)
    if (!found) {
      return -1
    }
    thePView := found[2]
    desktopNum_ofWindow := this._desktopNum_from_pView(thePView)
    return desktopNum_ofWindow
  }
  static goToDesktopOfWindow(wintitle, activateYourWindow := true) {
    found := this._tryGetValidWindow(wintitle)
    if (!found) {
      return -1
    }
    theHwnd := found[1]
    thePView := found[2]
    desktopNum_ofWindow := this._desktopNum_from_pView(thePView)
    this.goToDesktopNum(desktopNum_ofWindow, !activateYourWindow)
    if (activateYourWindow) {
      WinActivate "ahk_id " theHwnd
    }
  }
  static MoveWindowToDesktopNum(wintitle, desktopNum) {
    found := this._tryGetValidWindow(wintitle)
    if (!found) {
      return this.WindowInfo(0, desktopNum)
    }
    theHwnd := found[1]
    thePView := found[2]
    needActivateWindowUnder := false
    if (activeHwnd := WinExist("A")) {
      if (activeHwnd == theHwnd) {
        currentDesktopNum := this.getCurrentDesktopNum()
        if (!(currentDesktopNum == desktopNum)) {
          needActivateWindowUnder := true
        }
      }
    }
    IVirtualDesktop := this._GetDesktops_Obj().GetAt(desktopNum)
    this._dll_MoveViewToDesktop(thePView, IVirtualDesktop)
    if (needActivateWindowUnder) {
      firstWindowId := this._getFirstWindowInVD(currentDesktopNum, theHwnd)
      if (!firstWindowId) {
        firstWindowId := WinExist("ahk_class Progman ahk_exe explorer.exe")
      }
      this.SetForegroundWindow(firstWindowId)
    }
    return this.WindowInfo(theHwnd, desktopNum)
  }
  class WindowInfo {
    __New(hwnd, desktopNum) {
      this.hwnd := hwnd
      this.desktopNum := desktopNum
    }
    follow() {
      VD.goToDesktopNum(this.desktopNum, false)
      if (WinExist("ahk_id " this.hwnd)) {
        WinActivate
      }
    }
  }
  static getRelativeDesktopNum(anchor_desktopNum, relative_count) {
    Desktops_Obj := this._GetDesktops_Obj()
    count_Desktops := Desktops_Obj.GetCount()
    absolute_desktopNum := anchor_desktopNum + relative_count
    absolute_desktopNum := Mod(absolute_desktopNum, count_Desktops)
    if (absolute_desktopNum <= 0) {
      absolute_desktopNum := absolute_desktopNum + count_Desktops
    }
    return absolute_desktopNum
  }
  static MoveWindowToRelativeDesktopNum(wintitle, relative_count) {
    desktopNum_ofWindow := this.getDesktopNumOfWindow(wintitle)
    absolute_desktopNum := this.getRelativeDesktopNum(desktopNum_ofWindow, relative_count)
    return this.MoveWindowToDesktopNum(wintitle, absolute_desktopNum)
  }
  static gotoRelativeDesktopNum(relative_count) {
    this.goToDesktopNum(this.getRelativeDesktopNum(this.getCurrentDesktopNum(), relative_count))
  }
  static MoveWindowToCurrentDesktop(wintitle, activateYourWindow := true) {
    found := this._tryGetValidWindow(wintitle)
    if (!found) {
      return -1
    }
    theHwnd := found[1]
    thePView := found[2]
    currentDesktopNum := this.getCurrentDesktopNum()
    IVirtualDesktop := this._GetDesktops_Obj().GetAt(currentDesktopNum)
    this._dll_MoveViewToDesktop(thePView, IVirtualDesktop)
    if (activateYourWindow) {
      WinActivate "ahk_id " theHwnd
    }
  }
  static getCurrentDesktopNum() {
    IVirtualDesktop_ofCurrentDesktop := this._dll_GetCurrentDesktop()
    desktopNum := this._desktopNum_from_IVirtualDesktop(IVirtualDesktop_ofCurrentDesktop)
    return desktopNum
  }
  static createDesktop(goThere := false) {
    IVirtualDesktop_ofNewDesktop := this._dll_CreateDesktop()
    if (goThere) {
      desktopNum := this._desktopNum_from_IVirtualDesktop(IVirtualDesktop_ofNewDesktop)
      this.goToDesktopNum(desktopNum)
    }
  }
  static createUntil(howMany, goToLastlyCreated := false) {
    howManyThereAlreadyAre := this.getCount()
    if (howManyThereAlreadyAre >= howMany) {
      return
    }
    loop howMany - howManyThereAlreadyAre - 1 {
      this.createDesktop(false)
    }
    this.createDesktop(goToLastlyCreated)
  }
  static removeDesktop(desktopNum, fallback_desktopNum := false) {
    Desktops_Obj := this._GetDesktops_Obj()
    if (!fallback_desktopNum) {
      if (desktopNum > 1) {
        fallback_desktopNum := desktopNum - 1
      } else if (desktopNum < Desktops_Obj.GetCount()) {
        fallback_desktopNum := desktopNum + 1
      } else {
        return false
      }
    }
    IVirtualDesktop := Desktops_Obj.GetAt(desktopNum)
    IVirtualDesktop_fallback := Desktops_Obj.GetAt(fallback_desktopNum)
    this._dll_RemoveDesktop(IVirtualDesktop, IVirtualDesktop_fallback)
  }
  static IsExePinned(exe_path) {
    appIsPinned := this._dll_IsAppIdPinned(exe_path)
    return appIsPinned
  }
  static TogglePinExe(exe_path) {
    appIsPinned := this._dll_IsAppIdPinned(exe_path)
    if (appIsPinned) {
      this._dll_UnpinAppID(exe_path)
    } else {
      this._dll_PinAppID(exe_path)
    }
  }
  static PinExe(exe_path) {
    this._dll_PinAppID(exe_path)
  }
  static UnPinExe(exe_path) {
    this._dll_UnpinAppID(exe_path)
  }
  static IsWindowPinned(wintitle) {
    found := this._tryGetValidWindow(wintitle)
    if (!found) {
      return -1
    }
    thePView := found[2]
    viewIsPinned := this._dll_IsViewPinned(thePView)
    return viewIsPinned
  }
  static TogglePinWindow(wintitle) {
    found := this._tryGetValidWindow(wintitle)
    if (!found) {
      return -1
    }
    thePView := found[2]
    viewIsPinned := this._dll_IsViewPinned(thePView)
    if (viewIsPinned) {
      this._dll_UnPinView(thePView)
    } else {
      this._dll_PinView(thePView)
    }
  }
  static PinWindow(wintitle) {
    found := this._tryGetValidWindow(wintitle)
    if (!found) {
      return -1
    }
    thePView := found[2]
    this._dll_PinView(thePView)
  }
  static UnPinWindow(wintitle) {
    found := this._tryGetValidWindow(wintitle)
    if (!found) {
      return -1
    }
    thePView := found[2]
    this._dll_UnPinView(thePView)
  }
  static _dll_QueryInterface_Same(riid, ppvObject) {
    if (!ppvObject) {
      return 0x80070057
    }
    if ((NumGet(riid + 0, 0x0, "Int64") == 0 && NumGet(riid + 0, 0x8, "Int64") == 5044031582654955712) || (NumGet(riid + 0, 0x0, "Int64") == VD.IID_IVirtualDesktopNotification_n1 && NumGet(riid + 0, 0x8, "Int64") == VD.IID_IVirtualDesktopNotification_n2)) {
      NumPut("Ptr", this, ppvObject + 0, 0)
      VD._dll_AddRef_Same.Call(this)
      return 0
    }
    NumPut("Ptr", 0, ppvObject + 0, "Ptr")
    return 0x80004002
  }
  static _dll_AddRef_Same() {
    refCount := NumGet(this + 0, A_PtrSize, "UInt")
    refCount++
    NumPut("UInt", refCount, this + 0, A_PtrSize)
    return refCount
  }
  static _dll_Release_Same() {
    refCount := NumGet(this + 0, A_PtrSize, "UInt")
    refCount--
    NumPut("UInt", refCount, this + 0, A_PtrSize)
    return refCount
  }
  static _dll_VirtualDesktopCreated_normal(IVirtualDesktop_created) {
    VD.VirtualDesktopCreated.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_created))
    return 0
  }
  static _dll_VirtualDesktopDestroyBegin_normal(IVirtualDesktop_destroyed, IVirtualDesktop_fallback) {
    VD.VirtualDesktopDestroyBegin.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_destroyed), VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_fallback))
    return 0
  }
  static _dll_VirtualDesktopDestroyFailed_normal(IVirtualDesktop_destroyed, IVirtualDesktop_fallback) {
    VD.VirtualDesktopDestroyFailed.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_destroyed), VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_fallback))
    return 0
  }
  static _dll_VirtualDesktopDestroyed_normal(IVirtualDesktop_destroyed, IVirtualDesktop_fallback) {
    VD.VirtualDesktopDestroyed.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_destroyed), VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_fallback))
    return 0
  }
  static _dll_ViewVirtualDesktopChanged_normal(IApplicationView) {
    VD.ViewVirtualDesktopChanged.Call(IApplicationView)
    return 0
  }
  static _dll_CurrentVirtualDesktopChanged_normal(IVirtualDesktop_old, IVirtualDesktop_new) {
    VD.CurrentVirtualDesktopChanged.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_old), VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_new))
    return 0
  }
  static _dll_VirtualDesktopNameChanged_normal(IVirtualDesktop, HSTRING) {
    desktopName := StrGet(DllCall("combase\WindowsGetStringRawBuffer", "Ptr", HSTRING, "Uint*", &length := 0, "Ptr"), "UTF-16")
    DllCall("combase\WindowsDeleteString", "Ptr", HSTRING)
    VD.VirtualDesktopNameChanged.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop), desktopName)
    return 0
  }
  static _dll_VirtualDesktopWallpaperChanged_normal(IVirtualDesktop, HSTRING) {
    desktopName := StrGet(DllCall("combase\WindowsGetStringRawBuffer", "Ptr", HSTRING, "Uint*", &length := 0, "Ptr"), "UTF-16")
    DllCall("combase\WindowsDeleteString", "Ptr", HSTRING)
    VD.VirtualDesktopWallpaperChanged.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop), desktopName)
    return 0
  }
  static _dll_VirtualDesktopCreated_IObjectArray(IObjectArray, IVirtualDesktop_created) {
    VD.VirtualDesktopCreated.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_created))
    return 0
  }
  static _dll_VirtualDesktopDestroyBegin_IObjectArray(IObjectArray, IVirtualDesktop_destroyed, IVirtualDesktop_fallback) {
    VD.VirtualDesktopDestroyBegin.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_destroyed), VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_fallback))
    return 0
  }
  static _dll_VirtualDesktopDestroyFailed_IObjectArray(IObjectArray, IVirtualDesktop_destroyed, IVirtualDesktop_fallback) {
    VD.VirtualDesktopDestroyFailed.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_destroyed), VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_fallback))
    return 0
  }
  static _dll_VirtualDesktopDestroyed_IObjectArray(IObjectArray, IVirtualDesktop_destroyed, IVirtualDesktop_fallback) {
    VD.VirtualDesktopDestroyed.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_destroyed), VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_fallback))
    return 0
  }
  static _dll_CurrentVirtualDesktopChanged_IObjectArray(IObjectArray, IVirtualDesktop_old, IVirtualDesktop_new) {
    VD.CurrentVirtualDesktopChanged.Call(VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_old), VD._desktopNum_from_IVirtualDesktop(IVirtualDesktop_new))
    return 0
  }
  static RegisterDesktopNotifications() {
    methods_ptr := DllCall("GlobalAlloc", "Uint", 0x40, "Uint", this.IVirtualDesktopNotification_methods_count * A_PtrSize)
    NumPut("Ptr", CallbackCreate(this._dll_QueryInterface_Same, "F"), methods_ptr + 0 * A_PtrSize, "Ptr")
    NumPut("Ptr", CallbackCreate(this._dll_AddRef_Same, "F"), methods_ptr + 1 * A_PtrSize, "Ptr")
    NumPut("Ptr", CallbackCreate(this._dll_Release_Same, "F"), methods_ptr + 2 * A_PtrSize, "Ptr")
    NumPut("Ptr", CallbackCreate(this._dll_VirtualDesktopCreated, "F"), methods_ptr + this.idx_VirtualDesktopCreated * A_PtrSize, "Ptr")
    NumPut("Ptr", CallbackCreate(this._dll_VirtualDesktopDestroyBegin, "F"), methods_ptr + this.idx_VirtualDesktopDestroyBegin * A_PtrSize, "Ptr")
    NumPut("Ptr", CallbackCreate(this._dll_VirtualDesktopDestroyFailed, "F"), methods_ptr + this.idx_VirtualDesktopDestroyFailed * A_PtrSize, "Ptr")
    NumPut("Ptr", CallbackCreate(this._dll_VirtualDesktopDestroyed, "F"), methods_ptr + this.idx_VirtualDesktopDestroyed * A_PtrSize, "Ptr")
    NumPut("Ptr", CallbackCreate(this._dll_ViewVirtualDesktopChanged, "F"), methods_ptr + this.idx_ViewVirtualDesktopChanged * A_PtrSize, "Ptr")
    NumPut("Ptr", CallbackCreate(this._dll_CurrentVirtualDesktopChanged, "F"), methods_ptr + this.idx_CurrentVirtualDesktopChanged * A_PtrSize, "Ptr")
    if (this.idx_VirtualDesktopNameChanged > -1) {
      NumPut("Ptr", CallbackCreate(this._dll_VirtualDesktopNameChanged, "F"), methods_ptr + this.idx_VirtualDesktopNameChanged * A_PtrSize, "Ptr")
    }
    if (this.idx_VirtualDesktopWallpaperChanged > -1) {
      NumPut("Ptr", CallbackCreate(this._dll_VirtualDesktopWallpaperChanged, "F"), methods_ptr + this.idx_VirtualDesktopWallpaperChanged * A_PtrSize, "Ptr")
    }
    ptr := methods_ptr
    end := methods_ptr + this.IVirtualDesktopNotification_methods_count * A_PtrSize
    callback_noop := 0
    while (ptr < end) {
      if (!NumGet(ptr + 0, "Ptr")) {
        if (!callback_noop) {
          callback_noop := CallbackCreate(VD._No_Op, "F")
        }
        NumPut("Ptr", callback_noop, ptr + 0)
      }
      ptr += A_PtrSize
    }
    this.RegisterDesktopNotifications_Same(methods_ptr)
  }
  static _No_Op() {
  }
  static RegisterDesktopNotifications_Same(methods_ptr) {
    obj := DllCall("GlobalAlloc", "Uint", 0x00, "Uint", A_PtrSize + 4)
    NumPut("Ptr", methods_ptr, obj + 0, 0)
    NumPut("UInt", 0, obj + 0, A_PtrSize)
    IDesktopNotificationService := ComObjQuery(this.CImmersiveShell_IServiceProvider.Ptr, "{A501FDEC-4A09-464C-AE4E-1B9C21B84918}", "{0CD45E71-D927-4F15-8B0A-8FEF525337BF}")
    ptr_Register := this._vtable(IDesktopNotificationService.Ptr, 3)
    DllCall(ptr_Register, "Ptr", IDesktopNotificationService.Ptr, "Ptr", obj, "Uint*", &pdwCookie := 0)
  }
  static VirtualDesktopCreated(desktopNum := 0) {
  }
  static VirtualDesktopDestroyBegin(desktopNum_Destroyed := 0, desktopNum_Fallback := 0) {
  }
  static VirtualDesktopDestroyFailed(desktopNum_Destroyed := 0, desktopNum_Fallback := 0) {
  }
  static VirtualDesktopDestroyed(desktopNum_Destroyed := 0, desktopNum_Fallback := 0) {
  }
  static ViewVirtualDesktopChanged(pView := 0) {
  }
  static CurrentVirtualDesktopChanged(desktopNum_Old := 0, desktopNum_New := 0) {
  }
  static VirtualDesktopNameChanged(desktopNum := 0, desktopName := "") {
  }
  static VirtualDesktopWallpaperChanged(desktopNum := 0, desktopWallpaper := "") {
  }
  static SetForegroundWindow(hWnd) {
    if (DllCall("AllowSetForegroundWindow", "Uint", DllCall("GetCurrentProcessId"))) {
      DllCall("SetForegroundWindow", "Ptr", hwnd)
    } else {
      LCtrlDown := GetKeyState("LCtrl")
      RCtrlDown := GetKeyState("RCtrl")
      LShiftDown := GetKeyState("LShift")
      RShiftDown := GetKeyState("RShift")
      LWinDown := GetKeyState("LWin")
      RWinDown := GetKeyState("RWin")
      LAltDown := GetKeyState("LAlt")
      RAltDown := GetKeyState("RAlt")
      if ((LCtrlDown || RCtrlDown) && (LWinDown || RWinDown)) {
        toRelease := ""
        if (LShiftDown) {
          toRelease .= "{LShift Up}"
        }
        if (RShiftDown) {
          toRelease .= "{RShift Up}"
        }
        if (toRelease) {
          Send "{Blind}" toRelease
        }
      }
      BlockInput "On"
      Send "{LAlt Down}{LAlt Down}"
      DllCall("SetForegroundWindow", "Ptr", hwnd)
      toAppend := ""
      if (!LAltDown) {
        toAppend .= "{LAlt Up}"
      }
      if (RAltDown) {
        toAppend .= "{RAlt Down}"
      }
      if (LCtrlDown) {
        toAppend .= "{LCtrl Down}"
      }
      if (RCtrlDown) {
        toAppend .= "{RCtrl Down}"
      }
      if (LShiftDown) {
        toAppend .= "{LShift Down}"
      }
      if (RShiftDown) {
        toAppend .= "{RShift Down}"
      }
      if (LWinDown) {
        toAppend .= "{LWin Down}"
      }
      if (RWinDown) {
        toAppend .= "{RWin Down}"
      }
      if (toAppend) {
        Send "{Blind}" toAppend
      }
      BlockInput "Off"
    }
  }
  static _getFirstWindowInVD(desktopNum, excludeHwnd := 0) {
    bak_DetectHiddenWindows := A_DetectHiddenWindows
    DetectHiddenWindows true
    returnValue := 0
    outHwndList := WinGetList()
    GUID_Desktop := Buffer(16)
    IVirtualDesktop := this._GetDesktops_Obj().GetAt(desktopNum)
    ptr_GetId := this._vtable(IVirtualDesktop, this.idx_GetId)
    DllCall(ptr_GetId, "Ptr", IVirtualDesktop, "Ptr", GUID_Desktop)
    n1 := NumGet(GUID_Desktop, 0x0, "Int64")
    n2 := NumGet(GUID_Desktop, 0x8, "Int64")
    loop outHwndList.Length {
      theHwnd := outHwndList[A_Index] + 0
      if (theHwnd == excludeHwnd) {
        continue
      }
      arr_success_pView_hWnd := this._isValidWindow(theHwnd)
      if (arr_success_pView_hWnd[1] == 0) {
        thePView := arr_success_pView_hWnd[2]
        OutputVar_MinMax := WinGetMinMax("ahk_id " theHwnd)
        if (!(OutputVar_MinMax == -1)) {
          ptr_GetVirtualDesktopId := this._vtable(thePView, 25)
          DllCall(ptr_GetVirtualDesktopId, "Ptr", thePView, "Ptr", GUID_Desktop)
          if (n1 == NumGet(GUID_Desktop, 0x0, "Int64") && n2 == NumGet(GUID_Desktop, 0x8, "Int64")) {
            returnValue := theHwnd
            break
          }
        }
      }
    }
    DetectHiddenWindows bak_DetectHiddenWindows
    return returnValue
  }
  static _tryGetValidWindow(wintitle) {
    bak_DetectHiddenWindows := A_DetectHiddenWindows
    bak_TitleMatchMode := A_TitleMatchMode
    DetectHiddenWindows true
    SetTitleMatchMode "2"
    outHwndList := WinGetList(wintitle)
    returnValue := false
    loop outHwndList.Length {
      theHwnd := outHwndList[A_Index] + 0
      arr_success_pView_hWnd := this._isValidWindow(theHwnd)
      pView := arr_success_pView_hWnd[2]
      if (pView) {
        returnValue := [arr_success_pView_hWnd[3], pView]
        break
      }
    }
    SetTitleMatchMode bak_TitleMatchMode
    DetectHiddenWindows bak_DetectHiddenWindows
    return returnValue
  }
  static _desktopNum_from_IVirtualDesktop(IVirtualDesktop) {
    Desktops_Obj := this._GetDesktops_Obj()
    loop Desktops_Obj.GetCount() {
      IVirtualDesktop_ofDesktop := Desktops_Obj.GetAt(A_Index)
      if (IVirtualDesktop_ofDesktop == IVirtualDesktop) {
        return A_Index
      }
    }
    return 0
  }
  static _desktopNum_from_pView(thePView) {
    GUID_Desktop := Buffer(16)
    ptr_GetVirtualDesktopId := this._vtable(thePView, 25)
    DllCall(ptr_GetVirtualDesktopId, "Ptr", thePView, "Ptr", GUID_Desktop)
    n1 := NumGet(GUID_Desktop, 0x0, "Int64")
    n2 := NumGet(GUID_Desktop, 0x8, "Int64")
    Desktops_Obj := this._GetDesktops_Obj()
    loop Desktops_Obj.GetCount() {
      IVirtualDesktop := Desktops_Obj.GetAt(A_Index)
      ptr_GetId := this._vtable(IVirtualDesktop, this.idx_GetId)
      DllCall(ptr_GetId, "Ptr", IVirtualDesktop, "Ptr", GUID_Desktop)
      if (n1 == NumGet(GUID_Desktop, 0x0, "Int64") && n2 == NumGet(GUID_Desktop, 0x8, "Int64")) {
        return A_Index
      }
    }
    return 0
  }
  static _GetDesktops_Obj() {
    IObjectArray := this._dll_GetDesktops()
    return this.IObjectArray_Wrapper(IObjectArray, this.IID_IVirtualDesktop_ptr)
  }
  static _isValidWindow(hWnd, checkUpper := true) {
    static GetWindowLong_Name := (A_PtrSize == 8) ? "GetWindowLongPtrW" : "GetWindowLongW"
    returnValue := [1, 0, 0]
    breakToReturnFalse:
    loop 1 {
      dwStyle := DllCall(GetWindowLong_Name, "Ptr", hWnd, "Int", -16, "Ptr")
      if (!(dwStyle & 0x10000000)) {
        break breakToReturnFalse
      }
      dwExStyle := DllCall(GetWindowLong_Name, "Ptr", hWnd, "Int", -20, "Ptr")
      if (!(dwExStyle & 0x00040000)) {
        if (dwExStyle & 0x00000080 || dwExStyle & 0x08000000) {
          break breakToReturnFalse
        }
        if (checkUpper) {
          toCheck := []
          upHwnd := hWnd
          while (upHwnd := DllCall("GetWindow", "Ptr", upHwnd, "Uint", 4)) {
            if (upHwnd == 65552) {
              break breakToReturnFalse
            }
            toCheck.Push(upHwnd)
          }
          i := toCheck.Length + 1
          while (i-- > 1) {
            arr_success_pView_hWnd := this._isValidWindow(toCheck[i], false)
            if (arr_success_pView_hWnd[1] == 0) {
              arr_success_pView_hWnd[1] := 2
              returnValue := arr_success_pView_hWnd
              break breakToReturnFalse
            }
          }
        }
      }
      pView := this._dll_GetViewForHwnd(hWnd)
      if (!pView) {
        break breakToReturnFalse
      }
      returnValue := [0, pView, hWnd]
    }
    return returnValue
  }
  static _vtable(ppv, index) {
    Return NumGet(NumGet(0 + ppv, 0, "Ptr") + A_PtrSize * index, 0, "Ptr")
  }
  class IObjectArray_Wrapper {
    __New(IObjectArray, IID_Interface_ptr) {
      this.IObjectArray := IObjectArray
      this.IID_Interface_ptr := IID_Interface_ptr
      this.ptr_GetAt := VD._vtable(IObjectArray, 4)
    }
    __Delete() {
      ptr_Release := VD._vtable(this.IObjectArray, 2)
      DllCall(ptr_Release, "Ptr", this.IObjectArray)
    }
    GetAt(oneBasedIndex) {
      DllCall(this.ptr_GetAt, "Ptr", this.IObjectArray, "UInt", oneBasedIndex - 1, "Ptr", this.IID_Interface_ptr, "Ptr*", &IInterface := 0)
      return IInterface
    }
    GetCount() {
      ptr_GetCount := VD._vtable(this.IObjectArray, 3)
      Count := 0
      DllCall(ptr_GetCount, "Ptr", this.IObjectArray, "UInt*", &Count := 0)
      return Count
    }
  }
}
