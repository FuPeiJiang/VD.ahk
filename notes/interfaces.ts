const versions = [{
    //from 17763.1
    buildNumber:20348,
    revisionNumber:0,
    IID_IVirtualDesktopManagerInternal_str:"{f31574d6-b682-4cdc-bd56-1827860abec6}",
    IID_IVirtualDesktop_str:"{ff72ffdd-be7e-43fc-9c03-ad81681e88e4}",
    IID_IVirtualDesktopNotification_str:"{c179334c-4295-40d3-bea1-c654d965605a}",

    idx_MoveViewToDesktop:4, //DllCall(ptr_MoveViewToDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IApplicationView,"Ptr",IVirtualDesktop)
    idx_GetCurrentDesktop:6, //DllCall(ptr_GetCurrentDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr*",&IVirtualDesktop:0),
    idx_GetDesktops:7, //DllCall(ptr_GetDesktops,"Ptr",IVirtualDesktopManagerInternal,"Ptr*",&IObjectArray:0),
    idx_SwitchDesktop:9, //DllCall(ptr_SwitchDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop)
    idx_CreateDesktop:10, //DllCall(ptr_CreateDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr*",&IVirtualDesktop:0),
    idx_RemoveDesktop:11, //DllCall(ptr_RemoveDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",IVirtualDesktop_fallback)
    idx_SwitchDesktopWithAnimation:-1,
    _dll_MoveViewToDesktop:"_dll_MoveViewToDesktop_normal",
    _dll_GetCurrentDesktop:"_dll_GetCurrentDesktop_normal",
    _dll_GetDesktops:"_dll_GetDesktops_normal",
    _dll_SwitchDesktop:"_dll_SwitchDesktop_normal",
    _dll_CreateDesktop:"_dll_CreateDesktop_normal",
    _dll_RemoveDesktop:"_dll_RemoveDesktop_normal",

    idx_GetId:4, //DllCall(ptr_GetId,"Ptr",IVirtualDesktop,"Ptr",guid_buf)

    idx_VirtualDesktopNameChanged:-1,
    idx_SetDesktopName:-1,
    idx_GetName:-1,

    idx_VirtualDesktopWallpaperChanged:-1,
    idx_SetDesktopWallpaper:-1,
    idx_GetWallpaper:-1,

    idx_VirtualDesktopCreated:3, //params (IVirtualDesktop)
    idx_VirtualDesktopDestroyBegin:4, //params (IVirtualDesktop, IVirtualDesktop_fallback)
    idx_VirtualDesktopDestroyFailed:5, //params (IVirtualDesktop, IVirtualDesktop_fallback)
    idx_VirtualDesktopDestroyed:6, //params (IVirtualDesktop, IVirtualDesktop_fallback)
    idx_ViewVirtualDesktopChanged:7, //params (IApplicationView)
    idx_CurrentVirtualDesktopChanged:8, //params (IVirtualDesktop_old, IVirtualDesktop_new)
    _dll_VirtualDesktopCreated:"_dll_VirtualDesktopCreated_normal",
    _dll_VirtualDesktopDestroyBegin:"_dll_VirtualDesktopDestroyBegin_normal",
    _dll_VirtualDesktopDestroyFailed:"_dll_VirtualDesktopDestroyFailed_normal",
    _dll_VirtualDesktopDestroyed:"_dll_VirtualDesktopDestroyed_normal",
    _dll_ViewVirtualDesktopChanged:"_dll_ViewVirtualDesktopChanged_normal",
    _dll_CurrentVirtualDesktopChanged:"_dll_CurrentVirtualDesktopChanged_normal",

    IVirtualDesktopNotification_methods_count:9,
}, { //22000.51 to be more precise
    //from 20348.2227 - Windows Server 2022
    buildNumber:22000,
    revisionNumber:0,
    IID_IVirtualDesktopManagerInternal_str:"{094afe11-44f2-4ba0-976f-29a97e263ee0}",
    IID_IVirtualDesktop_str:"{62fdf88b-11ca-4afb-8bd8-2296dfae49e2}",
    IID_IVirtualDesktopNotification_str:"{f3163e11-6b04-433c-a64b-6f82c9094257}",

    idx_MoveViewToDesktop:4, //DllCall(ptr_MoveViewToDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IApplicationView,"Ptr",IVirtualDesktop)
    idx_GetCurrentDesktop:6, //DllCall(ptr_GetCurrentDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",HMONITOR,"Ptr*",&IVirtualDesktop:0),
    idx_GetDesktops:7, //DllCall(ptr_GetDesktops,"Ptr",IVirtualDesktopManagerInternal,"Ptr",HMONITOR,"Ptr*",&IObjectArray:0),
    idx_SwitchDesktop:9, //DllCall(ptr_SwitchDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",HMONITOR,"Ptr",IVirtualDesktop)
    idx_CreateDesktop:10, //DllCall(ptr_CreateDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",HMONITOR,"Ptr*",&IVirtualDesktop:0),
    idx_RemoveDesktop:11, //DllCall(ptr_RemoveDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",IVirtualDesktop_fallback)
    idx_SwitchDesktopWithAnimation:-1,
    _dll_MoveViewToDesktop:"_dll_MoveViewToDesktop_normal",
    _dll_GetCurrentDesktop:"_dll_GetCurrentDesktop_HMONITOR",
    _dll_GetDesktops:"_dll_GetDesktops_HMONITOR",
    _dll_SwitchDesktop:"_dll_SwitchDesktop_HMONITOR",
    _dll_CreateDesktop:"_dll_CreateDesktop_HMONITOR",
    _dll_RemoveDesktop:"_dll_RemoveDesktop_normal",

    idx_GetId:4, //DllCall(ptr_GetId,"Ptr",IVirtualDesktop,"Ptr",guid_buf)

    idx_VirtualDesktopNameChanged:8, //params (IVirtualDesktop, HSTRING)
    _dll_VirtualDesktopNameChanged:"_dll_VirtualDesktopNameChanged_normal",
    idx_SetDesktopName:14, //DllCall(ptr_SetDesktopName,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",HSTRING)
    idx_GetName:6, //DllCall(ptr_GetName,"Ptr",IVirtualDesktop,"Ptr*",HSTRING)

    idx_VirtualDesktopWallpaperChanged:-1,
    idx_SetDesktopWallpaper:-1,
    idx_GetWallpaper:-1,

    idx_VirtualDesktopCreated:3, //params (IVirtualDesktop)
    idx_VirtualDesktopDestroyBegin:4, //params (IVirtualDesktop, IVirtualDesktop_fallback)
    idx_VirtualDesktopDestroyFailed:5, //params (IVirtualDesktop, IVirtualDesktop_fallback)
    idx_VirtualDesktopDestroyed:6, //params (IVirtualDesktop, IVirtualDesktop_fallback)
    idx_ViewVirtualDesktopChanged:9, //params (IApplicationView)
    idx_CurrentVirtualDesktopChanged:10, //params (IVirtualDesktop_old, IVirtualDesktop_new)
    _dll_VirtualDesktopCreated:"_dll_VirtualDesktopCreated_normal",
    _dll_VirtualDesktopDestroyBegin:"_dll_VirtualDesktopDestroyBegin_normal",
    _dll_VirtualDesktopDestroyFailed:"_dll_VirtualDesktopDestroyFailed_normal",
    _dll_VirtualDesktopDestroyed:"_dll_VirtualDesktopDestroyed_normal",
    _dll_ViewVirtualDesktopChanged:"_dll_ViewVirtualDesktopChanged_normal",
    _dll_CurrentVirtualDesktopChanged:"_dll_CurrentVirtualDesktopChanged_normal",

    IVirtualDesktopNotification_methods_count:11,
}, { //22483.1000 to be more precise
    //from 22000.51
    buildNumber:22483,
    revisionNumber:0,
    IID_IVirtualDesktopManagerInternal_str:"{b2f925b9-5a0f-4d2e-9f4d-2b1507593c10}",
    IID_IVirtualDesktop_str:"{536d3495-b208-4cc9-ae26-de8111275bf8}",
    IID_IVirtualDesktopNotification_str:"{cd403e52-deed-4c13-b437-b98380f2b1e8}",

    idx_MoveViewToDesktop:4, //DllCall(ptr_MoveViewToDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IApplicationView,"Ptr",IVirtualDesktop)
    idx_GetCurrentDesktop:6, //DllCall(ptr_GetCurrentDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",HMONITOR,"Ptr*",&IVirtualDesktop:0),
    idx_GetDesktops:7, //DllCall(ptr_GetDesktops,"Ptr",IVirtualDesktopManagerInternal,"Ptr",HMONITOR,"Ptr*",&IObjectArray:0),
    idx_SwitchDesktop:9, //DllCall(ptr_SwitchDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",HMONITOR,"Ptr",IVirtualDesktop)
    idx_CreateDesktop:10, //DllCall(ptr_CreateDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",HMONITOR,"Ptr*",&IVirtualDesktop:0),
    idx_RemoveDesktop:12, //DllCall(ptr_RemoveDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",IVirtualDesktop_fallback)
    idx_SwitchDesktopWithAnimation:-1,
    _dll_MoveViewToDesktop:"_dll_MoveViewToDesktop_normal",
    _dll_GetCurrentDesktop:"_dll_GetCurrentDesktop_HMONITOR",
    _dll_GetDesktops:"_dll_GetDesktops_HMONITOR",
    _dll_SwitchDesktop:"_dll_SwitchDesktop_HMONITOR",
    _dll_CreateDesktop:"_dll_CreateDesktop_HMONITOR",
    _dll_RemoveDesktop:"_dll_RemoveDesktop_normal",

    idx_GetId:4, //DllCall(ptr_GetId,"Ptr",IVirtualDesktop,"Ptr",guid_buf)

    idx_VirtualDesktopNameChanged:9, //params (IVirtualDesktop, HSTRING)
    _dll_VirtualDesktopNameChanged:"_dll_VirtualDesktopNameChanged_normal",
    idx_SetDesktopName:15, //DllCall(ptr_SetDesktopName,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",HSTRING)
    idx_GetName:6, //DllCall(ptr_GetName,"Ptr",IVirtualDesktop,"Ptr*",HSTRING)

    idx_VirtualDesktopWallpaperChanged:12, //params (IVirtualDesktop, HSTRING)
    _dll_VirtualDesktopWallpaperChanged:"_dll_VirtualDesktopWallpaperChanged_normal",
    idx_SetDesktopWallpaper:16, //DllCall(ptr_SetDesktopWallpaper,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",HSTRING)
    idx_GetWallpaper:7, //DllCall(ptr_GetWallpaper,"Ptr",IVirtualDesktop,"Ptr*",HSTRING)

    idx_VirtualDesktopCreated:3, //params (IObjectArray, IVirtualDesktop)
    idx_VirtualDesktopDestroyBegin:4, //params (IObjectArray, IVirtualDesktop, IVirtualDesktop_fallback)
    idx_VirtualDesktopDestroyFailed:5, //params (IObjectArray, IVirtualDesktop, IVirtualDesktop_fallback)
    idx_VirtualDesktopDestroyed:6, //params (IObjectArray, IVirtualDesktop, IVirtualDesktop_fallback)
    idx_ViewVirtualDesktopChanged:10, //params (IApplicationView)
    idx_CurrentVirtualDesktopChanged:11, //params (IObjectArray, IVirtualDesktop_old, IVirtualDesktop_new)
    _dll_VirtualDesktopCreated:"_dll_VirtualDesktopCreated_IObjectArray",
    _dll_VirtualDesktopDestroyBegin:"_dll_VirtualDesktopDestroyBegin_IObjectArray",
    _dll_VirtualDesktopDestroyFailed:"_dll_VirtualDesktopDestroyFailed_IObjectArray",
    _dll_VirtualDesktopDestroyed:"_dll_VirtualDesktopDestroyed_IObjectArray",
    _dll_ViewVirtualDesktopChanged:"_dll_ViewVirtualDesktopChanged_normal",
    _dll_CurrentVirtualDesktopChanged:"_dll_CurrentVirtualDesktopChanged_IObjectArray",

    IVirtualDesktopNotification_methods_count:13,
}, {
    //from 22483.1000
    //from 22621.1778 (they're identical)
    buildNumber:22621,
    revisionNumber:2215,
    //yeah yeah, IID are the same as above, but vftable differs
    IID_IVirtualDesktopManagerInternal_str:"{b2f925b9-5a0f-4d2e-9f4d-2b1507593c10}",
    IID_IVirtualDesktop_str:"{536d3495-b208-4cc9-ae26-de8111275bf8}",
    IID_IVirtualDesktopNotification_str:"{cd403e52-deed-4c13-b437-b98380f2b1e8}",

    idx_MoveViewToDesktop:4, //DllCall(ptr_MoveViewToDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IApplicationView,"Ptr",IVirtualDesktop)
    idx_GetCurrentDesktop:6, //DllCall(ptr_GetCurrentDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",HMONITOR,"Ptr*",&IVirtualDesktop:0),
    idx_GetDesktops:8, //DllCall(ptr_GetDesktops,"Ptr",IVirtualDesktopManagerInternal,"Ptr",HMONITOR,"Ptr*",&IObjectArray:0),
    idx_SwitchDesktop:10, //DllCall(ptr_SwitchDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",HMONITOR,"Ptr",IVirtualDesktop)
    idx_CreateDesktop:11, //DllCall(ptr_CreateDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",HMONITOR,"Ptr*",&IVirtualDesktop:0),
    idx_RemoveDesktop:13, //DllCall(ptr_RemoveDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",IVirtualDesktop_fallback)
    idx_SwitchDesktopWithAnimation:-1,
    _dll_MoveViewToDesktop:"_dll_MoveViewToDesktop_normal",
    _dll_GetCurrentDesktop:"_dll_GetCurrentDesktop_HMONITOR",
    _dll_GetDesktops:"_dll_GetDesktops_HMONITOR",
    _dll_SwitchDesktop:"_dll_SwitchDesktop_HMONITOR",
    _dll_CreateDesktop:"_dll_CreateDesktop_HMONITOR",
    _dll_RemoveDesktop:"_dll_RemoveDesktop_normal",

    idx_GetId:4, //DllCall(ptr_GetId,"Ptr",IVirtualDesktop,"Ptr",guid_buf)

    idx_VirtualDesktopNameChanged:9, //params (IVirtualDesktop, HSTRING)
    _dll_VirtualDesktopNameChanged:"_dll_VirtualDesktopNameChanged_normal",
    idx_SetDesktopName:16, //DllCall(ptr_SetDesktopName,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",HSTRING)
    idx_GetName:6, //DllCall(ptr_GetName,"Ptr",IVirtualDesktop,"Ptr*",HSTRING)

    idx_VirtualDesktopWallpaperChanged:12, //params (IVirtualDesktop, HSTRING)
    _dll_VirtualDesktopWallpaperChanged:"_dll_VirtualDesktopWallpaperChanged_normal",
    idx_SetDesktopWallpaper:17, //DllCall(ptr_SetDesktopWallpaper,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",HSTRING)
    idx_GetWallpaper:7, //DllCall(ptr_GetWallpaper,"Ptr",IVirtualDesktop,"Ptr*",HSTRING)

    idx_VirtualDesktopCreated:3, //params (IObjectArray, IVirtualDesktop)
    idx_VirtualDesktopDestroyBegin:4, //params (IObjectArray, IVirtualDesktop, IVirtualDesktop_fallback)
    idx_VirtualDesktopDestroyFailed:5, //params (IObjectArray, IVirtualDesktop, IVirtualDesktop_fallback)
    idx_VirtualDesktopDestroyed:6, //params (IObjectArray, IVirtualDesktop, IVirtualDesktop_fallback)
    idx_ViewVirtualDesktopChanged:10, //params (IApplicationView)
    idx_CurrentVirtualDesktopChanged:11, //params (IObjectArray, IVirtualDesktop_old, IVirtualDesktop_new)
    _dll_VirtualDesktopCreated:"_dll_VirtualDesktopCreated_IObjectArray",
    _dll_VirtualDesktopDestroyBegin:"_dll_VirtualDesktopDestroyBegin_IObjectArray",
    _dll_VirtualDesktopDestroyFailed:"_dll_VirtualDesktopDestroyFailed_IObjectArray",
    _dll_VirtualDesktopDestroyed:"_dll_VirtualDesktopDestroyed_IObjectArray",
    _dll_ViewVirtualDesktopChanged:"_dll_ViewVirtualDesktopChanged_normal",
    _dll_CurrentVirtualDesktopChanged:"_dll_CurrentVirtualDesktopChanged_IObjectArray",

    IVirtualDesktopNotification_methods_count:13,
}, {
    //from 22621.2215
    buildNumber:22631,
    revisionNumber:3085,
    IID_IVirtualDesktopManagerInternal_str:"{a3175f2d-239c-4bd2-8aa0-eeba8b0b138e}",
    IID_IVirtualDesktop_str:"{3f07f4be-b107-441a-af0f-39d82529072c}",
    IID_IVirtualDesktopNotification_str:"{b287fa1c-7771-471a-a2df-9b6b21f0d675}",

    idx_MoveViewToDesktop:4, //DllCall(ptr_MoveViewToDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IApplicationView,"Ptr",IVirtualDesktop)
    idx_GetCurrentDesktop:6, //DllCall(ptr_GetCurrentDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr*",&IVirtualDesktop:0),
    idx_GetDesktops:7, //DllCall(ptr_GetDesktops,"Ptr",IVirtualDesktopManagerInternal,"Ptr*",&IObjectArray:0),
    idx_SwitchDesktop:9, //DllCall(ptr_SwitchDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop)
    idx_CreateDesktop:10, //DllCall(ptr_CreateDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr*",&IVirtualDesktop:0),
    idx_RemoveDesktop:12, //DllCall(ptr_RemoveDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",IVirtualDesktop_fallback)
    idx_SwitchDesktopWithAnimation:21, //DllCall(ptr_SwitchDesktopWithAnimation,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop)
    _dll_MoveViewToDesktop:"_dll_MoveViewToDesktop_normal",
    _dll_GetCurrentDesktop:"_dll_GetCurrentDesktop_normal",
    _dll_GetDesktops:"_dll_GetDesktops_normal",
    _dll_SwitchDesktop:"_dll_SwitchDesktop_normal",
    _dll_CreateDesktop:"_dll_CreateDesktop_normal",
    _dll_RemoveDesktop:"_dll_RemoveDesktop_normal",

    idx_GetId:4, //DllCall(ptr_GetId,"Ptr",IVirtualDesktop,"Ptr",guid_buf)

    idx_VirtualDesktopNameChanged:8, //params (IVirtualDesktop, HSTRING)
    _dll_VirtualDesktopNameChanged:"_dll_VirtualDesktopNameChanged_normal",
    idx_SetDesktopName:15, //DllCall(ptr_SetDesktopName,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",HSTRING)
    idx_GetName:5, //DllCall(ptr_GetName,"Ptr",IVirtualDesktop,"Ptr*",HSTRING)

    idx_VirtualDesktopWallpaperChanged:11, //params (IVirtualDesktop, HSTRING)
    _dll_VirtualDesktopWallpaperChanged:"_dll_VirtualDesktopWallpaperChanged_normal",
    idx_SetDesktopWallpaper:16, //DllCall(ptr_SetDesktopWallpaper,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",HSTRING)
    idx_GetWallpaper:6, //DllCall(ptr_GetWallpaper,"Ptr",IVirtualDesktop,"Ptr*",HSTRING)

    idx_VirtualDesktopCreated:3, //params (IVirtualDesktop)
    idx_VirtualDesktopDestroyBegin:4, //params (IVirtualDesktop, IVirtualDesktop_fallback)
    idx_VirtualDesktopDestroyFailed:5, //params (IVirtualDesktop, IVirtualDesktop_fallback)
    idx_VirtualDesktopDestroyed:6, //params (IVirtualDesktop, IVirtualDesktop_fallback)
    idx_ViewVirtualDesktopChanged:9, //params (IApplicationView)
    idx_CurrentVirtualDesktopChanged:10, //params (IVirtualDesktop_old, IVirtualDesktop_new)
    _dll_VirtualDesktopCreated:"_dll_VirtualDesktopCreated_normal",
    _dll_VirtualDesktopDestroyBegin:"_dll_VirtualDesktopDestroyBegin_normal",
    _dll_VirtualDesktopDestroyFailed:"_dll_VirtualDesktopDestroyFailed_normal",
    _dll_VirtualDesktopDestroyed:"_dll_VirtualDesktopDestroyed_normal",
    _dll_ViewVirtualDesktopChanged:"_dll_ViewVirtualDesktopChanged_normal",
    _dll_CurrentVirtualDesktopChanged:"_dll_CurrentVirtualDesktopChanged_normal",

    IVirtualDesktopNotification_methods_count:14,
}, {
    //from 22631.3085
    buildNumber:26100,
    revisionNumber:0,
    //the only difference with the above is IID_IVirtualDesktopNotification
    IID_IVirtualDesktopManagerInternal_str:"{53f5ca0b-158f-4124-900c-057158060b27}",
    IID_IVirtualDesktop_str:"{3f07f4be-b107-441a-af0f-39d82529072c}",
    IID_IVirtualDesktopNotification_str:"{b9e5e94d-233e-49ab-af5c-2b4541c3aade}",

    idx_MoveViewToDesktop:4, //DllCall(ptr_MoveViewToDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IApplicationView,"Ptr",IVirtualDesktop)
    idx_GetCurrentDesktop:6, //DllCall(ptr_GetCurrentDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr*",&IVirtualDesktop:0),
    idx_GetDesktops:7, //DllCall(ptr_GetDesktops,"Ptr",IVirtualDesktopManagerInternal,"Ptr*",&IObjectArray:0),
    idx_SwitchDesktop:9, //DllCall(ptr_SwitchDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop)
    idx_CreateDesktop:10, //DllCall(ptr_CreateDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr*",&IVirtualDesktop:0),
    idx_RemoveDesktop:12, //DllCall(ptr_RemoveDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",IVirtualDesktop_fallback)
    idx_SwitchDesktopWithAnimation:21, //DllCall(ptr_SwitchDesktopWithAnimation,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop)
    _dll_MoveViewToDesktop:"_dll_MoveViewToDesktop_normal",
    _dll_GetCurrentDesktop:"_dll_GetCurrentDesktop_normal",
    _dll_GetDesktops:"_dll_GetDesktops_normal",
    _dll_SwitchDesktop:"_dll_SwitchDesktop_normal",
    _dll_CreateDesktop:"_dll_CreateDesktop_normal",
    _dll_RemoveDesktop:"_dll_RemoveDesktop_normal",

    idx_GetId:4, //DllCall(ptr_GetId,"Ptr",IVirtualDesktop,"Ptr",guid_buf)

    idx_VirtualDesktopNameChanged:8, //params (IVirtualDesktop, HSTRING)
    _dll_VirtualDesktopNameChanged:"_dll_VirtualDesktopNameChanged_normal",
    idx_SetDesktopName:15, //DllCall(ptr_SetDesktopName,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",HSTRING)
    idx_GetName:5, //DllCall(ptr_GetName,"Ptr",IVirtualDesktop,"Ptr*",HSTRING)

    idx_VirtualDesktopWallpaperChanged:11, //params (IVirtualDesktop, HSTRING)
    _dll_VirtualDesktopWallpaperChanged:"_dll_VirtualDesktopWallpaperChanged_normal",
    idx_SetDesktopWallpaper:16, //DllCall(ptr_SetDesktopWallpaper,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",HSTRING)
    idx_GetWallpaper:6, //DllCall(ptr_GetWallpaper,"Ptr",IVirtualDesktop,"Ptr*",HSTRING)

    idx_VirtualDesktopCreated:3, //params (IVirtualDesktop)
    idx_VirtualDesktopDestroyBegin:4, //params (IVirtualDesktop, IVirtualDesktop_fallback)
    idx_VirtualDesktopDestroyFailed:5, //params (IVirtualDesktop, IVirtualDesktop_fallback)
    idx_VirtualDesktopDestroyed:6, //params (IVirtualDesktop, IVirtualDesktop_fallback)
    idx_ViewVirtualDesktopChanged:9, //params (IApplicationView)
    idx_CurrentVirtualDesktopChanged:10, //params (IVirtualDesktop_old, IVirtualDesktop_new)
    _dll_VirtualDesktopCreated:"_dll_VirtualDesktopCreated_normal",
    _dll_VirtualDesktopDestroyBegin:"_dll_VirtualDesktopDestroyBegin_normal",
    _dll_VirtualDesktopDestroyFailed:"_dll_VirtualDesktopDestroyFailed_normal",
    _dll_VirtualDesktopDestroyed:"_dll_VirtualDesktopDestroyed_normal",
    _dll_ViewVirtualDesktopChanged:"_dll_ViewVirtualDesktopChanged_normal",
    _dll_CurrentVirtualDesktopChanged:"_dll_CurrentVirtualDesktopChanged_normal",

    IVirtualDesktopNotification_methods_count:14,
}, {
    //from 26100.863
    buildNumber:99999,
    revisionNumber:0,
    IID_IVirtualDesktopManagerInternal_str:"{53f5ca0b-158f-4124-900c-057158060b27}",
    IID_IVirtualDesktop_str:"{3f07f4be-b107-441a-af0f-39d82529072c}",
    IID_IVirtualDesktopNotification_str:"{b9e5e94d-233e-49ab-af5c-2b4541c3aade}",

    idx_MoveViewToDesktop:4, //DllCall(ptr_MoveViewToDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IApplicationView,"Ptr",IVirtualDesktop)
    idx_GetCurrentDesktop:6, //DllCall(ptr_GetCurrentDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr*",&IVirtualDesktop:0),
    idx_GetDesktops:7, //DllCall(ptr_GetDesktops,"Ptr",IVirtualDesktopManagerInternal,"Ptr*",&IObjectArray:0),
    idx_SwitchDesktop:9, //DllCall(ptr_SwitchDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop)
    idx_CreateDesktop:11, //DllCall(ptr_CreateDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr*",&IVirtualDesktop:0),
    idx_RemoveDesktop:13, //DllCall(ptr_RemoveDesktop,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",IVirtualDesktop_fallback)
    idx_SwitchDesktopWithAnimation:22, //DllCall(ptr_SwitchDesktopWithAnimation,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop)
    _dll_MoveViewToDesktop:"_dll_MoveViewToDesktop_normal",
    _dll_GetCurrentDesktop:"_dll_GetCurrentDesktop_normal",
    _dll_GetDesktops:"_dll_GetDesktops_normal",
    _dll_SwitchDesktop:"_dll_SwitchDesktop_normal",
    _dll_CreateDesktop:"_dll_CreateDesktop_normal",
    _dll_RemoveDesktop:"_dll_RemoveDesktop_normal",

    idx_GetId:4, //DllCall(ptr_GetId,"Ptr",IVirtualDesktop,"Ptr",guid_buf)

    idx_VirtualDesktopNameChanged:8, //params (IVirtualDesktop, HSTRING)
    _dll_VirtualDesktopNameChanged:"_dll_VirtualDesktopNameChanged_normal",
    idx_SetDesktopName:16, //DllCall(ptr_SetDesktopName,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",HSTRING)
    idx_GetName:5, //DllCall(ptr_GetName,"Ptr",IVirtualDesktop,"Ptr*",HSTRING)

    idx_VirtualDesktopWallpaperChanged:11, //params (IVirtualDesktop, HSTRING)
    _dll_VirtualDesktopWallpaperChanged:"_dll_VirtualDesktopWallpaperChanged_normal",
    idx_SetDesktopWallpaper:17, //DllCall(ptr_SetDesktopWallpaper,"Ptr",IVirtualDesktopManagerInternal,"Ptr",IVirtualDesktop,"Ptr",HSTRING)
    idx_GetWallpaper:6, //DllCall(ptr_GetWallpaper,"Ptr",IVirtualDesktop,"Ptr*",HSTRING)

    idx_VirtualDesktopCreated:3, //params (IVirtualDesktop)
    idx_VirtualDesktopDestroyBegin:4, //params (IVirtualDesktop, IVirtualDesktop_fallback)
    idx_VirtualDesktopDestroyFailed:5, //params (IVirtualDesktop, IVirtualDesktop_fallback)
    idx_VirtualDesktopDestroyed:6, //params (IVirtualDesktop, IVirtualDesktop_fallback)
    idx_ViewVirtualDesktopChanged:9, //params (IApplicationView)
    idx_CurrentVirtualDesktopChanged:10, //params (IVirtualDesktop_old, IVirtualDesktop_new)
    _dll_VirtualDesktopCreated:"_dll_VirtualDesktopCreated_normal",
    _dll_VirtualDesktopDestroyBegin:"_dll_VirtualDesktopDestroyBegin_normal",
    _dll_VirtualDesktopDestroyFailed:"_dll_VirtualDesktopDestroyFailed_normal",
    _dll_VirtualDesktopDestroyed:"_dll_VirtualDesktopDestroyed_normal",
    _dll_ViewVirtualDesktopChanged:"_dll_ViewVirtualDesktopChanged_normal",
    _dll_CurrentVirtualDesktopChanged:"_dll_CurrentVirtualDesktopChanged_normal",

    IVirtualDesktopNotification_methods_count:14,
}]

const formatted = versions.map(v => {
const str = `        {
            buildNumber: ${v.buildNumber},
            revisionNumber: ${v.revisionNumber},
            IID_IVirtualDesktopManagerInternal_str: "${v.IID_IVirtualDesktopManagerInternal_str}",
            IID_IVirtualDesktop_str: "${v.IID_IVirtualDesktop_str}",
            IID_IVirtualDesktopNotification_str: "${v.IID_IVirtualDesktopNotification_str}",
            ; vtbl
            IVirtualDesktopManagerInternal: VD.IVirtualDesktopManagerInternal_${v._dll_GetCurrentDesktop === "_dll_GetCurrentDesktop_HMONITOR" ? "HMONITOR" : "Normal"},
            idx_MoveViewToDesktop:${v.idx_MoveViewToDesktop},
            idx_GetCurrentDesktop:${v.idx_GetCurrentDesktop},
            idx_GetDesktops:${v.idx_GetDesktops},
            idx_SwitchDesktop:${v.idx_SwitchDesktop},
            idx_CreateDesktop:${v.idx_CreateDesktop},
            idx_RemoveDesktop:${v.idx_RemoveDesktop},
            ;vtbl Notification
            IVirtualDesktopNotification: VD.IVirtualDesktopNotification_${v._dll_VirtualDesktopCreated === "_dll_VirtualDesktopCreated_IObjectArray" ? "IObjectArray" : "Normal"},
            IVirtualDesktopNotification_methods_count: ${v.IVirtualDesktopNotification_methods_count},
            idx_VirtualDesktopCreated:${v.idx_VirtualDesktopCreated},
            idx_VirtualDesktopDestroyed:${v.idx_VirtualDesktopDestroyed},
            idx_CurrentVirtualDesktopChanged:${v.idx_CurrentVirtualDesktopChanged},
        },`
return str
}).join('\n')

debugger