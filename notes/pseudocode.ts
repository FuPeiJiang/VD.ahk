function MoveWindowToDesktopNum(wintitle: string, desktopNum: number, follow = false, WinActivatePriority = "new window") {
    // enum WinActivatePriority
    // - new window ; WinActivate new
    // - old window - Exclude Desktop ; WinActivate old if old !== Desktop, else WinActivate new // (ensure a Window is always active)
    // - old window - Include Desktop ; WinActivate old

    // [follow = false]
    // srcVD (Window's VD) == dstVD ; do nothing

    // [follow = false]
    // activeWindow -> CurrentVD ; do nothing
    // activeWindow -> OtherVD ; WinActivate window under
    // non-activeWindow -> CurrentVD ; WinActivate following WinActivatePriority
    // - new window ; WinActivate
    // - old window - Exclude Desktop ; do nothing
    // - old window - Include Desktop ; do nothing
    // non-activeWindow -> OtherVD ; do nothing
    // otherVD-Window -> CurrentVD ; WinActivate following WinActivatePriority
    // - otherVD-Window -> CurrentVD (WinActivate new)
    // - otherVD-Window -> CurrentVD (WinActivate old)
    //   - otherVD-Window -> CurrentVD (WinActivate old - Exclude Desktop)
    //   - otherVD-Window -> CurrentVD (WinActivate old - Include Desktop)
    // otherVD-Window -> OtherVD ; do nothing

    // [follow = true]
    // dstVD == CurrentVD ; do not Switch

    // [follow = true]
    // activeWindow -> CurrentVD ; do nothing
    // activeWindow -> OtherVD ; Move then Switch then WinActive following WinActivatePriority
    // non-activeWindow -> CurrentVD ; Move then !Switch then WinActive following WinActivatePriority
    // non-activeWindow -> OtherVD ; Move then Switch then WinActive following WinActivatePriority
    // otherVD-Window -> CurrentVD ; Move then !Switch then WinActive following WinActivatePriority
    // otherVD-Window -> OtherVD ; Move then Switch then WinActive following WinActivatePriority
}