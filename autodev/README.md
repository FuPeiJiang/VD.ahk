# Virtual desktops implemented entirely in AutoDev JavaScript

**`VD.mjs` + `VDCom.mjs` are the complete runtime.** There is no custom DLL,
AutoHotkey script, PowerShell bridge, C/C++ extension or Node.js dependency.
These are AutoDev QuickJS modules (with `VD.d.mts` TypeScript declarations).
They call Windows **system-provided** `ole32.dll`, `oleaut32.dll`,
`combase.dll`, `version.dll`, `kernel32.dll`, and `user32.dll` using AutoDev's
built-in `ffi`. Windows' private virtual desktop COM interfaces are vtable-only,
so `VDCom.mjs` uses the documented `DispCallFunc` entry point to invoke them.

## Install

Keep `VD.mjs` and `VDCom.mjs` in the same directory as the entry module.
Run with AutoDev, **not** with Node or the browser:

```powershell
autodev run --module .\VD.example.mjs
```

```js
import { VD } from './VD.mjs';
const desktops = new VD();
try {
  console.log(desktops.getInfo());
  desktops.goToRelativeDesktopNum(1);  // wrap to next desktop
  // desktops.moveWindowToDesktopNum('A', 2, { follow: true });
  // desktops.togglePinWindow('A');
} finally {
  desktops.close();
}
```

`VD.hotkeys.mjs` demonstrates real native AutoDev hotkey handlers.
Use `autodev run --module --persistent .\VD.hotkeys.mjs` to keep those active.

## API

- `getCount()`, `getCurrentDesktopNum()`, `getVersion()`, `getInfo()`
- `goToDesktopNum(n)`, `goToRelativeDesktopNum(offset)`, `waitDesktopSwitched(n)`
- `createDesktop()`, `createUntil(count)`, `removeDesktop(n, fallback)`
- `getDesktopNumOfWindow(hwnd)`, `goToDesktopOfWindow(hwnd)`
- `moveWindowToDesktopNum(hwnd,n)`, `moveWindowToRelativeDesktopNum`, `moveWindowToCurrentDesktop`
- `isWindowPinned`, `pinWindow`, `unpinWindow`, `togglePinWindow`
- `isExePinned`, `pinExe`, `unpinExe`, `togglePinExe` (these accept **AppUserModelID**, not necessarily a filesystem executable path)
- `getNameFromDesktopNum`, `setNameToDesktopNum`
- `getWallpaperFromDesktopNum`, `setWallpaperToDesktopNum`
- `close()` releases AutoDev's loaded system-library handles.

Desktop numbers are one-based. Negative identifiers `-1` and `-2` identify
pinned applications/windows across all desktops and are not array indexes.
Window targets accept positive HWND, `{hwnd}`, `'A'` (foreground), or a uniquely
matching title. Ambiguous titles throw instead of silently targeting an
arbitrary window. `follow`/`activate` requests cannot override Windows' own
foreground activation policy.

## Windows compatibility and caveats

The layout is selected using **twinui.pcshell.dll's** build/revision, matching
the baseline `VD.ahk`; the Windows OS build can differ from this file version.
The underlying COM interfaces are **undocumented** and may change in future
Windows builds. The library releases references after every operation so it
does not keep stale COM pointers between calls, but it does not yet offer
native desktop notification subscriptions or preserve pinned foreground focus
history. These are areas for future AutoDev runtime enhancements.

## Tests

`smoke.mjs` creates, renames, switches, moves a test GUI, pins/unpins it and
removes the disposable desktop. **Never run the smoke test on your real
Windows desktop**. The KVM runner creates a throwaway qcow2 overlay and never
boots the base disk for writing:

```bash
./tests/kvm-runner.sh
```

Verified on the disposable Windows 11 guest with twinui version 26100.9168:
20 checks green for the pure JavaScript implementation. Only the read-only
`VD.example.mjs` was exercised on the live `desktop` host. Window-specific
focus restoration, wallpaper writes, app-ID pinning, and Explorer restart
fault injection remain outside the present smoke coverage.
