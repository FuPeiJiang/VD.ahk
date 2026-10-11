# VD for AutoDev (JavaScript)

A JavaScript/QuickJS implementation of Windows virtual-desktop control for **AutoDev**.
It replaces the AutoHotkey runtime with `VD.mjs` + a small x64 `VDNative.dll`
for the undocumented Windows COM virtual-desktop interfaces. **No AutoHotkey
installation is required**. Tested in a **disposable Windows 11 24H2 KVM guest
(build 26100.9168)** with AutoDev 0.1.0.

## Install and run

Copy `VD.mjs` and `VDNative.dll` into the same directory as your entry script.
Use AutoDev's module runner (not Node.js):

```powershell
autodev run --module .\example.mjs
```

```javascript
import { VD } from './VD.mjs';
const dir = autodev.script.path.replace(/[^\\/]+$/, '');
const desktops = new VD({ dllPath: dir + 'VDNative.dll' });
try {
  console.log(desktops.getInfo());
  desktops.goToRelativeDesktopNum(+1);            // switch right with wrap
  // desktops.moveWindowToDesktopNum('A', 2, { follow: true });
  // desktops.togglePinWindow('A');
} finally {
  desktops.close();
}
```

The explicit `dllPath` also works in AutoDev's `eval`/`repl` where
`autodev.script.path` is unset. Windows, not the QuickJS implementation,
controls foreground activation. An `activate`/`follow` request is best effort;
Windows can refuse focus changes.

## API

- `getCount`, `getCurrentDesktopNum`, `getVersion`, `getInfo`
- `goToDesktopNum`, `goToRelativeDesktopNum`, `waitDesktopSwitched`
- `createDesktop`, `createUntil`, `removeDesktop`
- `getDesktopNumOfWindow`, `goToDesktopOfWindow`
- `moveWindowToDesktopNum`, `moveWindowToRelativeDesktopNum`, `moveWindowToCurrentDesktop`
- `isWindowPinned`, `pinWindow`, `unpinWindow`, `togglePinWindow`
- `isExePinned`, `pinExe`, `unpinExe`, `togglePinExe`
- `getNameFromDesktopNum`, `setNameToDesktopNum`
- `getWallpaperFromDesktopNum`, `setWallpaperToDesktopNum`
- `close`

Desktop indexes are **1-based**. Invalid indexes throw before a COM operation;
`goToDesktopOfWindow` returns `-1` for pinned app and `-2` for pinned view.
Windows are resolved by HWND (preferred), `{ hwnd }`, `A` (foreground), or a
unique matching window title. Ambiguous titles throw instead of silently
selecting a random window. `getInfo` performs live queries instead of relying
on cached Explorer COM pointers.

## Rebuild native adapter

On Linux with MinGW-w64:

```bash
x86_64-w64-mingw32-gcc -std=c11 -O2 -shared -Wall -Wextra -Werror \
  -Wno-cast-function-type -o VDNative.dll native/vd_bridge.c -lole32 -lversion
```

The DLL is intentionally a small fixed-ABI adapter, not an injection/hooking
component. Windows' private virtual-desktop interfaces are non-`IDispatch`
interfaces, which AutoDev's regular `com.create(...).call(...)` cannot invoke.
The bridge creates and releases COM interfaces on every operation; no injected
Explorer DLL, persistent COM pointer, or global foreground hook is required.
ABI selection follows the reference `VD.ahk` build/revision table. Because
Microsoft does not provide a stable contract, compatibility with new Windows
builds must be reverified rather than assumed.

## Regressions

`smoke.mjs` uses real desktop state and intentionally creates/removes test
desktops. **Run only inside a disposable KVM overlay**, never as an unattended
smoke test on a personal desktop:

```bash
./tests/kvm-runner.sh
```

The runner boots a fresh throwaway qcow2 child of the pre-existing Windows 11
test image and uses Termirror to copy AutoDev and the candidate files. It
checks version detection, enumeration, bounds, switching, create/remove,
renaming, and a native AutoDev GUI's window/pinning behavior. The base image
is never attached read/write.

Hotkey example: `autodev run --module --persistent VD.hotkeys.mjs` registers
Alt+Shift+Left/Right using AutoDev hotkeys (not AutoHotkey). The module is
optional; close the persistent AutoDev session to unregister its hotkeys.
TypeScript editor declarations live in `VD.d.mts`.
