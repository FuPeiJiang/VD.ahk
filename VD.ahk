#Include %A_LineFile%\..\_VD.ahk
dummyFunction1() {
    static dummyStatic1 := VD.init()
}
; in YOUR app
; #Include VD.ahk
; or
; #Include _VD.ahk
; ...{startup code}
; VD.init()

; VD.ahk : calls `VD.init()` on #Include
; _VD.ahk : `VD.init()` when you want, like after a GUI has rendered, for startup performance reasons
