#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance force
ListLines Off
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetBatchLines -1
#KeyHistory 0


Gui, Add, Text, hwndHWND, fjiewjf
Gui, Add, Text, hwndHWND2, gjwioehjgi
Gui, Show, w1000 h100

what()
what() {
global
ok1:=ErrorLevel
ok2:=A_LastError
hdc:=DllCall("GetDC", "Ptr",HWND, "Ptr")
; hdc:=DllCall("GetWindowDC", "Ptr",HWND, "Ptr")
ret:=DllCall("Gdi32\SetBkColor", "Ptr",hdc, "Uint",0x0000FF00)
; ret:=DllCall("Gdi32\SetBkColor", "Ptr",hdc, "Uint",0x000000FF)
hBrush:=DllCall("Gdi32\CreateSolidBrush", "Uint",0x0000FF00, "Ptr")
ret:=DllCall("Gdi32\SelectObject", "Ptr",hdc, "Ptr",hBrush, "Ptr")
; DllCall("CreateRectRgn", "int",x1, "int",y1, "int",x2, "int",y2, "Ptr")
Rgn:=DllCall("Gdi32\CreateRectRgn", "int",0, "int",0, "int",10, "int",100, "Ptr")
; ret:=DllCall("GetWindowRgn", "Ptr",HWND, "Ptr",Rgn)
; ret:=DllCall("Gdi32\SelectObject", "Ptr",hdc, "Ptr",Rgn, "Ptr")
; ret:=DllCall("Gdi32\FillRgn", "Ptr",hdc, "Ptr",Rgn, "Ptr",hBrush)

VarSetCapacity(Rect, 16)  ; A RECT is a struct consisting of four 32-bit integers (i.e. 4*4=16).
DllCall("GetClientRect", "Ptr", HWND, "Ptr", &Rect)  ; WinExist() returns an HWND.


HFONT:=DllCall("SendMessage", "Ptr",hwnd, "UInt",WM_GETFONT:=0x0031, "Ptr",0, "Ptr",0)
ret:=DllCall("Gdi32\SelectObject", "Ptr",hdc, "Ptr",HFONT, "Ptr")

NumPut(NumGet(Rect, 0, "Int") - 2, Rect, 0, "Int")
NumPut(NumGet(Rect, 4, "Int") - 2, Rect, 4, "Int")
NumPut(NumGet(Rect, 8, "Int") + 2, Rect, 8, "Int")
NumPut(NumGet(Rect, 12, "Int") + 2, Rect, 12, "Int")

ret:=DllCall("FillRect", "Ptr",hdc, "Ptr",&Rect, "Ptr",hBrush)

NumPut(NumGet(Rect, 0, "Int") + 2, Rect, 0, "Int")
NumPut(NumGet(Rect, 4, "Int") + 2, Rect, 4, "Int")
NumPut(NumGet(Rect, 8, "Int") - 2, Rect, 8, "Int")
NumPut(NumGet(Rect, 12, "Int") - 2, Rect, 12, "Int")

; MsgBox % "Left " . NumGet(Rect, 0, "Int") . " Top " . NumGet(Rect, 4, "Int") . " Right " . NumGet(Rect, 8, "Int") . " Bottom " . NumGet(Rect, 12, "Int")

ret:=DllCall("DrawText", "Ptr",hdc, "Str",getText(HWND), "int",-1, "Ptr",&Rect, "Uint",0)
ok3:=okLastError()

; DT_CENTER:=0x00000001
; ,DT_NOCLIP:=0x00000100
; DllCall("DrawText", "Ptr",hdc, "Str",getText(HWND), "int",-1, "Ptr",&Rect, "Uint",DT_CENTER | DT_NOCLIP)



; DllCall("User32.dll\InvalidateRect", "Ptr", HWND, "Ptr", 0, "Int", 1)
; DllCall("User32.dll\RedrawWindow", "Ptr", HWND, "Ptr", 0, "Int", 1)
; ret:=DllCall("SendMessage", "Ptr",HWND, "UInt",WM_ERASEBKGND:=0x0014, "Ptr",hdc, "Ptr",0)

; hex:=Format("{:X}", ret)
; ret:=DllCall("SendMessage", "Ptr",HWND, "UInt",WM_ERASEBKGND:=0x0014, "Ptr",hdc, "Ptr",0)
; ret:=DllCall("SendMessage", "Ptr",HWND, "UInt",WM_ERASEBKGND:=0x0014, "Ptr",hdc, "Ptr",0)

}

return

f3::Exitapp

okLastError() {
    FORMAT_MESSAGE_ALLOCATE_BUFFER:=0x00000100
    ,FORMAT_MESSAGE_FROM_SYSTEM:=0x00001000
    ,FORMAT_MESSAGE_IGNORE_INSERTS:=0x00000200

    ret:=DllCall("FormatMessage", "Uint",FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS
    ,"Ptr",0
    ,"Uint",A_LastError
    ,"Uint",0x00000400 ; MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT)
    ,"Ptr*",lpMsgBuf
    ,"Uint",0
    ,"Ptr",0, "Uint")
    Return StrGet(lpMsgBuf)
}
Hex(num) {
    return hex:=Format("0x{:X}", num)
}
getText(hwnd) {
    lengthInCharacters:=DllCall("SendMessage", "Ptr",hwnd, "UInt",WM_GETTEXTLENGTH:=0x000E, "Ptr",0, "Ptr",0)

    VarSetCapacity(buf , 2*lengthInCharacters)
    numberOfCharsCopied:=DllCall("SendMessage", "Ptr",hwnd, "UInt",WM_GETTEXT:=0x000D, "Ptr",lengthInCharacters + 1, "Ptr",&buf)

    return buf
}
