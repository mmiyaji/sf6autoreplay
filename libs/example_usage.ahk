; example_usage.ahk (AutoHotkey v2)
#Include %A_ScriptDir%\ps_capture.ahk

; Capture frontmost window's client area to .\captures\client.png
F1::{
    hwnd := WinExist("A")
    out  := A_ScriptDir "\captures\client.png"
    PSCap.captureWindowClient(hwnd, out)
    MsgBox "Saved: " out
}

; Capture a custom rectangle to .\captures\roi.png
F2::{
    out := A_ScriptDir "\captures\roi.png"
    PSCap.captureRect(100,200,640,360, out)
    MsgBox "Saved: " out
}

; Capture primary monitor to .\captures\primary.png
F3::{
    out := A_ScriptDir "\captures\primary.png"
    PSCap.capturePrimary(out)
    MsgBox "Saved: " out
}

; Capture entire virtual screen (all monitors) to .\captures\all.png
F4::{
    out := A_ScriptDir "\captures\all.png"
    PSCap.captureVirtual(out)
    MsgBox "Saved: " out
}

; Capture with/without window shadow (for the active window)
F5::{
    hwnd := WinExist("A")
    out1 := A_ScriptDir "\captures\with_shadow.png"
    out2 := A_ScriptDir "\captures\no_shadow.png"
    PSCap.captureWindowWithShadow(hwnd, out1)
    PSCap.captureWindowNoShadow(hwnd, out2)
    MsgBox "Saved:`n" out1 "`n" out2
}
