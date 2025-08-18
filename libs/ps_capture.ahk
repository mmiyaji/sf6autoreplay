; ps_capture.ahk  — PowerShell-based screen capture helpers (AutoHotkey v2)
; 依存: PowerShell + .NET (System.Drawing)。追加ツール不要。
; 実行は -EncodedCommand を使うのでクォート問題なし。ExecutionPolicy はプロセス限定 Bypass。

#Requires AutoHotkey v2.0

class PSCap {
    static _dpiSet := false

    ; ----------------- Public API -----------------
    ; すべて成功時は outPath（絶対パス）を返し、失敗時は例外を投げます。
    static captureRect(x, y, w, h, outPath) {
        if (w <= 0 || h <= 0)
            throw Error("PSCap.captureRect: w/h must be > 0")
        if (!outPath)
            throw Error("PSCap.captureRect: outPath is required")
        this._ensureParentDir(outPath)
        this._ensureDpiAware()

        ps := this._psTemplate(x, y, w, h, this._psSaveLine(outPath))
        this._psRun(ps)
        return this._full(outPath)
    }

    static captureWindowClient(hwnd, outPath) {
        if !hwnd
            throw Error("PSCap.captureWindowClient: invalid hwnd")
        if (!outPath)
            throw Error("PSCap.captureWindowClient: outPath is required")
        this._ensureParentDir(outPath)
        this._ensureDpiAware()

        r := this._clientRectOnScreen(hwnd)
        if (r.w <= 0 || r.h <= 0)
            throw Error(Format("PSCap.captureWindowClient: invalid rect x={},y={},w={},h={}", r.x, r.y, r.w, r.h))

        ps := this._psTemplate(r.x, r.y, r.w, r.h, this._psSaveLine(outPath))
        this._psRun(ps)
        return this._full(outPath)
    }

    static captureWindowNoShadow(hwnd, outPath) {
        if !hwnd
            throw Error("PSCap.captureWindowNoShadow: invalid hwnd")
        if (!outPath)
            throw Error("PSCap.captureWindowNoShadow: outPath is required")
        this._ensureParentDir(outPath)
        this._ensureDpiAware()

        r := this._dwmExtendedFrame(hwnd)   ; タイトルバー含む／影なし
        if (r.w <= 0 || r.h <= 0)
            r := this._windowRect(hwnd)     ; フォールバック（影あり）

        ps := this._psTemplate(r.x, r.y, r.w, r.h, this._psSaveLine(outPath))
        this._psRun(ps)
        return this._full(outPath)
    }

    static captureWindowWithShadow(hwnd, outPath) {
        if !hwnd
            throw Error("PSCap.captureWindowWithShadow: invalid hwnd")
        if (!outPath)
            throw Error("PSCap.captureWindowWithShadow: outPath is required")
        this._ensureParentDir(outPath)
        this._ensureDpiAware()

        r := this._windowRect(hwnd)         ; 影あり
        ps := this._psTemplate(r.x, r.y, r.w, r.h, this._psSaveLine(outPath))
        this._psRun(ps)
        return this._full(outPath)
    }

    static capturePrimary(outPath) {
        if (!outPath)
            throw Error("PSCap.capturePrimary: outPath is required")
        this._ensureParentDir(outPath)
        this._ensureDpiAware()

        r := this._primaryScreenRect()
        ps := this._psTemplate(r.x, r.y, r.w, r.h, this._psSaveLine(outPath))
        this._psRun(ps)
        return this._full(outPath)
    }

    static captureVirtual(outPath) {
        if (!outPath)
            throw Error("PSCap.captureVirtual: outPath is required")
        this._ensureParentDir(outPath)
        this._ensureDpiAware()

        r := this._virtualScreenRect()
        ps := this._psTemplate(r.x, r.y, r.w, r.h, this._psSaveLine(outPath))
        this._psRun(ps)
        return this._full(outPath)
    }

    ; ----------------- Internals -----------------
    static _psTemplate(x, y, w, h, saveLine) {
        ; 改行は PowerShell の `n を明示
        ps := "$ErrorActionPreference='Stop';`n"
        ps .= "Add-Type -AssemblyName System.Drawing;`n"
        ps .= "$bmp=New-Object System.Drawing.Bitmap(" . w . "," . h . ");`n"
        ps .= "$g=[System.Drawing.Graphics]::FromImage($bmp);`n"
        ps .= "$g.CopyFromScreen(" . x . "," . y . ",0,0,$bmp.Size);`n"
        ps .= saveLine . "`n"
        ps .= "$g.Dispose();$bmp.Dispose();"
        return ps
    }

    static _psSaveLine(path) {
        ext := StrLower(RegExReplace(path, ".*\.", ""))
        p := this._psSingle(path)  ; 'C:\path\file.png' のように単一引用符で包む
        switch ext {
            case "jpg","jpeg":
                return "$bmp.Save(" p ",[System.Drawing.Imaging.ImageFormat]::Jpeg);"
            case "bmp":
                return "$bmp.Save(" p ",[System.Drawing.Imaging.ImageFormat]::Bmp);"
            case "gif":
                return "$bmp.Save(" p ",[System.Drawing.Imaging.ImageFormat]::Gif);"
            default:
                return "$bmp.Save(" p ",[System.Drawing.Imaging.ImageFormat]::Png);"
        }
    }

    static _psRun(psScript) {
        ; 文字列のクォート問題を避けるため、UTF-16LE Base64 の -EncodedCommand を使用
        pwsh := this._findPowerShell()
        enc  := this._psEncodeBase64(psScript)
        cmd  := pwsh . " -NoProfile -ExecutionPolicy Bypass -EncodedCommand " . enc
        exitCode := RunWait(cmd, , "Hide")
        if (exitCode != 0)
            throw Error(Format("PSCap: PowerShell failed (exit={})", exitCode))
    }

    static _psEncodeBase64(s) {
        ; UTF-16LE の生バイト長（終端ヌル除く）
        cb := StrLen(s) * 2
        if (cb = 0)
            return ""
        flags := 0x40000001  ; CRYPT_STRING_NOCRLF(0x40000000) | CRYPT_STRING_BASE64(0x1)
        cch := 0
        DllCall("Crypt32\CryptBinaryToStringW", "ptr", StrPtr(s), "uint", cb, "uint", flags, "ptr", 0, "uint*", &cch)
        buf := Buffer(cch * 2, 0)
        if !DllCall("Crypt32\CryptBinaryToStringW", "ptr", StrPtr(s), "uint", cb, "uint", flags, "ptr", buf.Ptr, "uint*", &cch)
            throw Error("PSCap: Base64 encoding failed")
        return StrGet(buf, "UTF-16")
    }

    static _findPowerShell() {
        p1 := A_WinDir "\System32\WindowsPowerShell\v1.0\powershell.exe"
        return FileExist(p1) ? ('"' . p1 . '"') : "pwsh.exe"
    }

    static _psSingle(s) {
        ; PowerShell 単一引用符リテラル。内部の ' は '' にエスケープ
        return "'" . StrReplace(s, "'", "''") . "'"
    }

    static _ensureDpiAware() {
        if this._dpiSet
            return
        if !DllCall("user32\SetProcessDpiAwarenessContext", "ptr", -4, "ptr")
            DllCall("user32\SetProcessDPIAware")
        this._dpiSet := true
    }

    static _ensureParentDir(path) {
        SplitPath path, , &dir
        if (dir && !DirExist(dir))
            DirCreate(dir)
    }

    ; ---- Rect helpers ----
    static _windowRect(hwnd) {
        rect := Buffer(16, 0)
        if !DllCall("user32\GetWindowRect", "ptr", hwnd, "ptr", rect.Ptr, "int")
            throw Error("GetWindowRect failed")
        left   := NumGet(rect,  0, "Int")
        top    := NumGet(rect,  4, "Int")
        right  := NumGet(rect,  8, "Int")
        bottom := NumGet(rect, 12, "Int")
        return { x: left, y: top, w: right - left, h: bottom - top }
    }

    static _dwmExtendedFrame(hwnd) {
        static DWMWA_EXTENDED_FRAME_BOUNDS := 9
        rect := Buffer(16, 0)
        hr := DllCall("dwmapi\DwmGetWindowAttribute"
                    , "ptr", hwnd
                    , "int", DWMWA_EXTENDED_FRAME_BOUNDS
                    , "ptr", rect.Ptr
                    , "int", rect.Size
                    , "int")
        if (hr != 0)
            return { x: 0, y: 0, w: 0, h: 0 }  ; 失敗時は無効矩形を返す
        left   := NumGet(rect,  0, "Int")
        top    := NumGet(rect,  4, "Int")
        right  := NumGet(rect,  8, "Int")
        bottom := NumGet(rect, 12, "Int")
        return { x: left, y: top, w: right - left, h: bottom - top }
    }

    static _clientRectOnScreen(hwnd) {
        rect := Buffer(16, 0)
        if !DllCall("user32\GetClientRect", "ptr", hwnd, "ptr", rect.Ptr, "int")
            throw Error("GetClientRect failed")
        DllCall("user32\MapWindowPoints", "ptr", hwnd, "ptr", 0, "ptr", rect.Ptr, "uint", 2, "int")
        left   := NumGet(rect,  0, "Int")
        top    := NumGet(rect,  4, "Int")
        right  := NumGet(rect,  8, "Int")
        bottom := NumGet(rect, 12, "Int")
        return { x: left, y: top, w: right - left, h: bottom - top }
    }

    static _primaryScreenRect() {
        hMon := DllCall("user32\MonitorFromPoint", "int64", 0, "uint", 1, "ptr")  ; MONITOR_DEFAULTTOPRIMARY=1
        return this._monitorRect(hMon)
    }

    static _virtualScreenRect() {
        x := DllCall("user32\GetSystemMetrics", "int", 76, "int")  ; SM_XVIRTUALSCREEN
        y := DllCall("user32\GetSystemMetrics", "int", 77, "int")  ; SM_YVIRTUALSCREEN
        w := DllCall("user32\GetSystemMetrics", "int", 78, "int")  ; SM_CXVIRTUALSCREEN
        h := DllCall("user32\GetSystemMetrics", "int", 79, "int")  ; SM_CYVIRTUALSCREEN
        return { x: x, y: y, w: w, h: h }
    }

    static _monitorRect(hMon) {
        if !hMon
            throw Error("MonitorFromPoint/Window failed")
        mi := Buffer(40 + 4*4, 0) ; MONITORINFO
        NumPut("UInt", mi.Size, mi, 0)
        if !DllCall("user32\GetMonitorInfoW", "ptr", hMon, "ptr", mi.Ptr, "int")
            throw Error("GetMonitorInfo failed")
        left   := NumGet(mi, 4,  "Int")
        top    := NumGet(mi, 8,  "Int")
        right  := NumGet(mi, 12, "Int")
        bottom := NumGet(mi, 16, "Int")
        return { x: left, y: top, w: right - left, h: bottom - top }
    }

    static _full(p) {
        ; 相対 → 絶対（必要なら）
        return InStr(p, ":") ? p : A_WorkingDir "\" p
    }
}
