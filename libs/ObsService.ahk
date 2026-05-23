FocusedTriggerOBS(keyToSend) {
    global OBSWinSelector, UseOBSRecording, AutoRefocusGame
    if !UseOBSRecording {
        Log("OBS: key send skipped (UseOBSRecording=false)")
        return false
    }
    prev := WinExist("A")
    if WinExist(OBSWinSelector) {
        WinActivate OBSWinSelector
        if !WinWaitActive(OBSWinSelector, , 0.5)
            Log("WARN: OBS activate timed out before key [" keyToSend "]")
        Sleep 100
        Send keyToSend
        Sleep 60
        if AutoRefocusGame
            EnsureFocusGame()
        else if prev
            WinActivate prev
        Log("OBS: key sent [" keyToSend "]")
        return true
    } else {
        TrayTip "OBS not found", OBSWinSelector " not found", 1200
        Log("ERROR: OBS window not found for key [" keyToSend "]")
        return false
    }
}

OBSStartRecording() {
    global Key_StartRec
    return OBSControl("start", Key_StartRec)
}

OBSStopRecording() {
    global Key_StopRec
    return OBSControl("stop", Key_StopRec)
}

OBSSplitRecording() {
    global Key_ToggleRec
    return OBSControl("split", Key_ToggleRec)
}

OBSGetRecordStatus() {
    return OBSControl("status", "")
}

OBSControl(action, fallbackKey := "") {
    global OBSControlMode

    if (StrLower(OBSControlMode) = "api")
        return OBSWebSocketRequest(action)

    if (fallbackKey != "")
        return FocusedTriggerOBS(fallbackKey)

    return false
}

OBSWebSocketRequest(action) {
    global OBSWebSocketHost, OBSWebSocketPort, OBSWebSocketPassword, OBSWebSocketTimeoutMs

    if (OBSWebSocketHost = "")
        OBSWebSocketHost := "127.0.0.1"
    if (OBSWebSocketPort <= 0)
        OBSWebSocketPort := 4455
    if (OBSWebSocketTimeoutMs <= 0)
        OBSWebSocketTimeoutMs := 5000

    script := A_ScriptDir "\libs\obs_ws.ps1"
    if !FileExist(script) {
        Log("ERROR: OBS API script not found: " script)
        return false
    }

    outPath := A_Temp "\sf6_obs_ws_" A_TickCount ".json"
    cmd := PSCmdQuote(OBSFindPowerShell()) " -NoProfile -ExecutionPolicy Bypass -File "
        . PSCmdQuote(script)
        . " -HostName " PSCmdQuote(OBSWebSocketHost)
        . " -Port " OBSWebSocketPort
        . " -Password " PSCmdQuote(OBSWebSocketPassword)
        . " -Action " PSCmdQuote(action)
        . " -TimeoutMs " OBSWebSocketTimeoutMs
        . " > " PSCmdQuote(outPath) " 2>&1"

    exitCode := RunWait(cmd, , "Hide")
    output := ""
    try output := Trim(FileRead(outPath, "UTF-8"))
    try FileDelete(outPath)

    if (exitCode = 0) {
        Log("OBS API: " action " ok " output)
        return true
    }

    Log("ERROR: OBS API " action " failed exit=" exitCode " output=" output)
    TrayTip "OBS API error", action " failed. See log.", 1600
    return false
}

OBSFindPowerShell() {
    p1 := A_WinDir "\System32\WindowsPowerShell\v1.0\powershell.exe"
    return FileExist(p1) ? p1 : "pwsh.exe"
}

PSCmdQuote(s) {
    q := Chr(34)
    return q StrReplace(s, q, q q) q
}

RolloverOBS(mode := "instant") {
    global UseOBSRecording, UseOBSToggleForRollover, Key_ToggleRec
    global gLastRolloverTick, OBSControlMode
    if !UseOBSRecording
        return

    ok := false
    if (StrLower(OBSControlMode) = "api") {
        Log("OBS: rollover via API split (" mode ")")
        ok := OBSSplitRecording()
        Sleep 800
    } else if (UseOBSToggleForRollover && Key_ToggleRec != "") {
        Log("OBS: rollover via toggle(one-shot) (" mode ")")
        ok := OBSSplitRecording()
        Sleep 800
    } else {
        Log("OBS: rollover via stop/start (" mode ")")
        okStop := OBSStopRecording()
        Sleep 900
        okStart := OBSStartRecording()
        ok := okStop && okStart
    }

    if !ok {
        Log("ERROR: OBS rollover failed; text file was not switched")
        TrayTip "Rollover failed", "OBS did not confirm file switch", 1600
        return
    }

    gLastRolloverTick := A_TickCount
    StartNewRecordingTextFile("rollover")
    TrayTip "Rollover", "Recording file switched (" mode ")", 1200
}

SendOBSTest() {
    global UseOBSRecording

    if !UseOBSRecording {
        TrayTip "OBS disabled", "OBS recording is disabled in settings", 1200
        Log("TEST: OBS test skipped (UseOBSRecording=false)")
        return
    }
    TrayTip "Test", "OBS start", 700
    OBSStartRecording()
    Sleep 700
    TrayTip "Test", "OBS stop", 700
    OBSStopRecording()
}

StartNewRecordingTextFile(reason := "start") {
    global MatchTextDir, gCurrentTextPath, gRecStartTick
    try DirCreate(MatchTextDir)
    ts := FormatTime(A_Now, "yyyyMMdd_HHmmss")
    gCurrentTextPath := MatchTextDir "\sf6_" ts ".txt"
    gRecStartTick := A_TickCount
    Log("TEXT: new output file -> " gCurrentTextPath " [" reason "]")
}
