FocusedTriggerOBS(keyToSend) {
    global OBSWinSelector, UseOBSRecording, AutoRefocusGame
    if !UseOBSRecording {
        Log("OBS: key send skipped (UseOBSRecording=false)")
        return
    }
    prev := WinExist("A")
    if WinExist(OBSWinSelector) {
        WinActivate OBSWinSelector
        WinWaitActive OBSWinSelector, , 500
        Sleep 100
        Send keyToSend
        Sleep 60
        if AutoRefocusGame
            EnsureFocusGame()
        else if prev
            WinActivate prev
        Log("OBS: key sent [" keyToSend "]")
    } else {
        TrayTip "OBS未検出", OBSWinSelector " が見つかりません", 1200
        Log("ERROR: OBS window not found for key [" keyToSend "]")
    }
}

RolloverOBS(mode := "instant") {
    global UseOBSRecording, UseOBSToggleForRollover, Key_ToggleRec
    global Key_StopRec, Key_StartRec, gLastRolloverTick
    if !UseOBSRecording
        return

    if (UseOBSToggleForRollover && Key_ToggleRec != "") {
        Log("OBS: rollover via toggle(one-shot) (" mode ")")
        FocusedTriggerOBS(Key_ToggleRec)
        Sleep 800
    } else {
        Log("OBS: rollover via stop/start (" mode ")")
        FocusedTriggerOBS(Key_StopRec)
        Sleep 900
        FocusedTriggerOBS(Key_StartRec)
    }

    gLastRolloverTick := A_TickCount
    StartNewRecordingTextFile("rollover")
    TrayTip "ローテ", "録画ファイルを切替（" (mode="instant"?"即時":"試合間") "）", 1200
}

SendOBSTest() {
    if !UseOBSRecording {
        TrayTip "OBS未使用", "設定でOBS録画がOFFです", 1200
        Log("TEST: OBS test skipped (UseOBSRecording=false)")
        return
    }
    TrayTip "テスト", "録画開始キー送信", 700
    FocusedTriggerOBS(Key_StartRec)
    Sleep 700
    TrayTip "テスト", "録画停止キー送信", 700
    FocusedTriggerOBS(Key_StopRec)
}

StartNewRecordingTextFile(reason := "start") {
    global MatchTextDir, gCurrentTextPath, gRecStartTick
    try DirCreate(MatchTextDir)
    ts := FormatTime(A_Now, "yyyyMMdd_HHmmss")
    gCurrentTextPath := MatchTextDir "\sf6_" ts ".txt"
    gRecStartTick := A_TickCount
    Log("TEXT: new output file -> " gCurrentTextPath " [" reason "]")
}