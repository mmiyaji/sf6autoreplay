ApplyGuiToVars() {
    global NextDirection, TotalMatches, MaxRunMinutes, RolloverMinutes, RolloverMode
    global ToleranceEnd, gUseFullROI, NextRepeats, NextIntervalMs
    global Key_StartRec, Key_StopRec, Key_ToggleRec, OBSWinSelector, Img_Ends
    global OBSControlMode, OBSWebSocketHost, OBSWebSocketPort, OBSWebSocketPassword, OBSWebSocketTimeoutMs
    global GameWinSelector, AutoRefocusGame, UseOBSRecording, UseOBSToggleForRollover, CheckOnStart_Game, CheckOnStart_OBS
    global CloseGameOnStop, GameExitTimeoutMs
    global LogEnabled, LogDir, AutoScrollLog
    global ResultSnapEnabled, ResultSnapDir
    global SaveOCREnabled, SaveOCRDir
    global SlackEnabled, SlackNotifyMethod, SlackWebhookUrl, SlackRouterUrl, SlackTimeoutMs
    global DiskCheckEnabled, DiskMinFreeGB, DiskCheckPath
    global PauseTimeoutMin

    NextDirection := ddlDir.Text
    TotalMatches := ToIntSafe(edtMatches.Text, TotalMatches)
    MaxRunMinutes := ToIntSafe(edtMaxMin.Text, MaxRunMinutes)
    RolloverMinutes := ToIntSafe(edtRollMin.Text, RolloverMinutes)
    RolloverMode := ddlRollMode.Text
    ToleranceEnd := ToIntSafe(edtTol.Text, ToleranceEnd)
    gUseFullROI := !!chkROI.Value
    NextRepeats := Max(1, ToIntSafe(edtNextRep.Text, NextRepeats))
    NextIntervalMs := Max(50, ToIntSafe(edtNextInt.Text, NextIntervalMs))
    Key_StartRec := edtStart.Text
    Key_StopRec := edtStop.Text
    Key_ToggleRec := edtToggle.Text
    OBSControlMode := ddlObsMode.Text
    OBSWebSocketHost := Trim(edtObsWsHost.Text)
    OBSWebSocketPort := ToIntSafe(edtObsWsPort.Text, OBSWebSocketPort)
    OBSWebSocketPassword := edtObsWsPassword.Text
    OBSWebSocketTimeoutMs := ToIntSafe(edtObsWsTimeout.Text, OBSWebSocketTimeoutMs)
    OBSWinSelector := edtObsSel.Text
    Img_Ends := NormalizeEndImagePaths(edtImgs.Value)
    GameWinSelector := edtGameSel.Text
    AutoRefocusGame := !!chkRefocus.Value
    UseOBSRecording := !!chkUseOBS.Value
    UseOBSToggleForRollover := !!chkUseToggle.Value
    CheckOnStart_Game := !!chkChkGame.Value
    CheckOnStart_OBS := !!chkChkOBS.Value
    CloseGameOnStop := !!chkCloseGame.Value
    PauseTimeoutMin := ToIntSafe(edtPauseTimeout.Value, PauseTimeoutMin)

    LogEnabled := !!chkLog.Value
    LogDir := edtLogDir.Text
    AutoScrollLog := !!chkAutoScroll.Value
    ResultSnapEnabled := !!chkSnap.Value
    ResultSnapDir := edtSnapDir.Text
    SaveOCREnabled := !!chkOCR.Value
    SaveOCRDir := edtOCRDir.Text

    DiskCheckEnabled := !!chkDiskCheckEnabled.Value
    DiskMinFreeGB := ToIntSafe(edtDiskMinFreeGB.Value, DiskMinFreeGB)
    DiskCheckPath := Trim(edtDiskCheckPath.Value)

    SlackEnabled := !!chkSlackEnabled.Value
    SlackNotifyMethod := SlackMethodFromLabel(ddlSlackMethod.Text)
    SlackWebhookUrl := Trim(edtSlackWebhook.Value)
    SlackRouterUrl := Trim(edtSlackRouter.Value)
    SlackTimeoutMs := ToIntSafe(edtSlackTimeout.Value, SlackTimeoutMs)
}

UpdateGuiFromVars() {
    global NextDirection, TotalMatches, MaxRunMinutes, RolloverMinutes, RolloverMode
    global ToleranceEnd, gUseFullROI, NextRepeats, NextIntervalMs
    global Key_StartRec, Key_StopRec, Key_ToggleRec, OBSWinSelector, Img_Ends
    global OBSControlMode, OBSWebSocketHost, OBSWebSocketPort, OBSWebSocketPassword, OBSWebSocketTimeoutMs
    global GameWinSelector, AutoRefocusGame, UseOBSRecording, UseOBSToggleForRollover, CheckOnStart_Game, CheckOnStart_OBS
    global CloseGameOnStop, PauseTimeoutMin
    global LogEnabled, LogDir, AutoScrollLog
    global ResultSnapEnabled, ResultSnapDir
    global SaveOCREnabled, SaveOCRDir
    global SlackEnabled, SlackNotifyMethod, SlackWebhookUrl, SlackRouterUrl, SlackTimeoutMs
    global DiskCheckEnabled, DiskMinFreeGB, DiskCheckPath

    ddlDir.Text := NextDirection
    edtMatches.Text := TotalMatches
    edtMaxMin.Text := MaxRunMinutes
    edtRollMin.Text := RolloverMinutes
    ddlRollMode.Text := RolloverMode
    edtTol.Text := ToleranceEnd
    chkROI.Value := gUseFullROI ? 1 : 0
    edtNextRep.Text := NextRepeats
    edtNextInt.Text := NextIntervalMs
    edtStart.Text := Key_StartRec
    edtStop.Text := Key_StopRec
    edtToggle.Text := Key_ToggleRec
    ddlObsMode.Text := OBSControlMode
    edtObsWsHost.Text := OBSWebSocketHost
    edtObsWsPort.Text := OBSWebSocketPort
    edtObsWsPassword.Text := OBSWebSocketPassword
    edtObsWsTimeout.Text := OBSWebSocketTimeoutMs
    edtObsSel.Text := OBSWinSelector
    edtImgs.Value := BuildEndImageGuiText(Img_Ends)
    edtGameSel.Text := GameWinSelector
    chkRefocus.Value := AutoRefocusGame ? 1 : 0
    chkUseOBS.Value := UseOBSRecording ? 1 : 0
    chkChkGame.Value := CheckOnStart_Game ? 1 : 0
    chkChkOBS.Value := CheckOnStart_OBS ? 1 : 0
    chkUseToggle.Value := UseOBSToggleForRollover ? 1 : 0
    chkCloseGame.Value := CloseGameOnStop ? 1 : 0
    edtPauseTimeout.Value := PauseTimeoutMin

    chkLog.Value := LogEnabled ? 1 : 0
    edtLogDir.Text := LogDir
    chkSnap.Value := ResultSnapEnabled ? 1 : 0
    edtSnapDir.Text := ResultSnapDir
    chkOCR.Value := SaveOCREnabled ? 1 : 0
    edtOCRDir.Text := SaveOCRDir
    chkAutoScroll.Value := AutoScrollLog ? 1 : 0

    chkDiskCheckEnabled.Value := DiskCheckEnabled ? 1 : 0
    edtDiskMinFreeGB.Value := DiskMinFreeGB
    edtDiskCheckPath.Value := DiskCheckPath

    chkSlackEnabled.Value := SlackEnabled ? 1 : 0
    ddlSlackMethod.Text := SlackMethodToLabel(SlackNotifyMethod)
    edtSlackWebhook.Value := SlackWebhookUrl
    edtSlackRouter.Value := SlackRouterUrl
    edtSlackTimeout.Value := SlackTimeoutMs

    UpdateOutputControlStates()
    UpdateSlackUIState()
    UpdatePauseBtn()
}

LoadConfig(path) {
    global NextDirection, TotalMatches, MaxRunMinutes, RolloverMinutes, RolloverMode
    global ToleranceEnd, gUseFullROI, NextRepeats, NextIntervalMs
    global Key_StartRec, Key_StopRec, Key_ToggleRec, OBSWinSelector, GameWinSelector, AutoRefocusGame
    global OBSControlMode, OBSWebSocketHost, OBSWebSocketPort, OBSWebSocketPassword, OBSWebSocketTimeoutMs
    global Img_Ends, UseOBSRecording, UseOBSToggleForRollover, CheckOnStart_Game, CheckOnStart_OBS
    global CloseGameOnStop, GameExitTimeoutMs
    global LogEnabled, LogDir, AutoScrollLog
    global ResultSnapEnabled, ResultSnapDir
    global SaveOCREnabled, SaveOCRDir
    global SlackEnabled, SlackNotifyMethod, SlackWebhookUrl, SlackRouterUrl, SlackTimeoutMs
    global DiskCheckEnabled, DiskMinFreeGB, DiskCheckPath
    global PauseTimeoutMin

    NextDirection := IniRead(path, "main", "NextDirection", NextDirection)
    TotalMatches := Integer(IniRead(path, "main", "TotalMatches", TotalMatches))
    MaxRunMinutes := Integer(IniRead(path, "main", "MaxRunMinutes", MaxRunMinutes))
    RolloverMinutes := Integer(IniRead(path, "main", "RolloverMinutes", RolloverMinutes))
    RolloverMode := IniRead(path, "main", "RolloverMode", RolloverMode)
    ToleranceEnd := Integer(IniRead(path, "main", "ToleranceEnd", ToleranceEnd))
    gUseFullROI := (Integer(IniRead(path, "main", "UseFullROI", gUseFullROI ? 1 : 0)) = 1)
    NextRepeats := Integer(IniRead(path, "main", "NextRepeats", NextRepeats))
    NextIntervalMs := Integer(IniRead(path, "main", "NextIntervalMs", NextIntervalMs))
    PauseTimeoutMin := Integer(IniRead(path, "main", "PauseTimeoutMin", PauseTimeoutMin))

    Key_StartRec := IniRead(path, "obs", "StartKey", Key_StartRec)
    Key_StopRec := IniRead(path, "obs", "StopKey", Key_StopRec)
    Key_ToggleRec := IniRead(path, "obs", "ToggleKey", Key_ToggleRec)
    OBSControlMode := IniRead(path, "obs", "ControlMode", OBSControlMode)
    if (OBSControlMode != "hotkey" && OBSControlMode != "api")
        OBSControlMode := "hotkey"
    OBSWebSocketHost := IniRead(path, "obs", "WebSocketHost", OBSWebSocketHost)
    OBSWebSocketPort := Integer(IniRead(path, "obs", "WebSocketPort", OBSWebSocketPort))
    OBSWebSocketPassword := IniRead(path, "obs", "WebSocketPassword", OBSWebSocketPassword)
    OBSWebSocketTimeoutMs := Integer(IniRead(path, "obs", "WebSocketTimeoutMs", OBSWebSocketTimeoutMs))
    OBSWinSelector := IniRead(path, "obs", "WindowSelector", OBSWinSelector)
    UseOBSRecording := (Integer(IniRead(path, "obs", "UseRecording", UseOBSRecording ? 1 : 0)) = 1)
    UseOBSToggleForRollover := (Integer(IniRead(path, "obs", "UseToggleForRollover", UseOBSToggleForRollover ? 1 : 0)) = 1)
    CheckOnStart_OBS := (Integer(IniRead(path, "obs", "CheckOnStart", CheckOnStart_OBS ? 1 : 0)) = 1)

    GameWinSelector := IniRead(path, "game", "WindowSelector", GameWinSelector)
    AutoRefocusGame := (Integer(IniRead(path, "game", "AutoRefocus", AutoRefocusGame ? 1 : 0)) = 1)
    CheckOnStart_Game := (Integer(IniRead(path, "game", "CheckOnStart", CheckOnStart_Game ? 1 : 0)) = 1)
    CloseGameOnStop := (Integer(IniRead(path, "game", "CloseOnStop", CloseGameOnStop ? 1 : 0)) = 1)
    GameExitTimeoutMs := Integer(IniRead(path, "game", "ExitTimeoutMs", GameExitTimeoutMs))

    imgs := IniRead(path, "images", "EndImages", "")
    if (imgs != "")
        Img_Ends := SplitList(imgs, ";")

    LogEnabled := (Integer(IniRead(path, "log", "Enabled", LogEnabled ? 1 : 0)) = 1)
    LogDir := IniRead(path, "log", "Dir", LogDir)
    ResultSnapEnabled := (Integer(IniRead(path, "log", "ResultSnapEnabled", ResultSnapEnabled ? 1 : 0)) = 1)
    ResultSnapDir := IniRead(path, "log", "ResultSnapDir", ResultSnapDir)
    AutoScrollLog := (Integer(IniRead(path, "log", "AutoScroll", AutoScrollLog ? 1 : 0)) = 1)

    SaveOCREnabled := (Integer(IniRead(path, "ocr", "SaveOCREnabled", SaveOCREnabled ? 1 : 0)) = 1)
    SaveOCRDir := IniRead(path, "ocr", "SaveOCRDir", SaveOCRDir)

    SlackEnabled := (Integer(IniRead(path, "slack", "Enabled", SlackEnabled ? 1 : 0)) = 1)
    SlackNotifyMethod := IniRead(path, "slack", "Method", SlackNotifyMethod)
    SlackWebhookUrl := IniRead(path, "slack", "WebhookUrl", SlackWebhookUrl)
    SlackRouterUrl := IniRead(path, "slack", "RouterUrl", SlackRouterUrl)
    SlackTimeoutMs := Integer(IniRead(path, "slack", "TimeoutMs", SlackTimeoutMs))
    if (SlackNotifyMethod != "webhook" && SlackNotifyMethod != "router")
        SlackNotifyMethod := (SlackWebhookUrl != "") ? "webhook" : ((SlackRouterUrl != "") ? "router" : "webhook")

    DiskCheckEnabled := (Integer(IniRead(path, "disk", "CheckEnabled", DiskCheckEnabled ? 1 : 0)) = 1)
    DiskMinFreeGB := Integer(IniRead(path, "disk", "MinFreeGB", DiskMinFreeGB))
    DiskCheckPath := IniRead(path, "disk", "CheckPath", DiskCheckPath)
}

SaveConfig(path) {
    global NextDirection, TotalMatches, MaxRunMinutes, RolloverMinutes, RolloverMode
    global ToleranceEnd, gUseFullROI, NextRepeats, NextIntervalMs
    global Key_StartRec, Key_StopRec, Key_ToggleRec, OBSWinSelector, GameWinSelector, AutoRefocusGame
    global OBSControlMode, OBSWebSocketHost, OBSWebSocketPort, OBSWebSocketPassword, OBSWebSocketTimeoutMs
    global Img_Ends, UseOBSRecording, UseOBSToggleForRollover, CheckOnStart_Game, CheckOnStart_OBS
    global CloseGameOnStop, GameExitTimeoutMs
    global LogEnabled, LogDir, AutoScrollLog
    global ResultSnapEnabled, ResultSnapDir
    global SaveOCREnabled, SaveOCRDir
    global SlackEnabled, SlackNotifyMethod, SlackWebhookUrl, SlackRouterUrl, SlackTimeoutMs
    global DiskCheckEnabled, DiskMinFreeGB, DiskCheckPath
    global PauseTimeoutMin

    ApplyGuiToVars()

    IniWrite(NextDirection, path, "main", "NextDirection")
    IniWrite(TotalMatches, path, "main", "TotalMatches")
    IniWrite(MaxRunMinutes, path, "main", "MaxRunMinutes")
    IniWrite(RolloverMinutes, path, "main", "RolloverMinutes")
    IniWrite(RolloverMode, path, "main", "RolloverMode")
    IniWrite(ToleranceEnd, path, "main", "ToleranceEnd")
    IniWrite(gUseFullROI ? 1 : 0, path, "main", "UseFullROI")
    IniWrite(NextRepeats, path, "main", "NextRepeats")
    IniWrite(NextIntervalMs, path, "main", "NextIntervalMs")
    IniWrite(PauseTimeoutMin, path, "main", "PauseTimeoutMin")

    IniWrite(Key_StartRec, path, "obs", "StartKey")
    IniWrite(Key_StopRec, path, "obs", "StopKey")
    IniWrite(Key_ToggleRec, path, "obs", "ToggleKey")
    IniWrite(OBSControlMode, path, "obs", "ControlMode")
    IniWrite(OBSWebSocketHost, path, "obs", "WebSocketHost")
    IniWrite(OBSWebSocketPort, path, "obs", "WebSocketPort")
    IniWrite(OBSWebSocketPassword, path, "obs", "WebSocketPassword")
    IniWrite(OBSWebSocketTimeoutMs, path, "obs", "WebSocketTimeoutMs")
    IniWrite(OBSWinSelector, path, "obs", "WindowSelector")
    IniWrite(UseOBSRecording ? 1 : 0, path, "obs", "UseRecording")
    IniWrite(UseOBSToggleForRollover ? 1 : 0, path, "obs", "UseToggleForRollover")
    IniWrite(CheckOnStart_OBS ? 1 : 0, path, "obs", "CheckOnStart")

    IniWrite(GameWinSelector, path, "game", "WindowSelector")
    IniWrite(AutoRefocusGame ? 1 : 0, path, "game", "AutoRefocus")
    IniWrite(CheckOnStart_Game ? 1 : 0, path, "game", "CheckOnStart")
    IniWrite(CloseGameOnStop ? 1 : 0, path, "game", "CloseOnStop")
    IniWrite(GameExitTimeoutMs, path, "game", "ExitTimeoutMs")

    IniWrite(JoinList(Img_Ends, ";"), path, "images", "EndImages")

    IniWrite(LogEnabled ? 1 : 0, path, "log", "Enabled")
    IniWrite(LogDir, path, "log", "Dir")
    IniWrite(ResultSnapEnabled ? 1 : 0, path, "log", "ResultSnapEnabled")
    IniWrite(ResultSnapDir, path, "log", "ResultSnapDir")
    IniWrite(AutoScrollLog ? 1 : 0, path, "log", "AutoScroll")

    IniWrite(SaveOCREnabled ? 1 : 0, path, "ocr", "SaveOCREnabled")
    IniWrite(SaveOCRDir, path, "ocr", "SaveOCRDir")

    IniWrite(SlackEnabled ? 1 : 0, path, "slack", "Enabled")
    IniWrite(SlackNotifyMethod, path, "slack", "Method")
    IniWrite(SlackWebhookUrl, path, "slack", "WebhookUrl")
    IniWrite(SlackRouterUrl, path, "slack", "RouterUrl")
    IniWrite(SlackTimeoutMs, path, "slack", "TimeoutMs")

    IniWrite(DiskCheckEnabled ? 1 : 0, path, "disk", "CheckEnabled")
    IniWrite(DiskMinFreeGB, path, "disk", "MinFreeGB")
    IniWrite(DiskCheckPath, path, "disk", "CheckPath")
}

NormalizeEndImagePaths(text) {
    paths := []
    for line in StrSplit(text, ["`r", "`n"], true) {
        line := Trim(line)
        if (line = "")
            continue
        if (SubStr(line, 1, 1) = "\" || InStr(line, ":\"))
            paths.Push(line)
        else
            paths.Push(A_ScriptDir "\" line)
    }
    return paths
}

BuildEndImageGuiText(paths) {
    rels := []
    for img in paths
        rels.Push(StrReplace(img, A_ScriptDir "\"))
    return StrJoin(rels, "`n")
}

UpdateOutputControlStates() {
    global chkLog, edtLogDir, btnLogDir
    global chkSnap, edtSnapDir, btnSnapDir
    global chkOCR, edtOCRDir, btnOCRDir

    if IsSet(chkLog) && IsSet(edtLogDir) && IsSet(btnLogDir) {
        edtLogDir.Enabled := (chkLog.Value = 1)
        btnLogDir.Enabled := (chkLog.Value = 1)
    }
    if IsSet(chkSnap) && IsSet(edtSnapDir) && IsSet(btnSnapDir) {
        edtSnapDir.Enabled := (chkSnap.Value = 1)
        btnSnapDir.Enabled := (chkSnap.Value = 1)
    }
    if IsSet(chkOCR) && IsSet(edtOCRDir) && IsSet(btnOCRDir) {
        edtOCRDir.Enabled := (chkOCR.Value = 1)
        btnOCRDir.Enabled := (chkOCR.Value = 1)
    }
}
