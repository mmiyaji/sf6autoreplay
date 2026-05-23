AutomationRun() {
    global gLoopCount, gSafeStopRequested
    global CloseGameOnStop

    if !AutomationValidateStartPrereqs()
        return

    AutomationBeginRun()

    Loop {
        if !AutomationShouldContinueLoop()
            break

        gLoopCount += 1
        Log("REPLAY: start #" gLoopCount)

        if !AutomationWaitIfPaused()
            break

        AutomationStartReplayPlayback()
        if !AutomationWaitForEndUi()
            break

        if gSafeStopRequested
            break

        AutomationAdvanceToNextReplay()
    }

    stopReason := AutomationResolveStopReason()
    AutomationFinalizeRun(stopReason)

    if CloseGameOnStop {
        Log("CLOSE: option enabled, trying to close GAME...")
        CloseGameApp()
    }

    SlackNotify(BuildSlackEndMessage(stopReason), (stopReason = "safe_stop" ? "warn" : "info"))
    Log("SEND: notify")
}

AutomationValidateStartPrereqs() {
    global gRunning, CheckOnStart_Game, CheckOnStart_OBS, UseOBSRecording
    global DiskCheckEnabled, Img_Ends

    if gRunning {
        TrayTip "実行中", "安全停止は Ctrl+Alt+X（またはGUI）", 1500
        return false
    }

    if !AutomationCheckRequiredWindows()
        return false

    for img in Img_Ends {
        if !FileExist(img) {
            MsgBox "終了判定画像が見つかりません:`n" img, "エラー", 16
            Log("ERROR: End image missing: " img)
            return false
        }
    }

    if (DiskCheckEnabled && UseOBSRecording && !AutomationCheckDiskSpace())
        return false

    return true
}

AutomationCheckRequiredWindows() {
    global CheckOnStart_Game, CheckOnStart_OBS, UseOBSRecording
    global GameWinSelector, OBSWinSelector

    if CheckOnStart_Game && !WinExist(GameWinSelector) {
        msg := "ゲームのウィンドウが見つかりません:`n" GameWinSelector
        MsgBox msg, "起動チェック", 48
        Log("WARN: Game window not found: " GameWinSelector)
        SlackNotify("⚠️ 起動チェック失敗: ゲームウィンドウが見つかりません`n" GameWinSelector, "warning")
        return false
    }

    if UseOBSRecording && CheckOnStart_OBS && !WinExist(OBSWinSelector) {
        msg := "OBSのウィンドウが見つかりません:`n" OBSWinSelector
        MsgBox msg, "起動チェック", 48
        Log("WARN: OBS window not found: " OBSWinSelector)
        SlackNotify("⚠️ 起動チェック失敗: OBSウィンドウが見つかりません`n" OBSWinSelector, "warning")
        return false
    }

    return true
}

AutomationCheckDiskSpace() {
    global DiskCheckPath, DiskMinFreeGB

    freeGB := DriveGetSpaceFree(DiskCheckPath) / 1024
    if freeGB < DiskMinFreeGB {
        msg := "ディスク空き容量が不足しています。`n空き: " Round(freeGB, 1) " GB / 最低: " DiskMinFreeGB " GB`n`n録画を開始しますか？"
        if MsgBox(msg, "ディスク空き容量チェック", 52) != "Yes" {
            Log("ABORT: disk space too low (" Round(freeGB, 1) "GB free)")
            return false
        }
    }

    Log("DISK: " Round(freeGB, 1) "GB free on " DiskCheckPath)
    return true
}

AutomationBeginRun() {
    global gRunning, gPaused, gRecording, gSafeStopRequested, gRunStartTick
    global gRolloverRequested, gLastRolloverTick, gLoopCount, gHardTimeoutCount
    global UseOBSRecording, Key_StartRec

    gRunning := true
    gPaused := false
    gSafeStopRequested := false
    gRolloverRequested := false
    gRunStartTick := A_TickCount
    gLoopCount := 0
    gHardTimeoutCount := 0

    CoordMode "Pixel", "Screen"
    RefocusGame(true)
    Log("START: automation")
    SlackNotify(BuildSlackStartMessage(), "info")

    if UseOBSRecording && !gRecording {
        if !OBSStartRecording() {
            Log("ERROR: OBS start recording failed; aborting automation")
            SlackNotify("OBS start recording failed; automation aborted", "warning")
            gRunning := false
            return
        }
        gRecording := true
        gLastRolloverTick := A_TickCount
        StartNewRecordingTextFile("start")
        TrayTip "録画開始", "通し録画を開始", 1200
        Log("OBS: start recording")
        Sleep 200
    } else {
        gLastRolloverTick := A_TickCount
        StartNewRecordingTextFile("run")
    }
}

AutomationShouldContinueLoop() {
    global gRunning, TotalMatches, gLoopCount, gSafeStopRequested

    if !gRunning
        return false
    if (TotalMatches > 0 && gLoopCount >= TotalMatches)
        return false

    AutomationApplyRunLimits()
    return !gSafeStopRequested
}

AutomationApplyRunLimits() {
    global MaxRunMinutes, gSafeStopRequested, gRunStartTick
    global UseOBSRecording, RolloverMinutes, gLastRolloverTick, RolloverMode, gRolloverRequested

    if (MaxRunMinutes > 0) {
        if !gSafeStopRequested && (A_TickCount - gRunStartTick) > (MaxRunMinutes * 60 * 1000) {
            gSafeStopRequested := true
            TrayTip "安全停止", "タイムリミット到達。次の終了UIで停止します", 1500
            Log("SAFE-STOP: due to MaxRunMinutes")
        }
    }

    if (UseOBSRecording && RolloverMinutes > 0 && !gSafeStopRequested) {
        if (A_TickCount - gLastRolloverTick) > (RolloverMinutes * 60 * 1000) {
            if (RolloverMode = "instant") {
                RolloverOBS("instant")
                SlackNotify("🎥 OBS rollover (instant) at elapsed=" FormatDuration(A_TickCount - gRunStartTick), "info")
            } else {
                gRolloverRequested := true
                TrayTip "ローテ予約", "次の終了UIで録画を切替", 1200
                Log("OBS: rollover requested (safe)")
            }
        }
    }
}

AutomationWaitIfPaused() {
    global gPaused, gRunning, PauseTimeoutMin, gPausedTick, gSafeStopRequested

    while gPaused && gRunning {
        Sleep 150
        if PauseTimeoutMin > 0 && gPausedTick > 0 {
            if (A_TickCount - gPausedTick) > (PauseTimeoutMin * 60000) {
                Log("PAUSE-TIMEOUT: " PauseTimeoutMin "分経過 -> 安全停止")
                SlackNotify("⏸ 一時停止タイムアウト（" PauseTimeoutMin "分）→ 安全停止", "warning")
                gPaused := false
                gSafeStopRequested := true
            }
        }
    }

    return gRunning
}

AutomationStartReplayPlayback() {
    global SaveOCREnabled, GameWinSelector
    global Key_Confirm, Delay_AfterFirstConfirm, Delay_AfterPlayKey

    EnsureFocusGame()
    Press(Key_Confirm, 80)
    Sleep Delay_AfterFirstConfirm
    EnsureFocusGame()

    if SaveOCREnabled {
        try {
            Sleep 150
            OCR_RecordCurrentMatch(GameWinSelector)
        } catch as e {
            Log("OCR: failed - " e.Message)
        }
    } else {
        Log("OCR: skipped (disabled)")
    }

    Press(Key_Confirm, 80)
    Sleep Delay_AfterPlayKey
}

AutomationWaitForEndUi() {
    global gRunning, gSafeStopRequested, gRolloverRequested, gHardTimeoutCount, gLoopCount
    global GameWinSelector, MatchHardTimeoutSec, gUseFullROI, ToleranceEnd, Img_Ends
    global Delay_BeforeNavigate, Delay_AfterBackKey, PollInterval
    global EndDownCount, EndConfirmCount, Key_Down, Key_Confirm
    global EnableLoadBlackWait, BlackDarknessThreshold, BlackMinBlackRatio, BlackGridX, BlackGridY
    global BlackCheckInterval, BlackMinWait, BlackStableAfter, BlackMaxWait
    global UseOBSRecording

    startTick := A_TickCount

    Loop {
        if !gRunning
            return false
        if !AutomationWaitIfPaused()
            return false

        detectHardTimeout := (A_TickCount - startTick) > (MatchHardTimeoutSec * 1000)
        if detectHardTimeout {
            Log("TIMEOUT: no end UI within " MatchHardTimeoutSec "s -> force handling same as detected")
            gHardTimeoutCount += 1
        }

        roi := (gUseFullROI ? GetROI_Full() : GetROI_End_Default(GameWinSelector))
        endUiDetected := FindAnyImage(Img_Ends, roi, ToleranceEnd, &fx, &fy)
        if endUiDetected or detectHardTimeout {
            if endUiDetected
                Log("DETECT: end UI at " fx "," fy)
            else
                Log("DETECT: end UI forced by hard timeout")
            Sleep Delay_BeforeNavigate

            EnsureFocusGame()
            Loop EndDownCount {
                Press(Key_Down, 60)
                Sleep 500
            }
            EnsureFocusGame()
            Loop EndConfirmCount {
                Press(Key_Confirm, 70)
                Sleep 180
            }
            Sleep Delay_AfterBackKey
            WaitImageDisappear(Img_Ends, roi, ToleranceEnd, 1200)

            if (EnableLoadBlackWait) {
                WaitWhileBlackByRatio_Window(GameWinSelector
                    , BlackDarknessThreshold, BlackMinBlackRatio, BlackGridX, BlackGridY
                    , BlackCheckInterval, BlackMinWait, BlackStableAfter, BlackMaxWait
                    , 0.30, 0.30)
            }

            if gRolloverRequested && !gSafeStopRequested && UseOBSRecording {
                RolloverOBS("safe")
                gRolloverRequested := false
                SlackNotify("🎥 OBS rollover (safe) after match #" gLoopCount, "info")
            }

            return true
        }

        Sleep PollInterval
    }
}

AutomationAdvanceToNextReplay() {
    global Delay_BetweenItems

    EnsureFocusGame()
    SendNextSelection()
    Sleep Delay_BetweenItems
}

AutomationFinalizeRun(stopReason) {
    global UseOBSRecording, gRecording, Key_StopRec
    global gRunning, gSafeStopRequested, gRolloverRequested

    if UseOBSRecording && gRecording {
        OBSStopRecording()
        gRecording := false
        TrayTip "録画停止", (stopReason = "safe_stop" ? "安全停止により停止" : "通し録画を停止"), 1200
        Log("OBS: stop recording")
    }

    gRunning := false
    gSafeStopRequested := false
    gRolloverRequested := false
    Log("END: automation")
}

AutomationResolveStopReason() {
    global TotalMatches, gLoopCount, gSafeStopRequested, gRunning

    if (TotalMatches > 0 && gLoopCount >= TotalMatches)
        return "reached_total"
    if gSafeStopRequested
        return "safe_stop"
    if !gRunning
        return "stopped"
    return "unknown"
}
