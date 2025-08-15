; ============================================================
; SF6 自動リプレイ：終了UIだけ画像検出 / 通し録画 / 決定2回で再生開始
; 検出強化＆診断ホットキー付き + 安全停止 + タイムリミット + 録画ローテーション
; 操作:
;   Ctrl+Alt+S  開始
;   Ctrl+Alt+X  安全停止（次の終了UIで録画停止して終了）
;   Ctrl+Alt+P  一時停止/再開
;   Ctrl+Alt+W  その場で検出テスト
;   Ctrl+Alt+R  ROI=全画面と既定の切替
;   Ctrl+Alt+E  S→F テスト送信
;   Ctrl+Alt+T  OBSキー送信テスト
;   Ctrl+Alt+Shift+X  即時停止（強制）
; ============================================================

; ---- 設定 ----
MatchHardTimeoutSec := 300       ; 最大5分（BO2想定）
PollInterval := 500              ; 終了UIポーリング間隔(ms)

; 再生開始（決定2回）
Delay_AfterFirstConfirm := 500   ; 1回目→2回目の間隔
Delay_AfterPlayKey      := 500   ; 2回目後の小休止

; 終了検出後の操作
Delay_BeforeNavigate := 180      ; 検出直後の短い待機（演出残り対策）
Delay_AfterBackKey   := 5000     ; 下→決定のあと待機
Delay_BetweenItems   := 220      ; 次のリプレイへ

; キー
Key_Confirm  := "F"              ; 決定
Key_Down     := "S"              ; 下

; 終了UIでの戻り操作
EndDownCount    := 1
EndConfirmCount := 1

; OBS 録画ホットキー（OBSの設定に合わせて）
Key_StartRec := "^{F7}"          ; Ctrl+F7
Key_StopRec  := "^{F8}"          ; Ctrl+F8

; タイムリミット（最終停止）
; ★ 0 の場合は時間判定をスキップし、回数条件（TotalMatches）のみ参照
MaxRunMinutes := 120

; ★録画ファイルの時間ローテーション
; 0 なら無効。>0 なら指定分を超えたら録画を切替
RolloverMinutes := 60
; "safe" = 次の終了UIで停止→すぐ開始（試合間で継ぎ目）/ "instant" = その場で停止→開始
RolloverMode := "safe"

TotalMatches := 50               ; 0=無限

; ---- 終了UI 画像 ----
Img_Ends := [
  A_ScriptDir "\assets\end_result1.png",
  A_ScriptDir "\assets\end_result2.png",
  A_ScriptDir "\assets\end_result3.png",
  A_ScriptDir "\assets\end_result4.png",
  A_ScriptDir "\assets\end_result5.png",
  A_ScriptDir "\assets\end_result6.png"
]

; 許容度
ToleranceEnd := 180

; ROI
GetROI_End_Default() {
    sw := A_ScreenWidth, sh := A_ScreenHeight
    return {x1: Round(sw*0.10), y1: Round(sh*0.45), x2: Round(sw*0.90), y2: Round(sh*0.98)}
}
GetROI_Full() {
    sw := A_ScreenWidth, sh := A_ScreenHeight
    return {x1: 0, y1: 0, x2: sw-1, y2: sh-1}
}

; ---- OBSウィンドウ識別（フォーカス一時移動）----
OBSWinSelector := "ahk_exe obs64.exe"   ; 必要に応じて "ahk_class OBSProject" 等に

; -------------- 内部 --------------
global gRunning := false
global gPaused := false
global gRecording := false
global gUseFullROI := true
global gSafeStopRequested := false       ; 最終停止（次の終了で停止）
global gRunStartTick := 0                ; 実行開始時刻

; ローテーション用
global gRolloverRequested := false       ; "safe" のとき：次の終了で切替
global gLastRolloverTick := 0            ; 直近の開始時刻（録画再スタート時に更新）

^!s:: StartAutomation()
^!x:: RequestSafeStop()                  ; 安全停止リクエスト
^!p:: TogglePause()

; 診断
^!w:: QuickDetectTest()
^!r:: ToggleFullROI()
^!e:: SendEndNavigateTest()
^!t:: SendOBSTest()

; 即時停止（強制）
^+!x:: ForceStopAutomation()

StartAutomation() {
    global gRunning, gPaused, gRecording, TotalMatches
    global gSafeStopRequested, gRunStartTick
    global RolloverMinutes, RolloverMode, gRolloverRequested, gLastRolloverTick
    global MaxRunMinutes

    if gRunning {
        TrayTip "実行中", "Stop（Ctrl+Alt+X）で安全停止できます", 1500
        return
    }
    for img in Img_Ends
        if !FileExist(img) {
            MsgBox "終了検出画像が見つかりません:`n" img, "エラー", 16
            return
        }

    gRunning := true
    gPaused := false
    gSafeStopRequested := false
    gRolloverRequested := false
    gRunStartTick := A_TickCount

    CoordMode "Pixel", "Screen"

    ; 通し録画開始（最初の1回）
    if !gRecording {
        FocusedTriggerOBS(Key_StartRec)
        gRecording := true
        gLastRolloverTick := A_TickCount
        TrayTip "録画開始", "通し録画を開始", 1200
        Sleep 200
    }

    loopCount := 0
    Loop {
        if !gRunning
            break
        if (TotalMatches > 0 && loopCount >= TotalMatches)
            break

        ; --- タイムリミット（最終停止） ---
        if (MaxRunMinutes > 0) {
            if !gSafeStopRequested && (A_TickCount - gRunStartTick) > (MaxRunMinutes * 60 * 1000) {
                gSafeStopRequested := true
                TrayTip "安全停止", "タイムリミット超過。次の終了UIで停止します", 1500
            }
        }

        ; --- 録画ローテ判定 ---
        if (RolloverMinutes > 0 && !gSafeStopRequested) {
            if (A_TickCount - gLastRolloverTick) > (RolloverMinutes * 60 * 1000) {
                if (RolloverMode = "instant") {
                    ; その場で停止→開始（途中で切れる可能性あり）
                    FocusedTriggerOBS(Key_StopRec)
                    Sleep 200
                    FocusedTriggerOBS(Key_StartRec)
                    gLastRolloverTick := A_TickCount
                    TrayTip "ローテ", "録画ファイルを切り替えました（即時）", 1200
                } else {
                    ; safe：次の終了UIで停止→開始
                    gRolloverRequested := true
                    TrayTip "ローテ待機", "次の終了UIで録画ファイルを切替", 1200
                }
            }
        }

        ; ★安全停止が要求されていたら、新規再生に入らず終了
        if gSafeStopRequested
            break

        loopCount += 1

        while gPaused && gRunning
            Sleep 150
        if !gRunning
            break

        ; ▼ 再生開始：決定2回
        Press(Key_Confirm, 80)
        Sleep Delay_AfterFirstConfirm
        Press(Key_Confirm, 80)
        Sleep Delay_AfterPlayKey

        ; --- 終了UI待ち ---
        startTick := A_TickCount
        Loop {
            if !gRunning
                break
            while gPaused && gRunning
                Sleep 150
            if !gRunning
                break

            ; 300秒たっても出なければ次へ（安全停止・ローテ待機があっても抜ける）
            if (A_TickCount - startTick) > (MatchHardTimeoutSec * 1000) {
                break
            }

            roi := (gUseFullROI ? GetROI_Full() : GetROI_End_Default())
            if FindAnyImage(Img_Ends, roi, ToleranceEnd, &fx, &fy) {
                ; 検出直後の短い待機
                Sleep Delay_BeforeNavigate

                ; 下→決定 で戻る
                Loop EndDownCount {
                    Press(Key_Down, 60)
                    Sleep 500
                }
                Loop EndConfirmCount {
                    Press(Key_Confirm, 70)
                    Sleep 180
                }
                Sleep Delay_AfterBackKey

                ; 終了UI消失確認
                WaitImageDisappear(Img_Ends, roi, ToleranceEnd, 1200)

                ; ★ここで "safe" ローテを実行（最終停止より優先度低）
                if gRolloverRequested && !gSafeStopRequested {
                    FocusedTriggerOBS(Key_StopRec)
                    Sleep 250
                    FocusedTriggerOBS(Key_StartRec)
                    gLastRolloverTick := A_TickCount
                    gRolloverRequested := false
                    TrayTip "ローテ", "録画ファイルを切り替えました（試合間）", 1500
                }
                break
            }
            Sleep PollInterval
        }

        ; ★ここで最終の安全停止が要求されていたら、次へ進まず終了
        if gSafeStopRequested
            break

        ; 次のリプレイへ
        Press(Key_Down, 60)
        Sleep Delay_BetweenItems
    }

    ; —— ループ終了（通常終了 / 安全停止 / タイムリミット）——
    if gRecording {
        FocusedTriggerOBS(Key_StopRec)
        gRecording := false
        TrayTip "録画停止", (gSafeStopRequested ? "安全停止により停止" : "通し録画を停止"), 1200
    }

    gRunning := false
    TrayTip "終了", (gSafeStopRequested ? "安全停止完了（試合終了後に停止）" : "自動連続再生を終了"), 1500
    gSafeStopRequested := false
    gRolloverRequested := false
}

; 安全停止：次の「終了UI」到達まで再生を続けてから止める
RequestSafeStop() {
    global gRunning, gRecording, gSafeStopRequested
    if !gRunning {
        if gRecording {
            FocusedTriggerOBS(Key_StopRec)
            gRecording := false
            TrayTip "録画停止", "（実行外）録画を停止しました", 1200
        }
        TrayTip "停止", "処理はすでに停止しています", 1000
        return
    }
    gSafeStopRequested := true
    TrayTip "安全停止", "次の終了UIまで待ってから停止します", 1500
}

; 即時停止：今すぐ止めたい場合
ForceStopAutomation() {
    global gRunning, gPaused, gRecording, Key_StopRec
    global gSafeStopRequested, gRolloverRequested
    gSafeStopRequested := false
    gRolloverRequested := false
    gRunning := false
    gPaused := false
    if gRecording {
        FocusedTriggerOBS(Key_StopRec)
        gRecording := false
        TrayTip "録画停止", "即時停止しました", 1200
    }
    TrayTip "停止", "スクリプトを停止しました", 1000
}

TogglePause() {
    global gPaused
    gPaused := !gPaused
    TrayTip (gPaused ? "一時停止" : "再開"), "", 900
}

; ---- 診断系 ----
QuickDetectTest() {
    roi := (gUseFullROI ? GetROI_Full() : GetROI_End_Default())
    if FindAnyImage(Img_Ends, roi, ToleranceEnd, &x, &y) {
        TrayTip "検出OK", "座標: " x "," y "  ROI=" (gUseFullROI?"Full":"Default"), 1200
    } else {
        TrayTip "未検出", "ROI=" (gUseFullROI?"Full":"Default") " / tol=" ToleranceEnd, 1200
    }
}
ToggleFullROI() {
    global gUseFullROI
    gUseFullROI := !gUseFullROI
    TrayTip "ROI切替", (gUseFullROI ? "全画面" : "中央～下"), 1000
}
SendEndNavigateTest() {
    global EndDownCount, EndConfirmCount, Key_Down, Key_Confirm
    Loop EndDownCount {
        Press(Key_Down, 60)
        Sleep 120
    }
    Loop EndConfirmCount {
        Press(Key_Confirm, 70)
        Sleep 180
    }
    TrayTip "送信", "S→F をテスト送信", 800
}

; ---- OBSキー送信テスト（フォーカス一時移動）----
SendOBSTest() {
    global Key_StartRec, Key_StopRec
    TrayTip "テスト", "録画開始キー送信: " Key_StartRec, 800
    FocusedTriggerOBS(Key_StartRec)
    Sleep 800
    TrayTip "テスト", "録画停止キー送信: " Key_StopRec, 800
    FocusedTriggerOBS(Key_StopRec)
}
FocusedTriggerOBS(keyToSend) {
    global OBSWinSelector
    prev := WinExist("A")
    if WinExist(OBSWinSelector) {
        WinActivate OBSWinSelector
        WinWaitActive OBSWinSelector, , 500
        Sleep 40
        Send keyToSend               ; 例: "{F7}" / "^{F7}"
        Sleep 60
        if prev {
            WinActivate prev
        }
    } else {
        TrayTip "OBS未検出", OBSWinSelector " が見つかりません", 1200
    }
}

; ---- 画像検索ユーティリティ ----
Press(key, holdMs := 60) {
    Send "{" key " down}"
    Sleep holdMs
    Send "{" key " up}"
}
FindAnyImage(imgList, roi, tol, &outX?, &outY?) {
    for img in imgList {
        if ImageSearch(&x, &y, roi.x1, roi.y1, roi.x2, roi.y2, "*n" tol " " img) {
            outX := x, outY := y
            return true
        }
    }
    return false
}
WaitImageDisappear(imgList, roi, tol, timeoutMs) {
    endTick := A_TickCount + timeoutMs
    while A_TickCount < endTick {
        if !FindAnyImage(imgList, roi, tol, &x, &y)
            return true
        Sleep 80
    }
    return false
}
