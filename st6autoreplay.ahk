; ============================================================
; SF6 自動リプレイ：終了UI検出 / 通し録画 / 決定2回
; 検出強化＆診断 + 安全停止 + タイムリミット + 録画ローテ
; + GUI（タブ式） + 共通ステータス（フラット表示） + 設定保存/読込
; + ログ出力（自動スクロール） + 起動チェック + 自動リフォーカス
; + 画像パス改行/相対対応 + SF6操作前のフォーカス保証
; AutoHotkey v2
; ============================================================

#SingleInstance Force
#Requires AutoHotkey v2.0

; ---------- 既定値 ----------
MatchHardTimeoutSec := 300
PollInterval := 500

Delay_AfterFirstConfirm := 500
Delay_AfterPlayKey      := 500
Delay_BeforeNavigate    := 180
Delay_AfterBackKey      := 6000
Delay_BetweenItems      := 2000

Key_Confirm  := "F"
Key_Down     := "S"
Key_Up       := "W"

EndDownCount    := 1
EndConfirmCount := 1

NextDirection  := "down"      ; "down" / "up"
NextRepeats    := 1
NextIntervalMs := 220

; ---- OBS ----
Key_StartRec   := "^{F7}"
Key_StopRec    := "^{F8}"
Key_ToggleRec  := "^{F9}"       ; 録画の開始/停止 切替（OBS側で割当）
UseOBSRecording := true
UseOBSToggleForRollover := true ; ローテはトグル1回で実行

MaxRunMinutes   := 60   ; 0 = 無効（時間条件）
RolloverMinutes := 30   ; 0 = 無効（ファイル切替）
RolloverMode    := "safe"  ; "safe" / "instant"

TotalMatches := 50      ; 0 = 無限（回数条件）

; 画像（GUIは相対+改行、内部は絶対配列）
Img_Ends := [
  A_ScriptDir "\assets\end_result1.png",
  A_ScriptDir "\assets\end_result2.png",
  A_ScriptDir "\assets\end_result3.png",
  A_ScriptDir "\assets\end_result4.png",
  A_ScriptDir "\assets\end_result5.png",
  A_ScriptDir "\assets\end_result6.png"
]
ToleranceEnd := 180

GetROI_End_Default() {
    sw := A_ScreenWidth, sh := A_ScreenHeight
    return {x1: Round(sw*0.10), y1: Round(sh*0.45), x2: Round(sw*0.90), y2: Round(sh*0.98)}
}
GetROI_Full() {
    sw := A_ScreenWidth, sh := A_ScreenHeight
    return {x1: 0, y1: 0, x2: sw-1, y2: sh-1}
}

; ---- ウィンドウ識別 ----
OBSWinSelector  := "ahk_exe obs64.exe"
codes := [83,116,114,101,101,116,70,105,103,104,116,101,114,54,46,101,120,101]
exe := ""
for c in codes
    exe .= Chr(c)
GameWinSelector := "ahk_exe " exe
AutoRefocusGame := true

; ---- 起動チェック ----
CheckOnStart_Game := true
CheckOnStart_OBS  := true  ; UseOBSRecording=false なら自動スキップ

; ---- 設定ファイル・ログ ----
ConfigPath := A_ScriptDir "\config.ini"
AutoSaveOnExit := true
LogEnabled := true
LogDir := A_ScriptDir "\logs"

; ---------- 内部状態 ----------
global gRunning := false
global gPaused := false
global gRecording := false
global gUseFullROI := true
global gSafeStopRequested := false
global gRunStartTick := 0
global gRolloverRequested := false
global gLastRolloverTick := 0
global gLoopCount := 0

; ============================================================
; GUI（タブ式 + 共通フラット・ステータス）
; ============================================================
main := Gui("+Resize +MinSize740x250", "SF6 自動リプレイ")
main.SetFont("s9")

tab := main.Add("Tab3", "x10 y10 w700 h440", ["基本設定","詳細設定","ログ"])

; -------------------- 基本設定タブ --------------------
tab.UseTab(1)
grpBasic := main.Add("GroupBox", "x20 y45 w680 h400", "基本設定")

main.Add("Text", "x35 y70 w200", "終了画像パス（相対 / 改行）")
edtImgs := main.Add("Edit", "x35 y90 w400 r7 -Wrap")

main.Add("Text", "x35 y210 w200", "録画開始キー / 停止キー")
edtStart := main.Add("Edit", "x35 y230 w80", Key_StartRec)
main.Add("Text", "x120 y233 w15 Center", "/")
edtStop  := main.Add("Edit", "x140 y230 w80", Key_StopRec)

main.Add("Text", "x250 y210 w185", "OBSウィンドウ識別子")
edtObsSel := main.Add("Edit", "x250 y230 w185", OBSWinSelector)

main.Add("Text", "x450 y210 w210", "ゲームウィンドウ識別子")
edtGameSel := main.Add("Edit", "x450 y230 w210", GameWinSelector)

chkUseOBS := main.Add("CheckBox", "x35 y270 w200", "OBS録画を使う")
chkUseOBS.Value := UseOBSRecording ? 1 : 0
chkRefocus := main.Add("CheckBox", "x250 y270 w230 +Wrap", "GUI/OBS後にゲームへ戻す")
chkRefocus.Value := AutoRefocusGame ? 1 : 0

chkChkGame := main.Add("CheckBox", "x35 y300 w240", "開始時にゲーム起動をチェック")
chkChkGame.Value := CheckOnStart_Game ? 1 : 0
chkChkOBS := main.Add("CheckBox", "x280 y300 w300", "開始時にOBS起動をチェック（OBS録画ON時）")
chkChkOBS.Value := CheckOnStart_OBS ? 1 : 0

; 操作ボタン
btnStart := main.Add("Button", "x35 y345 w120 h28", "開始")
btnSafe  := main.Add("Button", "x165 y345 w120 h28", "安全停止")
btnForce := main.Add("Button", "x295 y345 w120 h28", "即時停止")
btnPause := main.Add("Button", "x425 y345 w120 h28", "一時停止")
btnApply := main.Add("Button", "x555 y345 w120 h28", "適用")

; 設定ファイル
btnLoad := main.Add("Button", "x35 y385 w160 h28", "読込（INI）")
btnSave := main.Add("Button", "x205 y385 w160 h28", "保存（INI）")

; テスト
btnDetect := main.Add("Button", "x375 y385 w150 h28", "検出テスト")
btnOBSon  := main.Add("Button", "x535 y385 w70 h28", "OBS開始")
btnOBSoff := main.Add("Button", "x610 y385 w70 h28", "OBS停止")

; -------------------- 詳細設定タブ --------------------
tab.UseTab(2)
grpAdv := main.Add("GroupBox", "x20 y45 w680 h500", "詳細設定")

; 上段
main.Add("Text", "x35 y70 w160", "次の選択方向（一覧）")
ddlDir := main.Add("DropDownList", "x35 y88 w120", ["down","up"])

main.Add("Text", "x180 y70 w200", "回数（TotalMatches / 0=無限）")
edtMatches := main.Add("Edit", "x180 y88 w120")

main.Add("Text", "x325 y70 w200", "タイムリミット（分 / 0=無効）")
edtMaxMin := main.Add("Edit", "x325 y88 w120")

main.Add("Text", "x470 y70 w200", "録画ローテ（分 / 0=無効）")
edtRollMin := main.Add("Edit", "x470 y88 w120")

main.Add("Text", "x35 y120 w160", "ローテ方式")
ddlRollMode := main.Add("DropDownList", "x35 y138 w120", ["safe","instant"])

; 中段
main.Add("Text", "x180 y120 w200", "Tolerance（0-255）")
edtTol := main.Add("Edit", "x180 y138 w120")

chkROI := main.Add("CheckBox", "x325 y138 w220", "ROI = 全画面（診断向け）")

main.Add("Text", "x470 y120 w200", "次移動の回数 / 間隔(ms)")
edtNextRep := main.Add("Edit", "x470 y138 w50")
main.Add("Text", "x525 y141 w20 Center", "×")
edtNextInt := main.Add("Edit", "x548 y138 w70")

; 下段（遅延）
main.Add("Text", "x35 y180 w200", "開始→決定2回の間隔(ms)")
edtD1 := main.Add("Edit", "x35 y198 w120", Delay_AfterFirstConfirm)
main.Add("Text", "x180 y180 w200", "決定後の小休止(ms)")
edtD2 := main.Add("Edit", "x180 y198 w120", Delay_AfterPlayKey)
main.Add("Text", "x325 y180 w200", "終了検出後の待機(ms)")
edtD3 := main.Add("Edit", "x325 y198 w120", Delay_BeforeNavigate)
main.Add("Text", "x470 y180 w200", "戻り後の待機(ms)")
edtD4 := main.Add("Edit", "x470 y198 w120", Delay_AfterBackKey)

main.Add("Text", "x35 y230 w200", "次のリプレイへ(ms)")
edtD5 := main.Add("Edit", "x35 y248 w120", Delay_BetweenItems)

; ログ設定
main.Add("Text", "x180 y230 w220", "ログ出力フォルダ")
edtLogDir := main.Add("Edit", "x180 y248 w260", LogDir)
chkLog := main.Add("CheckBox", "x450 y248 w200", "ログをファイルに保存")
chkLog.Value := LogEnabled ? 1 : 0

; OBS切替キーとローテ使用
main.Add("Text", "x180 y280 w200", "OBS切替キー（任意）")
edtToggle := main.Add("Edit", "x180 y298 w120", Key_ToggleRec)
chkUseToggle := main.Add("CheckBox", "x325 y298 w220", "ローテは切替キーで行う")
chkUseToggle.Value := UseOBSToggleForRollover ? 1 : 0

; -------------------- ログタブ --------------------
tab.UseTab(3)
grpLog := main.Add("GroupBox", "x20 y45 w680 h500", "ログ")
logBox := main.Add("Edit", "x35 y70 w650 h460 ReadOnly -Wrap -VScroll", "")

tab.UseTab() ; タブ終了

; ---- 共通フラット・ステータス（タブ外・枠なしの Text）----
statusText := main.Add("Text", "x10 y460 w700 h24 vStatusText", "")
statusText.SetFont("s9", "Segoe UI")
SetStatus("準備完了")

; ---- 画面初期化 ----
if FileExist(ConfigPath) {
    LoadConfig(ConfigPath)
}
UpdateGuiFromVars()
main.Show("NA")

; ---- ステータス更新タイマー ----
UpdateStatusText()
SetTimer(UpdateStatusText, 1000)

; ============================================================
; ボタン動作（最後でゲームに戻す）
; ============================================================
btnStart.OnEvent("Click", (*) => (ApplyGuiToVars(), StartAutomation(), RefocusGame()))
btnSafe.OnEvent("Click",  (*) => (RequestSafeStop(), RefocusGame()))
btnForce.OnEvent("Click", (*) => (ForceStopAutomation(), RefocusGame()))
btnPause.OnEvent("Click", (*) => (TogglePause(), UpdatePauseBtn(), RefocusGame()))
btnApply.OnEvent("Click", (*) => (ApplyGuiToVars(), RefocusGame()))
btnLoad.OnEvent("Click",  (*) => (LoadConfig(ConfigPath), UpdateGuiFromVars(), TrayTip("読込","設定を読み込みました",900)))
btnSave.OnEvent("Click",  (*) => (ApplyGuiToVars(), SaveConfig(ConfigPath), TrayTip("保存","設定を保存しました",900)))

btnDetect.OnEvent("Click",(*) => (QuickDetectTest(), RefocusGame()))
btnOBSon.OnEvent("Click", (*) => (FocusedTriggerOBS(Key_StartRec), RefocusGame()))
btnOBSoff.OnEvent("Click",(*) => (FocusedTriggerOBS(Key_StopRec),  RefocusGame()))

; GUI 開閉ホットキー
^!g:: (main.Visible := !main.Visible, !main.Visible ? "" : RefocusGame())

; ============================================================
; ホットキー
; ============================================================
^!s:: (StartAutomation(), RefocusGame())
^!x:: (RequestSafeStop(), RefocusGame())
^!p:: (TogglePause(), UpdatePauseBtn(), RefocusGame())
^!w:: (QuickDetectTest(), RefocusGame())
^!r:: (ToggleFullROI(), RefocusGame())
^!e:: (SendEndNavigateTest(), RefocusGame())
^!t:: (SendOBSTest(), RefocusGame())
^+!x:: (ForceStopAutomation(), RefocusGame())
^!u::  ; up/downトグル
{
    global NextDirection
    NextDirection := (StrLower(NextDirection) = "down") ? "up" : "down"
    TrayTip "次の選択方向", (NextDirection="up"?"上へ":"下へ"), 1200
    RefocusGame()
}
^!l:: (LoadConfig(ConfigPath), UpdateGuiFromVars(), TrayTip("読込","設定を読み込みました",900), RefocusGame())
^!k:: (ApplyGuiToVars(), SaveConfig(ConfigPath), TrayTip("保存","設定を保存しました",900), RefocusGame())

OnExit(ExitHandler)

; ============================================================
; メイン処理
; ============================================================
StartAutomation() {
    global gRunning, gPaused, gRecording, TotalMatches
    global gSafeStopRequested, gRunStartTick
    global RolloverMinutes, RolloverMode, gRolloverRequested, gLastRolloverTick
    global MaxRunMinutes, gLoopCount
    global UseOBSRecording, CheckOnStart_Game, CheckOnStart_OBS

    if gRunning {
        TrayTip "実行中", "安全停止は Ctrl+Alt+X（またはGUI）", 1500
        return
    }
    ; 起動チェック
    if CheckOnStart_Game && !WinExist(GameWinSelector) {
        MsgBox "ゲームのウィンドウが見つかりません:`n" GameWinSelector, "起動チェック", 48
        Log("WARN: Game window not found: " GameWinSelector)
        return
    }
    if UseOBSRecording && CheckOnStart_OBS && !WinExist(OBSWinSelector) {
        MsgBox "OBSのウィンドウが見つかりません:`n" OBSWinSelector, "起動チェック", 48
        Log("WARN: OBS window not found: " OBSWinSelector)
        return
    }

    for img in Img_Ends
        if !FileExist(img) {
            MsgBox "終了検出画像が見つかりません:`n" img, "エラー", 16
            Log("ERROR: End image missing: " img)
            return
        }

    gRunning := true
    gPaused := false
    gSafeStopRequested := false
    gRolloverRequested := false
    gRunStartTick := A_TickCount
    gLoopCount := 0

    CoordMode "Pixel", "Screen"
    RefocusGame(true)
    Log("START: automation")

    ; 録画開始
    if UseOBSRecording && !gRecording {
        FocusedTriggerOBS(Key_StartRec)
        gRecording := true
        gLastRolloverTick := A_TickCount
        TrayTip "録画開始", "通し録画を開始", 1200
        Log("OBS: start recording")
        Sleep 200
    } else {
        gLastRolloverTick := A_TickCount
    }

    Loop {
        if !gRunning
            break
        if (TotalMatches > 0 && gLoopCount >= TotalMatches)
            break

        ; タイムリミット
        if (MaxRunMinutes > 0) {
            if !gSafeStopRequested && (A_TickCount - gRunStartTick) > (MaxRunMinutes * 60 * 1000) {
                gSafeStopRequested := true
                TrayTip "安全停止", "タイムリミット超過。次の終了UIで停止します", 1500
                Log("SAFE-STOP: due to MaxRunMinutes")
            }
        }

        ; 録画ローテ
        if (UseOBSRecording && RolloverMinutes > 0 && !gSafeStopRequested) {
            if (A_TickCount - gLastRolloverTick) > (RolloverMinutes * 60 * 1000) {
                if (RolloverMode = "instant") {
                    RolloverOBS("instant")              ; Ctrl+F9 1回
                } else {
                    gRolloverRequested := true
                    TrayTip "ローテ待機", "次の終了UIで録画を切替", 1200
                    Log("OBS: rollover requested (safe)")
                }
            }
        }

        if gSafeStopRequested
            break

        gLoopCount += 1
        Log("REPLAY: start #" gLoopCount)

        while gPaused && gRunning
            Sleep 150
        if !gRunning
            break

        ; 再生開始（決定2回）※操作直前に必ずフォーカス
        EnsureFocusSF6()
        Press(Key_Confirm, 80)
        Sleep Delay_AfterFirstConfirm
        EnsureFocusSF6()
        Press(Key_Confirm, 80)
        Sleep Delay_AfterPlayKey

        ; 終了UI待ち
        startTick := A_TickCount
        Loop {
            if !gRunning
                break
            while gPaused && gRunning
                Sleep 150
            if !gRunning
                break

            if (A_TickCount - startTick) > (MatchHardTimeoutSec * 1000) {
                Log("TIMEOUT: no end UI within " MatchHardTimeoutSec "s")
                break
            }

            roi := (gUseFullROI ? GetROI_Full() : GetROI_End_Default())
            if FindAnyImage(Img_Ends, roi, ToleranceEnd, &fx, &fy) {
                Log("DETECT: end UI at " fx "," fy)
                Sleep Delay_BeforeNavigate

                ; 終了UIで戻る操作 ※直前にフォーカス
                EnsureFocusSF6()
                Loop EndDownCount {
                    Press(Key_Down, 60)
                    Sleep 500
                }
                EnsureFocusSF6()
                Loop EndConfirmCount {
                    Press(Key_Confirm, 70)
                    Sleep 180
                }
                Sleep Delay_AfterBackKey
                WaitImageDisappear(Img_Ends, roi, ToleranceEnd, 1200)

                if gRolloverRequested && !gSafeStopRequested && UseOBSRecording {
                    RolloverOBS("safe")                 ; Ctrl+F9 1回
                    gRolloverRequested := false
                }
                break
            }
            Sleep PollInterval
        }

        if gSafeStopRequested
            break

        ; 次のリプレイへ ※直前にフォーカス
        EnsureFocusSF6()
        SendNextSelection()
        Sleep Delay_BetweenItems
    }

    ; 終了処理
    if UseOBSRecording && gRecording {
        FocusedTriggerOBS(Key_StopRec)
        gRecording := false
        TrayTip "録画停止", (gSafeStopRequested ? "安全停止により停止" : "通し録画を停止"), 1200
        Log("OBS: stop recording")
    }

    gRunning := false
    gSafeStopRequested := false
    gRolloverRequested := false
    Log("END: automation")
}

; ---- OBSローテ：Ctrl+F9 を1回だけ送る ----
RolloverOBS(mode := "instant") {
    global UseOBSRecording, UseOBSToggleForRollover, Key_ToggleRec
    global Key_StopRec, Key_StartRec, gLastRolloverTick
    if !UseOBSRecording
        return

    if (UseOBSToggleForRollover && Key_ToggleRec != "") {
        Log("OBS: rollover via toggle(one-shot) (" mode ")")
        FocusedTriggerOBS(Key_ToggleRec)   ; ★ 1回だけ送る（OBSが停止→新規開始まで実行）
        Sleep 800                          ; 切替安定待ち（必要に応じ調整 800～1200ms）
    } else {
        Log("OBS: rollover via stop/start (" mode ")")
        FocusedTriggerOBS(Key_StopRec)
        Sleep 900
        FocusedTriggerOBS(Key_StartRec)
    }

    gLastRolloverTick := A_TickCount
    TrayTip "ローテ", "録画ファイルを切替（" (mode="instant"?"即時":"試合間") "）", 1200
}

SendNextSelection() {
    global NextDirection, Key_Down, Key_Up, NextRepeats, NextIntervalMs
    dir := StrLower(NextDirection)
    key := (dir = "up") ? Key_Up : Key_Down
    Loop Max(1, NextRepeats) {
        Press(key, 60)
        Sleep NextIntervalMs
    }
}

RequestSafeStop() {
    global gRunning, gRecording, gSafeStopRequested
    if !gRunning {
        if UseOBSRecording && gRecording {
            FocusedTriggerOBS(Key_StopRec)
            gRecording := false
            TrayTip "録画停止", "（実行外）録画を停止しました", 1200
            Log("OBS: stop recording (out of run)")
        }
        TrayTip "停止", "既に停止しています", 1000
        return
    }
    gSafeStopRequested := true
    TrayTip "安全停止", "次の終了UIで停止します", 1500
    Log("SAFE-STOP: requested")
}

ForceStopAutomation() {
    global gRunning, gPaused, gRecording, Key_StopRec
    global gSafeStopRequested, gRolloverRequested
    gSafeStopRequested := false
    gRolloverRequested := false
    gRunning := false
    gPaused := false
    if UseOBSRecording && gRecording {
        FocusedTriggerOBS(Key_StopRec)
        gRecording := false
        TrayTip "録画停止", "即時停止しました", 1200
        Log("OBS: stop recording (force)")
    }
    TrayTip "停止", "スクリプトを停止しました", 1000
    Log("FORCE-STOP: script")
}

TogglePause() {
    global gPaused
    gPaused := !gPaused
    TrayTip (gPaused ? "一時停止" : "再開"), "", 900
    Log(gPaused ? "PAUSE" : "RESUME")
}

; ============================================================
; 設定 保存/読込
; ============================================================
SaveConfig(path) {
    global NextDirection, TotalMatches, MaxRunMinutes, RolloverMinutes, RolloverMode
    global ToleranceEnd, gUseFullROI, NextRepeats, NextIntervalMs
    global Key_StartRec, Key_StopRec, Key_ToggleRec, OBSWinSelector, GameWinSelector, AutoRefocusGame
    global Img_Ends, UseOBSRecording, UseOBSToggleForRollover, CheckOnStart_Game, CheckOnStart_OBS
    global LogEnabled, LogDir
    IniWrite(NextDirection,  path, "main", "NextDirection")
    IniWrite(TotalMatches,   path, "main", "TotalMatches")
    IniWrite(MaxRunMinutes,  path, "main", "MaxRunMinutes")
    IniWrite(RolloverMinutes,path, "main", "RolloverMinutes")
    IniWrite(RolloverMode,   path, "main", "RolloverMode")
    IniWrite(ToleranceEnd,   path, "main", "ToleranceEnd")
    IniWrite(gUseFullROI?1:0,path, "main", "UseFullROI")
    IniWrite(NextRepeats,    path, "main", "NextRepeats")
    IniWrite(NextIntervalMs, path, "main", "NextIntervalMs")
    IniWrite(Key_StartRec,   path, "obs",  "StartKey")
    IniWrite(Key_StopRec,    path, "obs",  "StopKey")
    IniWrite(Key_ToggleRec,  path, "obs",  "ToggleKey")
    IniWrite(OBSWinSelector, path, "obs",  "WindowSelector")
    IniWrite(UseOBSRecording?1:0, path, "obs", "UseRecording")
    IniWrite(UseOBSToggleForRollover?1:0, path, "obs", "UseToggleForRollover")
    IniWrite(CheckOnStart_OBS?1:0, path, "obs", "CheckOnStart")
    IniWrite(GameWinSelector,           path, "game", "WindowSelector")
    IniWrite(AutoRefocusGame?1:0,       path, "game", "AutoRefocus")
    IniWrite(CheckOnStart_Game?1:0,     path, "game", "CheckOnStart")
    IniWrite(JoinList(Img_Ends, ";"),   path, "images", "EndImages")
    IniWrite(LogEnabled?1:0, path, "log", "Enabled")
    IniWrite(LogDir,         path, "log", "Dir")
}

LoadConfig(path) {
    global NextDirection, TotalMatches, MaxRunMinutes, RolloverMinutes, RolloverMode
    global ToleranceEnd, gUseFullROI, NextRepeats, NextIntervalMs
    global Key_StartRec, Key_StopRec, Key_ToggleRec, OBSWinSelector, GameWinSelector, AutoRefocusGame
    global Img_Ends, UseOBSRecording, UseOBSToggleForRollover, CheckOnStart_Game, CheckOnStart_OBS
    global LogEnabled, LogDir
    NextDirection  := IniRead(path, "main", "NextDirection", NextDirection)
    TotalMatches   := Integer(IniRead(path, "main", "TotalMatches", TotalMatches))
    MaxRunMinutes  := Integer(IniRead(path, "main", "MaxRunMinutes", MaxRunMinutes))
    RolloverMinutes:= Integer(IniRead(path, "main", "RolloverMinutes", RolloverMinutes))
    RolloverMode   := IniRead(path, "main", "RolloverMode", RolloverMode)
    ToleranceEnd   := Integer(IniRead(path, "main", "ToleranceEnd", ToleranceEnd))
    gUseFullROI    := (Integer(IniRead(path, "main", "UseFullROI", gUseFullROI?1:0))=1)
    NextRepeats    := Integer(IniRead(path, "main", "NextRepeats", NextRepeats))
    NextIntervalMs := Integer(IniRead(path, "main", "NextIntervalMs", NextIntervalMs))
    Key_StartRec   := IniRead(path, "obs", "StartKey", Key_StartRec)
    Key_StopRec    := IniRead(path, "obs", "StopKey",  Key_StopRec)
    Key_ToggleRec  := IniRead(path, "obs", "ToggleKey", Key_ToggleRec)
    OBSWinSelector := IniRead(path, "obs", "WindowSelector", OBSWinSelector)
    UseOBSRecording := (Integer(IniRead(path, "obs", "UseRecording", UseOBSRecording?1:0))=1)
    UseOBSToggleForRollover := (Integer(IniRead(path, "obs", "UseToggleForRollover", UseOBSToggleForRollover?1:0))=1)
    CheckOnStart_OBS := (Integer(IniRead(path, "obs", "CheckOnStart", CheckOnStart_OBS?1:0))=1)
    GameWinSelector := IniRead(path, "game", "WindowSelector", GameWinSelector)
    AutoRefocusGame := (Integer(IniRead(path, "game", "AutoRefocus", AutoRefocusGame?1:0))=1)
    CheckOnStart_Game := (Integer(IniRead(path, "game", "CheckOnStart", CheckOnStart_Game?1:0))=1)
    imgs := IniRead(path, "images", "EndImages", "")
    if (imgs != "") {
        Img_Ends := SplitList(imgs, ";")
    }
    LogEnabled := (Integer(IniRead(path, "log", "Enabled", LogEnabled?1:0))=1)
    LogDir     := IniRead(path, "log", "Dir", LogDir)
}

JoinList(arr, sep:=";") {
    out := ""
    for i, v in arr {
        out .= (i>1 ? sep : "") . v
    }
    return out
}
SplitList(txt, sep:=";") {
    parts := StrSplit(txt, sep)
    out := []
    for p in parts {
        p := Trim(p)
        if (p != "")
            out.Push(p)
    }
    return out
}

ApplyGuiToVars() {
    global NextDirection, TotalMatches, MaxRunMinutes, RolloverMinutes, RolloverMode
    global ToleranceEnd, gUseFullROI, NextRepeats, NextIntervalMs
    global Key_StartRec, Key_StopRec, Key_ToggleRec, OBSWinSelector, Img_Ends
    global GameWinSelector, AutoRefocusGame, UseOBSRecording, UseOBSToggleForRollover, CheckOnStart_Game, CheckOnStart_OBS
    global LogEnabled, LogDir
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
    Key_StopRec  := edtStop.Text
    OBSWinSelector := edtObsSel.Text
    ; 画像GUI(改行)→絶対パス配列
    Img_Ends := []
    txt := edtImgs.Value
    for line in StrSplit(txt, ["`r","`n"], true) {
        line := Trim(line)
        if (line != "") {
            if (SubStr(line,1,1)="\" || InStr(line,":\")) {
                Img_Ends.Push(line)
            } else {
                Img_Ends.Push(A_ScriptDir "\" line)
            }
        }
    }
    GameWinSelector := edtGameSel.Text
    AutoRefocusGame := !!chkRefocus.Value
    UseOBSRecording := !!chkUseOBS.Value
    Key_ToggleRec := edtToggle.Text
    UseOBSToggleForRollover := !!chkUseToggle.Value
    CheckOnStart_Game := !!chkChkGame.Value
    CheckOnStart_OBS  := !!chkChkOBS.Value
    LogEnabled := !!chkLog.Value
    LogDir := edtLogDir.Text
}
UpdateGuiFromVars() {
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
    edtStop.Text  := Key_StopRec
    edtObsSel.Text := OBSWinSelector
    ; 配列→相対パス改行
    rels := []
    for img in Img_Ends {
        rel := StrReplace(img, A_ScriptDir "\")
        rels.Push(rel)
    }
    edtImgs.Value := StrJoin(rels, "`n")
    edtGameSel.Text := GameWinSelector
    chkRefocus.Value := AutoRefocusGame ? 1 : 0
    chkUseOBS.Value := UseOBSRecording ? 1 : 0
    chkChkGame.Value := CheckOnStart_Game ? 1 : 0
    chkChkOBS.Value := CheckOnStart_OBS ? 1 : 0
    edtToggle.Text := Key_ToggleRec
    chkUseToggle.Value := UseOBSToggleForRollover ? 1 : 0
    chkLog.Value := LogEnabled ? 1 : 0
    edtLogDir.Text := LogDir
    UpdatePauseBtn()
}

StrJoin(arr, sep:="`n") {
    s := ""
    for i,v in arr
        s .= (i>1?sep:"") v
    return s
}

ToIntSafe(txt, def) {
    if (txt = "")
        return def
    try {
        return Integer(txt)
    } catch as e {
        return def
    }
}

UpdatePauseBtn() {
    btnPause.Text := (gPaused ? "再開" : "一時停止")
}
UpdateStatusText() {
    global gRunning, gPaused, gRecording, gRunStartTick, gLoopCount, UseOBSRecording
    runMin := (gRunStartTick>0) ? Round((A_TickCount - gRunStartTick)/60000, 1) : 0
    base := "状態: " (gRunning ? (gPaused ? "一時停止中" : "実行中") : "停止")
        . " / 録画: " (UseOBSRecording ? (gRecording ? "ON" : "OFF") : "未使用")
        . " / 経過: " runMin "分"
        . " / ループ: " gLoopCount
    SetStatus(base)
    UpdatePauseBtn()
}

ExitHandler(*) {
    global AutoSaveOnExit, ConfigPath
    if AutoSaveOnExit {
        ApplyGuiToVars()
        SaveConfig(ConfigPath)
    }
}

; ============================================================
; 診断・OBS送信・ユーティリティ
; ============================================================
QuickDetectTest() {
    roi := (gUseFullROI ? GetROI_Full() : GetROI_End_Default())
    if FindAnyImage(Img_Ends, roi, ToleranceEnd, &x, &y) {
        TrayTip "検出OK", "座標: " x "," y "  ROI=" (gUseFullROI?"Full":"Default"), 1200
        Log("TEST: detect OK at " x "," y)
    } else {
        TrayTip "未検出", "ROI=" (gUseFullROI?"Full":"Default") " / tol=" ToleranceEnd, 1200
        Log("TEST: detect NG")
    }
}
ToggleFullROI() {
    global gUseFullROI
    gUseFullROI := !gUseFullROI
    chkROI.Value := gUseFullROI ? 1 : 0
    TrayTip "ROI切替", (gUseFullROI ? "全画面" : "中央～下"), 1000
    Log("ROI: " (gUseFullROI ? "Full" : "Default"))
}
SendEndNavigateTest() {
    global EndDownCount, EndConfirmCount, Key_Down, Key_Confirm
    EnsureFocusSF6()
    Loop EndDownCount {
        Press(Key_Down, 60)
        Sleep 120
    }
    EnsureFocusSF6()
    Loop EndConfirmCount {
        Press(Key_Confirm, 70)
        Sleep 180
    }
    TrayTip "送信", "S→F をテスト送信", 800
    Log("TEST: sent S→F")
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

; ===== フォーカス制御 =====
EnsureFocusSF6() {
    global GameWinSelector
    if WinExist(GameWinSelector) {
        try WinRestore GameWinSelector
        WinActivate GameWinSelector
        if WinWaitActive(GameWinSelector, , 500) {
            Sleep 60
            return true
        }
    }
    TrayTip "警告", "SF6のウィンドウが見つかりません / アクティブ化できません", 1200
    Log("FOCUS: SF6 window not found or activate failed: " GameWinSelector)
    return false
}

RefocusGame(force := false) {
    global AutoRefocusGame
    if (!force && !AutoRefocusGame)
        return
    EnsureFocusSF6()
}

RestoreFocus(prevHwnd) {
    ; 旧互換：使わないが残しておく
    EnsureFocusSF6()
}

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
        Sleep 100                 ; 少し長めに安定待ち
        Send keyToSend
        Sleep 60
        ; OBS操作後は自動でSF6へ戻す（設定に従う）
        if AutoRefocusGame
            EnsureFocusSF6()
        else if prev
            WinActivate prev
        Log("OBS: key sent [" keyToSend "]")
    } else {
        TrayTip "OBS未検出", OBSWinSelector " が見つかりません", 1200
        Log("ERROR: OBS window not found for key [" keyToSend "]")
    }
}

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
        if !FindAnyImage(Img_Ends, roi, tol, &x, &y)
            return true
        Sleep 80
    }
    return false
}

; ===== 共通ステータス（フラット） & ログ =====
SetStatus(text) {
    global statusText
    if IsSet(statusText) && statusText
        statusText.Value := text
}
TruncForStatus(msg, maxChars := 60) {
    msg := RegExReplace(msg, "\R", " ")
    return (StrLen(msg) > maxChars) ? SubStr(msg, 1, maxChars-1) "…" : msg
}
Log(msg) {
    global LogEnabled, LogDir, logBox
    ; 画面（ログタブ）へ追記（末尾へ自動スクロール）
    if IsSet(logBox) && logBox {
        try {
            logBox.Value .= FormatTime(A_Now, "HH:mm:ss") " - " msg "`r`n"
            static EM_SETSEL := 0x00B1, EM_SCROLLCARET := 0x00B7
            SendMessage EM_SETSEL, -1, -1, , logBox.Hwnd
            SendMessage EM_SCROLLCARET, 0, 0, , logBox.Hwnd
        }
    }
    ; 最新行を下部ステータスへ
    SetStatus("最新: " TruncForStatus(msg))

    if !LogEnabled
        return
    ; ファイルへ
    try {
        DirCreate(LogDir)
        stamp := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
        file := LogDir "\" FormatTime(A_Now, "yyyyMMdd") ".log"
        FileAppend(stamp " - " msg "`r`n", file, "UTF-8-RAW")
    } catch as e {
        ; 失敗は握りつぶし
    }
}
