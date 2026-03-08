; ============================================================
; 自動リプレイ：終了UI検出 / 通し録画 / 決定2回
; 検出強化＆診断 + 安全停止 + タイムリミット + 録画ローテ
; + GUI（タブ式） + 共通ステータス（フラット表示） + 設定保存/読込
; + ログ出力（自動スクロール） + 起動チェック + 自動リフォーカス
; + 画像パス改行/相対対応 + GAME操作前のフォーカス保証
; AutoHotkey v2
; ============================================================
;@Ahk2Exe-SetName       SF6 Auto Replay
;@Ahk2Exe-SetDescription SF6 Auto Replay Recorder
;@Ahk2Exe-SetVersion    1.0.0.0
;@Ahk2Exe-SetCompanyName mmiyaji
;@Ahk2Exe-SetCopyright  (c) 2025 mmiyaji
;@Ahk2Exe-SetMainIcon   icons\rec_icon.ico
;@Ahk2Exe-SetFileVersion 1.0.0.0

#SingleInstance Force
#Requires AutoHotkey v2.0
#Include %A_ScriptDir%\libs\OCR.ahk
#Include %A_ScriptDir%\libs\ps_capture.ahk
; ============================================================
; 変数定義
; ============================================================

; ---------- 既定値 ----------
MatchHardTimeoutSec := 300
PollInterval := 500

Delay_AfterFirstConfirm := 500
Delay_AfterPlayKey      := 500
Delay_BeforeNavigate    := 180
Delay_AfterBackKey      := 4000
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
RolloverMode    := "instant"  ; "safe" / "instant"

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

CloseGameOnStop := false          ; 録画停止時にGAMEを終了する（オプション）
GameExitTimeoutMs := 5000         ; WinClose/Alt+F4 待機タイムアウト(ms)

; ---- ロード黒画面待ち（真っ黒検知で待機を延長）----
EnableLoadBlackWait    := true   ; 黒画面待機を有効にする
BlackDarknessThreshold := 32     ; 0-255 各チャンネルのしきい値（小さいほど“黒”判定が厳格）
BlackGrid              := 8      ; サンプリング格子（8→8x8=64点をチェック）
BlackBrightAllowance   := 5      ; “黒でない”点をこの数まで許容（ノイズ対策）
BlackCheckInterval     := 200    ; ポーリング間隔(ms)
BlackMinWait           := 500    ; 黒検出時の最低待機(ms)
BlackStableAfter       := 1500    ; “黒が明けた”後に進むまでの安定時間(ms)
BlackMaxWait           := 15000   ; 最大待機(ms) 上限
BlackGridX             := 8       ; 横方向サンプル数
BlackGridY             := 6       ; 縦方向サンプル数
BlackMinBlackRatio     := 0.70    ; この割合以上が黒なら「黒画面中」
; 画面の上下左右を何割クロップするか（UIの明るい縁を除外）
BlackCropMarginX       := 0.15
BlackCropMarginY       := 0.15

; ---- ログ関連 既定値 ----
AutoScrollLog := true

; ---- OCR 設定（既定値）----
OCRLang      := "ja-JP"           ; 日本語。英語UIなら "en-US"
OCRScale     := 3.0               ; 1.5～2.0 で精度↑（重くはなる）
OCRGray      := true              ; グレースケール化
; MatchCSVPath := A_ScriptDir "\match_results.csv"
MatchTextPath := A_ScriptDir "\match_results.txt"
IncludeRawInText := false          ; true でOCR全文を末尾に1行で添付
TextSeparator := "----------------------------------------"
; ---- テキスト出力 ----
MatchTextDir := A_ScriptDir "\results"  ; 出力フォルダ

; --- OCRデバッグ ---
UseOCRDebug := true   ; ハイライトと詳細ログを有効化（不要になったら false）

; --- 名前帯のROI係数（必要に応じてここだけ弄れば調整可能）---
NameTopFrac      := 0.26   ; ← 0.30/0.46 から上寄せ
NameHFrac        := 0.08   ; 少し厚めに
LeftNameX1Frac   := 0.10
LeftNameX2Frac   := 0.46
RightNameX1Frac  := 0.56
RightNameX2Frac  := 0.90

; ---- Slack ----
SlackEnabled   := false
SlackRouterUrl := ""
SlackTimeoutMs := 5000

; ===== パネル（結果ウィンドウ）全体の相対位置 =====
Clamp01(v) => Max(0, Min(1, v+0.0))

; --- 画面上のクライアント矩形を取得（DPI考慮, ピクセル）---
; --- PANEL基準の割合 → 画面ピクセル矩形 ---
ROI_PANEL := { x:0.06, y:0.08, w:0.88, h:0.78 }

; 右上の時刻は「再生回数」を含まないよう細めに
ROI_ReplayID := { n:"ReplayID", x1:0.15, y1:0.045, x2:0.33, y2:0.11 }
ROI_Time     := { n:"Time",     x1:0.67, y1:0.032, x2:0.82, y2:0.085 }

; プレイヤー名（クラブ名を避けるため“やや上寄せ＋薄め”）
ROI_L_Name   := { n:"L_Name",   x1:0.09, y1:0.275, x2:0.33, y2:0.335 }
ROI_R_Name   := { n:"R_Name",   x1:0.742, y1:0.275, x2:0.90, y2:0.335 }

; MR/LP（数字＋MR/LPがある行だけ）
ROI_L_Rating := { n:"L_Rating", x1:0.20, y1:0.190, x2:0.32, y2:0.24 }
ROI_R_Rating := { n:"R_Rating", x1:0.86, y1:0.190, x2:0.95, y2:0.24 }

; WIN/LOSE（中央の文字列。左右で重ならないように）
ROI_L_Result := { n:"L_Result", x1:0.33, y1:0.40,  x2:0.43, y2:0.46 }
ROI_R_Result := { n:"R_Result", x1:0.60, y1:0.40,  x2:0.70, y2:0.46 }


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

ResultSnapEnabled := true
ResultSnapDir := A_ScriptDir "\snapshots"

SaveOCREnabled := true
SaveOCRDir := A_ScriptDir "\results"

; ---------- 内部状態 ----------
global gRunning := false
global gPaused := false
global gRecording := false
global gUseFullROI := false
global gSafeStopRequested := false
global gRunStartTick := 0
global gRolloverRequested := false
global gLastRolloverTick := 0
global gLoopCount := 0
global gStatusLast := ""
global gLastLogText := ""
global gCurrentTextPath := ""           ; 現在の録画セグメントの出力ファイル
global gRecStartTick := 0               ; 録画開始Tick（経過時刻の基準）
global OCRDebugSaveWin := true


; ============================================================
; GUI（タブ式 + 共通フラット・ステータス）
; ============================================================
main := Gui("+Resize +MinSize720x250", "SF6 自動リプレイ")
main.SetFont("s9")
global IconPath := A_ScriptDir "\icons\rec_icon.ico"
if FileExist(IconPath) {
    ; main.SetIcon(IconPath)
    TraySetIcon(IconPath)
}

; ▼タブ本体：高さを少し低くして下段の常時ボタン領域を確保
tab := main.Add("Tab3", "x10 y10 w700 h410", ["基本設定","操作設定","出力設定","ログ","テスト","Slack"])

; -------------------- 基本設定タブ --------------------
tab.UseTab(1)
; ▼タブ内のグループの高さも調整（タブ高410に収まるように）
grpBasic := main.Add("GroupBox", "x20 y45 w680 h360", "基本設定")

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
chkRefocus := main.Add("CheckBox", "x300 y270 w230 +Wrap", "GUI/OBS後にゲームへ戻す")
chkRefocus.Value := AutoRefocusGame ? 1 : 0

chkChkGame := main.Add("CheckBox", "x35 y300 w240", "開始時にゲーム起動をチェック")
chkChkGame.Value := CheckOnStart_Game ? 1 : 0

chkChkOBS := main.Add("CheckBox", "x300 y295 w240", "開始時にOBS起動をチェック（OBS録画ON時）")
chkChkOBS.Value := CheckOnStart_OBS ? 1 : 0

chkCloseGame := main.Add("CheckBox", "x35 y330 w240", "録画停止時にゲームを終了する")
chkCloseGame.Value := CloseGameOnStop ? 1 : 0

; ▼操作（適用/読込/保存/OBS開始/停止）
btnApply := main.Add("Button", "x295 y355 w120 h28", "適用")
btnLoad  := main.Add("Button", "x35  y355 w120 h28", "読込（INI）")
btnSave  := main.Add("Button", "x165 y355 w120 h28", "保存（INI）")
btnOBSon := main.Add("Button", "x425 y355 w120 h28", "OBS録画開始")
btnOBSoff:= main.Add("Button", "x555 y355 w120 h28", "OBS録画停止")

; -------------------- 詳細設定(操作)タブ --------------------
tab.UseTab(2)

grpOps := main.Add("GroupBox", "x20 y45 w680 h360", "詳細設定（操作）")

; ── 実行（左側）
grpRun := main.Add("GroupBox", "x35 y70 w305 h250", "実行")
main.Add("Text", "x50  y95  w140", "次の選択方向（一覧）")
ddlDir := main.Add("DropDownList", "x50  y113 w120", ["down","up"])

main.Add("Text", "x195 y95  w140", "回数（0=無限）")
edtMatches := main.Add("Edit", "x195 y113 w120")

main.Add("Text", "x50  y147 w140", "タイムリミット（分）")
edtMaxMin := main.Add("Edit", "x50  y165 w120")

main.Add("Text", "x195 y147 w140", "録画ローテ（分）")
edtRollMin := main.Add("Edit", "x195 y165 w120")

main.Add("Text", "x50  y199 w140", "ローテ方式")
ddlRollMode := main.Add("DropDownList", "x50  y217 w120", ["safe","instant"])

main.Add("Text", "x195 y199 w140", "OBS切替キー（任意）")
edtToggle := main.Add("Edit", "x195 y217 w120", Key_ToggleRec)

chkUseToggle := main.Add("CheckBox", "x50  y250 w265", "ローテは切替キーで行う(OBS詳細設定)")
main.Add("Text", "x70  y270 w265", "オフの場合は録画停止→再開操作")
chkUseToggle.Value := UseOBSToggleForRollover ? 1 : 0
chkUseToggle.OnEvent("Click", (*) => (UseOBSToggleForRollover := (chkUseToggle.Value=1)))

; ── 検出 / 遷移（右側）
grpDetect := main.Add("GroupBox", "x355 y70 w330 h250", "検出 / 遷移")
main.Add("Text", "x370 y95  w140", "Tolerance（0-255）")
edtTol := main.Add("Edit", "x370 y113 w120")

chkROI := main.Add("CheckBox", "x535 y113 w125", "ROI = 全画面")

main.Add("Text", "x370 y147 w160", "次移動の回数 / 間隔(ms)")
edtNextRep := main.Add("Edit", "x370 y165 w50")
main.Add("Text", "x425 y168 w20 Center", "×")
edtNextInt := main.Add("Edit", "x450 y165 w70")

main.Add("Text", "x535 y147 w140", "開始→決定2回間隔")
edtD1 := main.Add("Edit", "x535 y165 w140", Delay_AfterFirstConfirm)

main.Add("Text", "x370 y199 w140", "決定後の小休止(ms)")
edtD2 := main.Add("Edit", "x370 y217 w140", Delay_AfterPlayKey)

main.Add("Text", "x535 y199 w140", "終了検出後の待機(ms)")
edtD3 := main.Add("Edit", "x535 y217 w140", Delay_BeforeNavigate)

main.Add("Text", "x370 y243 w140", "戻り後の待機(ms)")
edtD4 := main.Add("Edit", "x370 y261 w140", Delay_AfterBackKey)

main.Add("Text", "x535 y243 w140", "次のリプレイへ(ms)")
edtD5 := main.Add("Edit", "x535 y261 w140", Delay_BetweenItems)

; -------------------- 詳細設定(出力)タブ --------------------
tab.UseTab(3)

grpOut  := main.Add("GroupBox", "x20 y45 w680 h360", "詳細設定（出力）")
grpSave := main.Add("GroupBox", "x35 y70 w650 h260", "保存設定")

; ===== レイアウト基準 =====
; 左カラム: x=50 / 右カラム: x=400
; 入力幅は両カラムとも Edit=220, Button=60（隙間10）
; チェックボックスは Edit の下段に左右で揃える
; OCR は下段左から、チェックは右カラム位置で揃える

; 1) ログ（左カラム）
main.Add("Text",  "x50  y95  w130", "ログ出力フォルダ")
edtLogDir := main.Add("Edit",   "x50  y115 w220", LogDir)
btnLogDir := main.Add("Button", "x280 y115 w60",  "参照…")
btnLogDir.OnEvent("Click", (*) => (
    dir := DirSelect(LogDir, 1, "ログ出力フォルダの選択"),
    dir ? (edtLogDir.Value := dir, LogDir := dir) : ""
))
chkLog := main.Add("CheckBox", "x50  y150 w270", "ログをファイルに保存")
chkLog.Value := LogEnabled ? 1 : 0
chkLog.OnEvent("Click", (*) => (
    LogEnabled := (chkLog.Value=1),
    edtLogDir.Enabled := LogEnabled,
    btnLogDir.Enabled := LogEnabled
))
edtLogDir.Enabled := LogEnabled, btnLogDir.Enabled := LogEnabled

; 2) キャプチャ（右カラム）
main.Add("Text",  "x380 y95  w130", "キャプチャ保存先")
edtSnapDir := main.Add("Edit",   "x380 y115 w220", ResultSnapDir)
btnSnapDir := main.Add("Button", "x610 y115 w60",  "参照…")
btnSnapDir.OnEvent("Click", (*) => (
    dir := DirSelect(ResultSnapDir, 1, "キャプチャ保存先の選択"),
    dir ? (edtSnapDir.Value := dir, ResultSnapDir := dir) : ""
))
chkSnap := main.Add("CheckBox", "x380 y150 w270", "マッチリザルト画面をキャプチャ保存")
chkSnap.Value := ResultSnapEnabled ? 1 : 0
chkSnap.OnEvent("Click", (*) => (
    ResultSnapEnabled := (chkSnap.Value=1),
    edtSnapDir.Enabled := ResultSnapEnabled,
    btnSnapDir.Enabled := ResultSnapEnabled
))
edtSnapDir.Enabled := ResultSnapEnabled, btnSnapDir.Enabled := ResultSnapEnabled

; 3) リザルト（OCR）— 下段（左=パス、右=ON/OFF）
main.Add("Text",  "x50  y190 w130", "OCR保存先")
edtOCRDir := main.Add("Edit",   "x50  y210 w220", SaveOCRDir)
btnOCRDir := main.Add("Button", "x280 y210 w60",  "参照…")
btnOCRDir.OnEvent("Click", (*) => (
    dir := DirSelect(SaveOCRDir, 1, "OCR結果の保存先の選択"),
    dir ? (edtOCRDir.Value := dir, SaveOCRDir := dir) : ""
))
chkOCR := main.Add("CheckBox", "x50 y245 w270", "OCRのマッチリザルトを保存")
chkOCR.Value := SaveOCREnabled ? 1 : 0
chkOCR.OnEvent("Click", (*) => (
    SaveOCREnabled := (chkOCR.Value=1),
    edtOCRDir.Enabled := SaveOCREnabled,
    btnOCRDir.Enabled := SaveOCREnabled
))
edtOCRDir.Enabled := SaveOCREnabled, btnOCRDir.Enabled := SaveOCREnabled

; -------------------- ログタブ --------------------
tab.UseTab(4)
grpLog := main.Add("GroupBox", "x20 y45 w680 h360", "ログ")
logBox := main.Add("Edit", "x35 y70 w650 h330 ReadOnly -Wrap +VScroll +HScroll", "")

btnLogClear := main.Add("Button", "x425 y40 w100 h26", "表示クリア")
chkAutoScroll := main.Add("CheckBox", "x545 y40 w140 h26", "最新へ自動スクロール")
chkAutoScroll.Value := AutoScrollLog ? 1 : 0

btnLogClear.OnEvent("Click", (*) => (
    logBox.Value := "",
    TrayTip("ログ", "表示をクリアしました", 900)
))

; -------------------- テストタブ --------------------
tab.UseTab(5)
grpTest := main.Add("GroupBox", "x20 y45 w680 h360", "テスト")
btnDetect := main.Add("Button", "x35 y80 w150 h30", "マッチ終了検出テスト")
btnTestBlack := main.Add("Button", "x35 y120 w150 h30", "黒画面待機テスト")
btnOCRTest := main.Add("Button", "x35 y160 w150 h30", "リザルトOCRテスト")
btnTestName := main.Add("Button", "x35 y200 w150 h30", "リザルトOCRテスト(詳細)")

; -------------------- Slackタブ --------------------
tab.UseTab(6)
grpSlack := main.Add("GroupBox", "x20 y45 w680 h360", "Slack")
; --- 有効/無効 ---
chkSlackEnabled := main.Add("CheckBox"
    , "x35 y80 vUI_SlackEnabled"
    , "開始／録画切り替え／終了時のSlack通知を有効化")

; --- Router URL ---
main.Add("Text", "x35 y120", "Router URL")
edtSlackRouter := main.Add("Edit"
    , "x130 y116 w540 vUI_SlackRouterUrl")

    ; デプロイ案内（クリックでGitHubを開く）
main.Add("Text", "x35 y250", "Slack通知を使うには slack-message-router を事前にデプロイしてください：")
main.Add("Link", "x35 y275 w640", '<a href="https://github.com/mmiyaji/slack-message-router">https://github.com/mmiyaji/slack-message-router</a>')

; --- Timeout ---
main.Add("Text", "x35 y160", "Timeout (ms)")
edtSlackTimeout := main.Add("Edit"
    , "x130 y156 w120 Number vUI_SlackTimeoutMs")
main.Add("Text", "x260 y160", "例: 2000")

; --- テスト送信 ---
btnSlackTest := main.Add("Button"
    , "x35 y210 w150 h30"
    , "Slackテスト送信")
btnSlackTest.OnEvent("Click", SlackTestSend)

; --- ON/OFF連動（初期状態） ---
UpdateSlackUIState()

UpdateSlackUIState() {
    global chkSlackEnabled, edtSlackRouter, edtSlackTimeout
    enabled := chkSlackEnabled.Value
    edtSlackRouter.Enabled := enabled
    edtSlackTimeout.Enabled := enabled
}
chkSlackEnabled.OnEvent("Click", (*) => UpdateSlackUIState())
SlackTestSend(*) {
    global chkSlackEnabled, edtSlackRouter, edtSlackTimeout
    global SlackEnabled, SlackRouterUrl, SlackTimeoutMs

    ; UI → 変数に反映（保存前でもテスト可能）
    SlackEnabled   := chkSlackEnabled.Value
    SlackRouterUrl := Trim(edtSlackRouter.Value)
    SlackTimeoutMs := Integer(edtSlackTimeout.Value)

    if (!SlackEnabled) {
        MsgBox("Slack通知が無効です", "テスト送信", 48)
        return
    }
    if (SlackRouterUrl = "") {
        MsgBox("Router URL が未設定です", "テスト送信", 48)
        return
    }

    SlackNotify("✅ sf6autoreplay Slack通知テスト`n" A_ComputerName, "info")
}

; -------------------- タブ終了 --------------------
tab.UseTab()

main.OnEvent("Close", (*) => ExitApp())

main.OnEvent("Size", OnMainResize)
OnMainResize(gui, minMax, w, h) {
    margin := 10
    bottomBarH := 70

    ; タブ全体を拡張
    tab.Move(margin, margin, w - margin*2, h - bottomBarH - margin*2)

    ; ログタブの領域（存在すれば動かす）
    try {
        grpLog.Move(20, 45, w - 40, h - bottomBarH - 55)
        logBox.Move(35, 70, w - 70, h - bottomBarH - 95)
        btnLogClear.Move(w - 280, 40)
        chkAutoScroll.Move(w - 170, 40)
    }

    ; 下段のボタン群
    btnY := h - bottomBarH + 5
    gap := 10, bw := 120, bh := 28
    btnStart.Move( 15,                btnY, bw, bh)
    btnSafe.Move(  15 + (bw+gap),     btnY, bw, bh)
    btnForce.Move( 15 + 2*(bw+gap),   btnY, bw, bh)
    btnPause.Move( 15 + 3*(bw+gap),   btnY, bw, bh)

    statusText.Move(15, btnY + 35, w - 30, 24)
}

; ▼タブ外（常時表示）操作ボタン：どのタブでも使える
; 位置はリサイズイベントで追随させるので、初期値は仮でOK
btnStart := main.Add("Button", "x15  y430 w120 h28", "開始")
btnSafe  := main.Add("Button", "x145 y430 w120 h28", "安全停止")
btnForce := main.Add("Button", "x275 y430 w120 h28", "即時停止")
btnPause := main.Add("Button", "x405 y430 w120 h28", "一時停止")

; ▼共通ステータス（タブ外）
statusText := main.Add("Text", "x15 y465 w700 h24 vStatusText", "")
statusText.SetFont("s9", "Segoe UI")
SetStatus("準備完了")

; ---- 画面初期化 ----
if FileExist(ConfigPath) {
    LoadConfig(ConfigPath)
}
UpdateGuiFromVars()
InitOCR()
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
btnApply.OnEvent("Click", (*) => (ApplyGuiToVars()))
btnLoad.OnEvent("Click",  (*) => (LoadConfig(ConfigPath), UpdateGuiFromVars(), TrayTip("読込","設定を読み込みました",900)))
btnSave.OnEvent("Click",  (*) => (ApplyGuiToVars(), SaveConfig(ConfigPath), TrayTip("保存","設定を保存しました",900)))

btnDetect.OnEvent("Click",(*) => (QuickDetectTest(), RefocusGame()))
btnOBSon.OnEvent("Click", (*) => (FocusedTriggerOBS(Key_StartRec), RefocusGame()))
btnOBSoff.OnEvent("Click",(*) => (FocusedTriggerOBS(Key_StopRec),  RefocusGame()))
btnTestBlack.OnEvent("Click", (*) => (RefocusGame(), TestBlackWait()))
btnOCRTest.OnEvent("Click", (*) => (RefocusGame(), OCR_TestButton()))
btnTestName.OnEvent("Click", (*) => (RefocusGame(), OCR_TestResultButton(GameWinSelector)))

; ============================================================
; ホットキー
; ============================================================
; GUI 開閉ホットキー
^!g:: (main.Visible := !main.Visible, !main.Visible ? "" : RefocusGame())
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
; メインロジック
; ============================================================

;-- 関数: StartAutomation()
;   目的: 開始の処理を開始する。
;   引数/返り値: 定義参照
StartAutomation() {
    global gRunning, gPaused, gRecording, TotalMatches
    global gSafeStopRequested, gRunStartTick
    global RolloverMinutes, RolloverMode, gRolloverRequested, gLastRolloverTick
    global MaxRunMinutes, gLoopCount
    global UseOBSRecording, CheckOnStart_Game, CheckOnStart_OBS
    global CloseGameOnStop, GameExitTimeoutMs

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
    SlackNotify(BuildSlackStartMessage(), "info")

    ; 録画開始
    if UseOBSRecording && !gRecording {
        FocusedTriggerOBS(Key_StartRec)
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
                    SlackNotify("🎥 OBS rollover (instant) at elapsed=" FormatDuration(A_TickCount - gRunStartTick), "info")
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
        EnsureFocusGame()
        Press(Key_Confirm, 80)
        Sleep Delay_AfterFirstConfirm
        EnsureFocusGame()
        if SaveOCREnabled {
            try {
                Sleep 150  ; わずかに安定待ち
                OCR_RecordCurrentMatch(GameWinSelector)
                
            } catch as e {
                Log("OCR: failed - " e.Message)
            }
        } else {
            Log("OCR: skipped (disabled)")
        }
        Press(Key_Confirm, 80)
        Sleep Delay_AfterPlayKey

        ; 終了UI待ち
        startTick := A_TickCount
        detectHardTimeout := false
        Loop {
            if !gRunning
                break
            while gPaused && gRunning
                Sleep 150
            if !gRunning
                break

            if (A_TickCount - startTick) > (MatchHardTimeoutSec * 1000) {
                Log("TIMEOUT: no end UI within " MatchHardTimeoutSec "s -> force handling same as detected")
                detectHardTimeout := true
            } else {
                detectHardTimeout := false
            }

            roi := (gUseFullROI ? GetROI_Full() : GetROI_End_Default(GameWinSelector))
            if FindAnyImage(Img_Ends, roi, ToleranceEnd, &fx, &fy) or detectHardTimeout{
                Log("DETECT: end UI at " fx "," fy)
                Sleep Delay_BeforeNavigate

                ; 終了UIで戻る操作 ※直前にフォーカス
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
                        , 0.30, 0.30)  ; 中央30%×30%をサンプリング
                }
                if gRolloverRequested && !gSafeStopRequested && UseOBSRecording {
                    RolloverOBS("safe")                 ; Ctrl+F9 1回
                    gRolloverRequested := false
                    SlackNotify("🎥 OBS rollover (safe) after match #" gLoopCount, "info")
                }
                break
            }
            Sleep PollInterval
        }

        if gSafeStopRequested
            break

        ; 次のリプレイへ ※直前にフォーカス
        EnsureFocusGame()
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
    ; 録画停止時にゲームを終了（オプション）
    if CloseGameOnStop {
        Log("CLOSE: option enabled, trying to close GAME...")
        CloseGameApp()
    }
    gRunning := false
    gSafeStopRequested := false
    gRolloverRequested := false
    Log("END: automation")

    ; --- stop reason 判定（ループを抜けた後、終了処理の前後どちらでもOK）
    stopReason := "unknown"
    if (TotalMatches > 0 && gLoopCount >= TotalMatches)
        stopReason := "reached_total"
    else if (gSafeStopRequested)
        stopReason := "safe_stop"
    else if (!gRunning)
        stopReason := "stopped"

    ; 終了ログの直後などで
    SlackNotify(BuildSlackEndMessage(stopReason), (stopReason="safe_stop" ? "warn" : "info"))
    Log("SEND: notify")
}

FormatDuration(ms) {
    totalSec := Floor(ms / 1000)
    h := Floor(totalSec / 3600)
    m := Floor(Mod(totalSec, 3600) / 60)
    s := Mod(totalSec, 60)
    if (h > 0)
        return h "h" m "m" s "s"
    else if (m > 0)
        return m "m" s "s"
    else
        return s "s"
}

BuildSlackStartMessage() {
    global TotalMatches, MaxRunMinutes, RolloverMinutes, RolloverMode, UseOBSRecording
    global GameWinSelector, OBSWinSelector

    total := (TotalMatches > 0) ? TotalMatches : "∞"
    maxrun := (MaxRunMinutes > 0) ? (MaxRunMinutes "min") : "∞"
    rollover := (UseOBSRecording && RolloverMinutes > 0) ? (RolloverMinutes "min/" RolloverMode) : "off"

    msg :=  ( "🎮 sf6autoreplay START`n"
    . "total=" total ", maxRun=" maxrun ", rollover=" rollover ", obs=" (UseOBSRecording ? "on" : "off") "`n"
    . "gameSel=" GameWinSelector )

    if (UseOBSRecording)
        msg .= "`nobsSel=" OBSWinSelector

    return msg
}

BuildSlackEndMessage(stopReason) {
    global gLoopCount, TotalMatches, gRunStartTick, gSafeStopRequested
    global UseOBSRecording, gRecording, RolloverMinutes, RolloverMode

    total := (TotalMatches > 0) ? TotalMatches : "∞"
    elapsed := FormatDuration(A_TickCount - gRunStartTick)

    msg := ( "🛑 sf6autoreplay END`n"
    . "reason=" stopReason "`n"
    . "done=" gLoopCount "/" total ", elapsed=" elapsed )

    ; 状態も少しだけ載せる（デバッグに効く）
    msg .= "`nobs=" (UseOBSRecording ? "on" : "off") ", rec=" (gRecording ? "on" : "off")
    if (UseOBSRecording && RolloverMinutes > 0)
        msg .= ", rollover=" RolloverMinutes "min/" RolloverMode

    return msg
}

; ===============================
; 整形テキスト出力
; ===============================
AppendMatchText2(dt, id, mode
    , leftName, leftVal, leftKind, rightName, rightVal, rightKind
    , winnerSide, raw) {
    global MatchTextPath, IncludeRawInText, TextSeparator

    dt   := (dt   != "" ? dt   : FormatTime(A_Now, "yyyy/MM/dd HH:mm"))
    mode := (mode != "" ? mode : "-")

    leftTag  := (winnerSide="left"  ? "[WIN]" : (winnerSide="right" ? "[LOSE]" : "[?]"))
    rightTag := (winnerSide="right" ? "[WIN]" : (winnerSide="left"  ? "[LOSE]" : "[?]"))

    winnerName := (winnerSide="left"  ? leftName  : (winnerSide="right" ? rightName : ""))
    winnerVal  := (winnerSide="left"  ? leftVal   : (winnerSide="right" ? rightVal  : ""))
    winnerKind := (winnerSide="left"  ? leftKind  : (winnerSide="right" ? rightKind : ""))
    loserName  := (winnerSide="left"  ? rightName : (winnerSide="right" ? leftName  : ""))
    loserVal   := (winnerSide="left"  ? rightVal  : (winnerSide="right" ? leftVal   : ""))
    loserKind  := (winnerSide="left"  ? rightKind : (winnerSide="right" ? leftKind  : ""))

    text := TextSeparator "`r`n"
          . dt "  " mode "`r`n"
          . "Replay ID: " id "`r`n"
          . "Winner: " FormatNameRating(winnerName, winnerVal, winnerKind) "`r`n"
          . "Loser : " FormatNameRating(loserName,  loserVal,  loserKind)  "`r`n"
          . "Left  : " FormatNameRating(leftName,  leftVal,  leftKind)  " " leftTag  "`r`n"
          . "Right : " FormatNameRating(rightName, rightVal, rightKind) " " rightTag "`r`n"

    if IncludeRawInText
        text .= "Raw  : " RegExReplace(raw, "\R", " ") "`r`n"

    text .= "`r`n"
    FileAppend(text, MatchTextPath, "UTF-8-RAW")
    Log("OCR: wrote formatted text -> " MatchTextPath)
}

; フォルダ作成＋録画ごとの一意ファイルを作る
; ミリ秒→ 00:00:00 形式
; 1行：「経過  [W/?/L] LeftName (1899 MR) vs [W/?/L] RightName (1716 LP) [ReplayID]」
AppendMatchLine(replayID
    , leftName, leftVal, leftKind
    , rightName, rightVal, rightKind
    , winnerSide) {

    global gCurrentTextPath, gRecStartTick

    if (gCurrentTextPath = "")
        StartNewRecordingTextFile("auto")  ; 念のため

    elapsed  := MsToHMS(A_TickCount - gRecStartTick)

    ; 検出結果に合わせて左右にタグを付与
    leftTag  := (winnerSide="left"  ? "[W]" : (winnerSide="right" ? "[L]" : "[?]"))
    rightTag := (winnerSide="right" ? "[W]" : (winnerSide="left"  ? "[L]" : "[?]"))

    ; 名前+レーティング（値が無い時は名前のみ）
    leftDisp  := FormatNameRating( leftName,  leftVal,  (leftKind!=""  ? leftKind  : "MR") )
    rightDisp := FormatNameRating( rightName, rightVal, (rightKind!="" ? rightKind : "MR") )

    line := elapsed "  "
          . leftTag  " " leftDisp
          . " vs "
          . rightTag " " rightDisp
          . (replayID!="" ? " [" replayID "]" : "")
          . "`r`n"

    FileAppend(line, gCurrentTextPath, "UTF-8-RAW")
    Log("TEXT: " RegExReplace(line, "\R", " "))
}

; ROI 内が「ほぼ黒」かを判定（格子サンプリング）
; ROI内の黒率(0.0-1.0)を返す（サンプル点は格子状）
; ROIが「黒率 >= minRatio」の間は待機。明けたら安定時間だけ余分に待ってから復帰。
WaitWhileBlackByRatio_Window(winSel
    , darkness := 32, minRatio := 0.70, gx := 8, gy := 6
    , pollMs := 120, minWait := 500, stableAfter := 400, maxWait := 8000
    , fracX := 0.30, fracY := 0.30) {

    roi := GetROI_Load_WindowCenter(winSel, fracX, fracY)
    t0 := A_TickCount, sawBlack := false

    ; 黒の間待つ
    Loop {
        ratio := BlackRatio(roi, darkness, gx, gy)
        if (ratio >= minRatio) {
            Log(Format("ロード画面検知: {:.0f}% (しきい値{:.0f}%)", ratio*100, minRatio*100))
            sawBlack := true
        } else {
            Log(Format("ロード画面終了検知: {:.0f}% (しきい値{:.0f}%)", ratio*100, minRatio*100))
            break
        }
        if (A_TickCount - t0) >= maxWait
            break
        Sleep pollMs
    }

    if (sawBlack) {
        ; 最低待機を保証
        remain := minWait - (A_TickCount - t0)
        if (remain > 0)
            Sleep remain
        ; 黒明け後の安定化
        okStart := A_TickCount
        Loop {
            ratio := BlackRatio(roi, darkness, gx, gy)
            if (ratio >= minRatio)
                okStart := A_TickCount   ; まだ黒 → 安定カウントをリセット
            if (A_TickCount - okStart) >= stableAfter
                break
            if (A_TickCount - t0) >= maxWait
                break
            Sleep pollMs
        }
    }
}

; 黒画面の間は待機し、明けてから安定時間までは更に待つ
WaitWhileBlack(roi, darkness := 32, grid := 8, brightAllowance := 5
    , pollMs := 120, minWait := 500, stableAfter := 400, maxWait := 8000) {

    t0 := A_TickCount
    seenBlack := false

    ; まず黒なら最低待機を確保しつつ黒が明けるのを待つ
    Loop {
        isBlack := IsRegionMostlyBlack(roi, darkness, grid, brightAllowance)
        if (isBlack) {
            seenBlack := true
        } else {
            break
        }
        if (A_TickCount - t0) >= maxWait
            break
        Sleep pollMs
    }

    ; 一度でも黒を見ていたら最低待機保証
    if (seenBlack) {
        ; minWait まで届いていなければ追加で待つ
        remain := minWait - (A_TickCount - t0)
        if (remain > 0)
            Sleep remain

        ; さらに“黒が明けた後”に少し安定時間を入れる
        okStart := A_TickCount
        Loop {
            if IsRegionMostlyBlack(roi, darkness, grid, brightAllowance) {
                okStart := A_TickCount  ; また黒くなった → 安定計測をリセット
            }
            if (A_TickCount - okStart) >= stableAfter
                break
            if (A_TickCount - t0) >= maxWait
                break
            Sleep pollMs
        }
    }
}

;=============================================================================
; [ブロック] 起動チェック/初期化
; 説明: 起動時の環境チェックや初期化処理。
;=============================================================================

;-- 関数: InitOCR()
;   目的: OCRを初期化する。
;   引数/返り値: 定義参照
InitOCR() {
    ; シングルトン初期化: OCRが有効化されている時だけ初期化し、二重実行を避ける
    global SaveOCREnabled, OCRLang
    static _ready := false
    static _sig := ""  ; 設定のシグネチャ（例: 言語）

    ; OCRがOFFなら何もしない
    if !SaveOCREnabled
        return false

    ; 現在の設定からシグネチャを生成（必要に応じて項目を追加）
    sig := StrLower(Trim(OCRLang ""))

    ; 初回 or 設定変更時のみ初期化
    if (!_ready || sig != _sig) {
        try {
            ; 例: 言語ロード。必要に応じて他の初期化処理を追加
            OCR.LoadLanguage(OCRLang)
            _sig := sig
            _ready := true
            try Log("InitOCR: initialized (lang=" OCRLang ")")
        } catch Error as e {
            _ready := false
            try Log("InitOCR: failed -> " e.Message)
        }
    }
    return _ready
}

;=============================================================================
; [ブロック] GUI/ウィンドウ・ログ表示
; 説明: GUI構築、ステータスやログ出力などの表示処理。
;=============================================================================

;-- 関数: ApplyGuiToVars()
;   目的: GUIに関する処理を行う。
;   引数/返り値: 定義参照
ApplyGuiToVars() {
    global NextDirection, TotalMatches, MaxRunMinutes, RolloverMinutes, RolloverMode
    global ToleranceEnd, gUseFullROI, NextRepeats, NextIntervalMs
    global Key_StartRec, Key_StopRec, Key_ToggleRec, OBSWinSelector, Img_Ends
    global GameWinSelector, AutoRefocusGame, UseOBSRecording, UseOBSToggleForRollover, CheckOnStart_Game, CheckOnStart_OBS
    global CloseGameOnStop, GameExitTimeoutMs
    global LogEnabled, LogDir, AutoScrollLog
    global ResultSnapEnabled, ResultSnapDir
    global SaveOCREnabled, SaveOCRDir
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
    CloseGameOnStop := !!chkCloseGame.Value
    LogEnabled := !!chkLog.Value
    LogDir := edtLogDir.Text
    AutoScrollLog := !!chkAutoScroll.Value
    ResultSnapEnabled := !!chkSnap.Value
    ResultSnapDir := edtSnapDir.Text
    SaveOCREnabled := !!chkOCR.Value
    SaveOCRDir := edtOCRDir.Text
}
;-- 関数: BuildStatusBase()
;   目的: UIを組み立てる。
;   引数/返り値: 定義参照
BuildStatusBase() {
    global gRunning, gPaused, gRecording, gRunStartTick, gLoopCount, UseOBSRecording
    runMin := (gRunStartTick>0) ? Round((A_TickCount - gRunStartTick)/60000, 1) : 0
    return "状態: " (gRunning ? (gPaused ? "□一時停止中" : "●実行中") : "■停止")
        . " / 録画: " (UseOBSRecording ? (gRecording ? "●ON" : "■OFF") : "- 未使用")
        . " / 経過: " runMin "分"
        . " / ループ: " gLoopCount
}

;-- 関数: CurrentStatusText()
;   目的: ステータスに関する処理を行う。
;   引数/返り値: 定義参照
CurrentStatusText() {
    global gLastLogText
    base := BuildStatusBase()
    if (gLastLogText != "")
        base .= "  |  最新: " TruncForStatus(gLastLogText, 60)  ; 長いときは丸め
    return base
}
;-- 関数: GetROI_Load_WindowCenter(winSel, fracX := 0.30, fracY := 0.30)
;   目的: ウィンドウを読み込む。
;   引数/返り値: 定義参照
GetROI_Load_WindowCenter(winSel, fracX := 0.30, fracY := 0.30) {
    ; クライアント領域の画面座標を取得（AHK v2）
    if !WinExist(winSel)
        throw Error("Window not found: " winSel)
    WinGetClientPos(&cx, &cy, &cw, &ch, winSel)  ; 画面座標系のクライアント矩形
    if (cw <= 0 || ch <= 0)
        throw Error("Invalid client size.")

    w := Max(1, Round(cw * fracX))
    h := Max(1, Round(ch * fracY))
    x1 := cx + Round((cw - w) / 2)
    y1 := cy + Round((ch - h) / 2)
    x2 := x1 + w
    y2 := y1 + h
    return {x1:x1, y1:y1, x2:x2, y2:y2}
}

;-- 関数: SetStatus(text)
;   目的: ステータスに関する処理を行う。
;   引数/返り値: 定義参照
SetStatus(text) {
    global statusText, gStatusLast
    if !IsSet(statusText) || !statusText
        return
    if (text = gStatusLast)
        return  ; 同一なら再描画しない
    statusText.Value := text
    gStatusLast := text
}
;-- 関数: TruncForStatus(msg, maxChars := 60)
;   目的: ステータスに関する処理を行う。
;   引数/返り値: 定義参照
TruncForStatus(msg, maxChars := 60) {
    msg := RegExReplace(msg, "\R", " ")
    return (StrLen(msg) > maxChars) ? SubStr(msg, 1, maxChars-1) "…" : msg
}
;-- 関数: UpdateGuiFromVars()
;   目的: GUIを更新する。
;   引数/返り値: 定義参照
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
    chkCloseGame.Value := CloseGameOnStop ? 1 : 0
    chkLog.Value := LogEnabled ? 1 : 0
    edtLogDir.Text := LogDir
    chkSnap.Value := ResultSnapEnabled ? 1 : 0
    edtSnapDir.Text := ResultSnapDir
    chkOCR.Value := SaveOCREnabled ? 1 : 0
    edtOCRDir.Text := SaveOCRDir
    chkAutoScroll.Value := AutoScrollLog ? 1 : 0
    UpdatePauseBtn()
}

;-- 関数: UpdateStatusText()
;   目的: ステータスを更新する。
;   引数/返り値: 定義参照
UpdateStatusText() {
    ; いまの状態 + 最新ログ を合成して出す
    SetStatus(CurrentStatusText())
    UpdatePauseBtn()  ; ラベル差分更新（既存の関数）
}


;=============================================================================
; [ブロック] OBS/録画制御
; 説明: OBS連携による録画開始・停止・状態監視。
;=============================================================================

;-- 関数: FocusedTriggerOBS(keyToSend)
;   目的: OBSに関する処理を行う。
;   引数/返り値: 定義参照
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
        ; OBS操作後は自動でGAMEへ戻す（設定に従う）
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

;-- 関数: RolloverOBS(mode := "instant")
;   目的: OBSに関する処理を行う。
;   引数/返り値: 定義参照
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
    StartNewRecordingTextFile("rollover")
    TrayTip "ローテ", "録画ファイルを切替（" (mode="instant"?"即時":"試合間") "）", 1200
}

;-- 関数: SendOBSTest()
;   目的: OBSに関する処理を行う。
;   引数/返り値: 定義参照
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

;-- 関数: StartNewRecordingTextFile(reason := "start")
;   目的: 録画を開始する。
;   引数/返り値: 定義参照
StartNewRecordingTextFile(reason := "start") {
    global MatchTextDir, gCurrentTextPath, gRecStartTick
    try DirCreate(MatchTextDir)
    ts := FormatTime(A_Now, "yyyyMMdd_HHmmss")
    gCurrentTextPath := MatchTextDir "\sf6_" ts ".txt"
    gRecStartTick := A_TickCount
    Log("TEXT: new output file -> " gCurrentTextPath " [" reason "]")
}


;=============================================================================
; [ブロック] OCR/画面認識
; 説明: OCRや画像文字認識に関連する処理。
;=============================================================================

;-- 関数: OCR_RecordCurrentMatch(winSel, showHighlight := false)
;   目的: OCRに関する処理を行う。
;   引数/返り値: 定義参照
OCR_RecordCurrentMatch(winSel, showHighlight := false, test := false) {
    global OCRLang, OCRScale, OCRGray
    global ROI_ReplayID, ROI_Time, ROI_L_Name, ROI_R_Name
    global ROI_L_Rating, ROI_R_Rating, ROI_L_Result, ROI_R_Result

    global gCurrentTextPath, gRecStartTick
    if (gCurrentTextPath = "") and (test = false)
        StartNewRecordingTextFile("auto")  ; 念のため

    if (OCRDebugSaveWin)
        CaptureWithPSCapture(winSel, test ? "test" : "", "", "client")
    ; if (OCRDebugSaveWin) and (test = false)
    ;     CaptureWithMiniCap(winSel, "", "", "client")
    ; CaptureWithNirCmd(winSel)
    ; if (OCRDebugSaveWin)
    ;     CaptureWithNirCmd_Client(winSel)

    opts := {lang:OCRLang, scale:Max(2.5, OCRScale), grayscale:OCRGray}
    opts_en := {lang:OCRLang, scale:Max(2.5, OCRScale), grayscale:OCRGray}

    ; 1) 各ROIからOCR
    idR := RectFromFrac(winSel, ROI_ReplayID),   resID := OCR.FromRect(idR.x,idR.y,idR.w,idR.h,opts_en)
    tmR := RectFromFrac(winSel, ROI_Time),       resTM := OCR.FromRect(tmR.x,tmR.y,tmR.w,tmR.h,opts_en)
    lnR := RectFromFrac(winSel, ROI_L_Name),     resLN := OCR.FromRect(lnR.x,lnR.y,lnR.w,lnR.h,opts)
    rnR := RectFromFrac(winSel, ROI_R_Name),     resRN := OCR.FromRect(rnR.x,rnR.y,rnR.w,rnR.h,opts)
    lrR := RectFromFrac(winSel, ROI_L_Rating),   resLR := OCR.FromRect(lrR.x,lrR.y,lrR.w,lrR.h,opts_en)
    rrR := RectFromFrac(winSel, ROI_R_Rating),   resRR := OCR.FromRect(rrR.x,rrR.y,rrR.w,rrR.h,opts_en)
    lwR := RectFromFrac(winSel, ROI_L_Result),   resLW := OCR.FromRect(lwR.x,lwR.y,lwR.w,lwR.h,opts_en)
    rwR := RectFromFrac(winSel, ROI_R_Result),   resRW := OCR.FromRect(rwR.x,rwR.y,rwR.w,rwR.h,opts_en)

    idTxt := RegExReplace(OCR_Normalize(resID.Text), "\s+")
    timeTxt := resTM.Text
    lName   := OCR_ExtractNameSmart(resLN.Text)
    rName   := OCR_ExtractNameSmart(resRN.Text)
    lRate   := OCR_ExtractRating(resLR.Text)     ; -> {value, kind("MR"/"LP"/"")}
    rRate   := OCR_ExtractRating(resRR.Text)
    lf      := OCR_GetSideFlags(resLW.Text)      ; -> {win:bool, lose:bool}
    rf      := OCR_GetSideFlags(resRW.Text)

    winnerSide := ""
    if (lf.win || rf.lose)
        winnerSide := "left"
    else if (rf.win || lf.lose)
        winnerSide := "right"

    ; 2) ログ
    Log(Format(
        "OCR: ID={1} | TM={2} | LN={3} | RN={4} | LR={5} {6} | RR={7} {8} | winnerSide={9}"
    , RegExReplace(idTxt,  "\R"," ")
    , RegExReplace(timeTxt,"\R"," ")
    , lName
    , rName
    , lRate.value, (lRate.kind!="" ? lRate.kind : "MR")
    , rRate.value, (rRate.kind!="" ? rRate.kind : "MR")
    , winnerSide
    ))

    ; 3) デバッグハイライト（空のときは何もしない安全版）
    if showHighlight {
        SafeHighlight(resID,  800, "Gold")
        SafeHighlight(resTM,  800, "SkyBlue")
        SafeHighlight(resLN,  800, "HotPink")
        SafeHighlight(resRN,  800, "Orange")
        SafeHighlight(resLR,  800, "Lime")
        SafeHighlight(resRR,  800, "Aqua")
        SafeHighlight(resLW,  800, "Silver")
        SafeHighlight(resRW,  800, "Gray")
    }

    ; 4) 書き出し（左右の抽出順は維持、勝敗タグだけ反映）
    ;    ※ AppendMatchLine は既存のものをそのまま使用
    if (test = false) {
        AppendMatchLine(
            OCR_ExtractReplayID(idTxt)                 ; replayID
        , lName,  lRate.value,  (lRate.kind!=""?lRate.kind:"MR")
        , rName,  rRate.value,  (rRate.kind!=""?rRate.kind:"MR")
        , winnerSide
        )
    }

    ; 5) 呼び出し側でも使えるよう返却
    return {
        replayID:    idTxt
      , dateTime:    OCR_ExtractDateTime(timeTxt)
      , leftName:    lName,    rightName:  rName
      , leftRating:  lRate,    rightRating: rRate
      , winnerSide:  winnerSide
      , raw: { id:idTxt, tm:timeTxt, ln:resLN.Text, rn:resRN.Text, lr:resLR.Text, rr:resRR.Text, lw:resLW.Text, rw:resRW.Text }
    }
}

;-- 関数: AppendMatchText(dt, id, mode, leftName, leftMR, rightName, rightMR, winnerSide, raw)
;   目的: テキストに関する処理を行う。
;   引数/返り値: 定義参照
AppendMatchText(dt, id, mode, leftName, leftMR, rightName, rightMR, winnerSide, raw) {
    global MatchTextPath, IncludeRawInText, TextSeparator

    dt   := (dt   != "" ? dt   : FormatTime(A_Now, "yyyy/MM/dd HH:mm"))
    mode := (mode != "" ? mode : "-")

    leftTag  := (winnerSide="left"  ? "[WIN]" : (winnerSide="right" ? "[LOSE]" : "[?]"))
    rightTag := (winnerSide="right" ? "[WIN]" : (winnerSide="left"  ? "[LOSE]" : "[?]"))

    winner   := (winnerSide="left"  ? leftName  : (winnerSide="right" ? rightName : ""))
    loser    := (winnerSide="left"  ? rightName : (winnerSide="right" ? leftName  : ""))
    winnerMR := (winnerSide="left"  ? leftMR    : (winnerSide="right" ? rightMR   : ""))
    loserMR  := (winnerSide="left"  ? rightMR   : (winnerSide="right" ? leftMR    : ""))

    text := TextSeparator "`r`n"
          . dt "  " mode "`r`n"
          . "Replay ID: " id "`r`n"
          . "Winner: " FormatNameMR(winner, winnerMR) "`r`n"
          . "Loser : " FormatNameMR(loser,  loserMR)  "`r`n"
          . "Left  : " FormatNameMR(leftName,  leftMR)  " " leftTag  "`r`n"
          . "Right : " FormatNameMR(rightName, rightMR) " " rightTag "`r`n"

    if IncludeRawInText
        text .= "Raw  : " RegExReplace(raw, "\R", " ") "`r`n"

    text .= "`r`n"
    FileAppend(text, MatchTextPath, "UTF-8-RAW")
    Log("OCR: wrote formatted text -> " MatchTextPath)
}

;-- 関数: OCR_DebugNamePick(t, label := "")
;   目的: OCRに関する処理を行う。
;   引数/返り値: 定義参照
OCR_DebugNamePick(t, label := "") {
    jpr := "一-龯ぁ-んァ-ヶー"
    alr := "A-Za-z"

    Log("DBG[" label "]: raw=" RegExReplace(t, "\R", " "))

    ; 軽いクリーニング（OCR_ExtractNameSmart と同じ流れ）
    cleaned := ""
    for line in StrSplit(t, ["`r","`n"], true) {
        line := Trim(line)
        if (line = "")
            continue
        if RegExMatch(line, "i)MR|LP|WIN|LOSE|RANKED|CASUAL|REPLAY|リプレイID")
            continue
        if RegExMatch(line, "^[0-9/:.\- ,]+$")
            continue
        cleaned .= line "`n"
    }
    if (cleaned = "")
        cleaned := t
    cleaned := RegExReplace(cleaned, "(?<=[" jpr alr "])\s+(?=[" jpr alr "])", "")

    noSp := RegExReplace(cleaned, "\s+", "")
    Log("DBG[" label "]: cleaned=" RegExReplace(cleaned, "\R", " "))
    Log("DBG[" label "]: noSp=" noSp)

    cand := []
    if RegExMatch(noSp, "([ァ-ヶー]{2,})", &m1) {
        Log("DBG[" label "]: cand[kata]=" m1[1])
        cand.Push(m1[1])
    }
    if RegExMatch(noSp, "([A-Za-z" jpr "]{2,})", &m2) {
        Log("DBG[" label "]: cand[mixed]=" m2[1])
        cand.Push(m2[1])
    }
    if RegExMatch(noSp, "([A-Za-z]{2,})", &m3) {
        Log("DBG[" label "]: cand[alpha]=" m3[1])
        cand.Push(m3[1])
    }

    longest := ""
    for v in cand
        if StrLen(v) > StrLen(longest)
            longest := v
    longest := RegExReplace(longest, "(?:\d{2,}|MR|LP)+$", "")

    Log("DBG[" label "]: pick=" Trim(longest))
    return Trim(longest)
}
;-- 関数: OCR_DebugNames(winSel)
;   目的: OCRに関する処理を行う。
;   引数/返り値: 定義参照
OCR_DebugNames(winSel) {
    global OCRLang, OCRScale, OCRGray
    global NameTopFrac, NameHFrac, LeftNameX1Frac, LeftNameX2Frac, RightNameX1Frac, RightNameX2Frac

    WinGetClientPos(&cx, &cy, &cw, &ch, winSel)
    x := cx + Round(cw * 0.16),  y := cy + Round(ch * 0.08)
    w := Round(cw * 0.30),       h := Round(ch * 0.78)

    ; ヘッダー（参考表示）
    hdrH   := Round(h * 0.14)
    resHdr := OCR.FromRect(x, y, w, hdrH, {lang:OCRLang, scale:OCRScale, grayscale:OCRGray})
    txtHdr := resHdr.Text

    ; 左右情報帯（参考表示）
    sideTop := y + Round(h * 0.16)
    sideH   := Round(h * 0.24)
    leftX   := x + Round(w * 0.06), leftW  := Round(w * 0.37)
    rightX  := x + Round(w * 0.57), rightW := Round(w * 0.37)

    resLeft  := OCR.FromRect(leftX,  sideTop, leftW,  sideH, {lang:OCRLang, scale:OCRScale, grayscale:OCRGray})
    resRight := OCR.FromRect(rightX, sideTop, rightW, sideH, {lang:OCRLang, scale:OCRScale, grayscale:OCRGray})

    ; ★ 名前帯（係数で決定）
    nameTop := y + Round(h * NameTopFrac)
    nameH   := Round(h * NameHFrac)

    lnX1 := x + Round(w * LeftNameX1Frac)
    lnX2 := x + Round(w * LeftNameX2Frac)
    rnX1 := x + Round(w * RightNameX1Frac)
    rnX2 := x + Round(w * RightNameX2Frac)

    resNameLeft  := OCR.FromRect(lnX1, nameTop, lnX2-lnX1, nameH, {lang:OCRLang, scale:OCRScale, grayscale:OCRGray})
    resNameRight := OCR.FromRect(rnX1, nameTop, rnX2-rnX1, nameH, {lang:OCRLang, scale:OCRScale, grayscale:OCRGray})

    ; ログ
    Log(Format("DBG: name-ROI L=({1},{2},{3},{4}) R=({5},{6},{7},{8})"
        , lnX1, nameTop, (lnX2-lnX1), nameH
        , rnX1, nameTop, (rnX2-rnX1), nameH))

    Log("DBG: HDR=" RegExReplace(txtHdr, "\R", " "))
    Log("DBG: L-info=" RegExReplace(resLeft.Text,  "\R", " "))
    Log("DBG: R-info=" RegExReplace(resRight.Text, "\R", " "))

    resHdr.Highlight(3000, "Lime")
    resLeft.Highlight(3000, "Aqua")
    resRight.Highlight(3000, "Lime")
    ; resNameLeft.Highlight(3000, "Aqua")
    ; resNameRight.Highlight(3000, "Lime")

    ; ハイライトで矩形を可視化
    SafeHighlight(resHdr,      1000, "Yellow")
    SafeHighlight(resLeft,     1000, "Lime")
    SafeHighlight(resRight,    1000, "Aqua")
    SafeHighlight(resNameLeft, 1000, "Fuchsia")
    SafeHighlight(resNameRight,1000, "Orange")

    ; 名前の選択過程を詳細ログ
    nameLeft  := OCR_DebugNamePick(resNameLeft.Text,  "L-name")
    nameRight := OCR_DebugNamePick(resNameRight.Text, "R-name")

    Log("DBG: result L-name='" nameLeft "' / R-name='" nameRight "'")
}

;-- 関数: OCR_ExtractDateTime(t)
;   目的: OCRに関する処理を行う。
;   引数/返り値: 定義参照
OCR_ExtractDateTime(t) {
    if RegExMatch(t, "(\d{4}/\d{2}/\d{2}\s+\d{2}:\d{2})", &m)
        return m[1]
    return ""
}
;-- 関数: OCR_ExtractMode(t)
;   目的: OCRに関する処理を行う。
;   引数/返り値: 定義参照
OCR_ExtractMode(t) {
    if RegExMatch(t, "(RANKED\s+MATCH|CASUAL\s+MATCH|ランクマッチ|カジュアル)", &m)
        return m[1]
    return ""
}
;-- 関数: OCR_ExtractMR(t)
;   目的: OCRに関する処理を行う。
;   引数/返り値: 定義参照
OCR_ExtractMR(t) {
    r := OCR_ExtractRating(t)
    return r.value
}

;-- 関数: OCR_ExtractReplayID(t)
;   目的: リプレイに関する処理を行う。
;   引数/返り値: 定義参照
OCR_ExtractReplayID(t) {
    if RegExMatch(t, "リプレイID\s*([A-Z0-9]{6,12})", &m)
        return m[1]
    if RegExMatch(t, "Replay\s*ID\s*([A-Z0-9]{6,12})", &m)
        return m[1]
    return ""
}

;-- 関数: OCR_ExtractNameFromInfo(t)
;   目的: OCRに関する処理を行う。
;   引数/返り値: 定義参照
OCR_ExtractNameFromInfo(t) {
    if !t
        return ""
    norm := OCR_Normalize(t)
    ; ノイズ語を除去
    norm := RegExReplace(norm, "\b(?:WIN|LOSE|RANKED|CASUAL|REPLAY|リプレイID|PC)\b", " ")
    norm := RegExReplace(norm, "[\u3000\s]+", " ")  ; 連続空白つぶし

    ; 典型:  <name> <digits> (MR|LP)
    patNum   := "(\d{2,5}(?:,\d{3})*)"      ; 1899 / 1,899 など
    patKind  := "(M\s*R|L\s*P)"             ; MR / M R / LP / L P
    if RegExMatch(norm, "i)([A-Z0-9一-龯ぁ-んァ-ヶー・\[\]\(\)'\-_/]+?)\s*" patNum "\s*" patKind, &m) {
        name := m[1]
    } else if RegExMatch(norm, "i)" patKind "\s*" patNum "\s*([A-Z0-9一-龯ぁ-んァ-ヶー・\[\]\(\)'\-_/]+)", &m) {
        name := m[3]
    } else {
        ; 最長語フォールバック（記号類をある程度許す）
        cleaned := RegExReplace(norm, "[^A-Z0-9一-龯ぁ-んァ-ヶー・\[\]\(\)'\-_/ ]", " ")
        longest := ""
        for word in StrSplit(cleaned, " ", true)
            if StrLen(word) > StrLen(longest)
                longest := word
        name := longest
    }

    ; 末尾の余計な数字/ラベルを落とす
    name := RegExReplace(name, "\s*(\d{2,5}(?:,\d{3})*)\s*(M\s*R|L\s*P)\s*$", "")
    ; 先頭のゴミ語
    name := RegExReplace(name, "^(?:PC|ONLINE)\b\s*", "")
    return Trim(name)
}

;-- 関数: OCR_ExtractNameSmart(t)
;   目的: OCRに関する処理を行う。
;   引数/返り値: 定義参照
OCR_ExtractNameSmart(t) {
    if !t
        return ""

    ; 軽い補正（よくある置換）
    t := StrReplace(t, "」", "J")
    t := StrReplace(t, "』", "J")
    t := StrReplace(t, "′", "")
    t := StrReplace(t, "，", ",")
    ; 「・」は名前で使うことがあるので維持

    ; 行ごとにノイズ除去
    cleaned := ""
    for line in StrSplit(t, ["`r","`n"], true) {
        line := Trim(line)
        if (line = "")
            continue
        if RegExMatch(line, "i)MR|WIN|LOSE|RANKED|CASUAL|REPLAY|リプレイID")
            continue
        if RegExMatch(line, "^[0-9/:.\- ,]+$") ; 日時・数字のみ
            continue
        cleaned .= line "`n"
    }
    if (cleaned = "")
        cleaned := t

    ; 和文/英字の間に入った不要空白を連結
    jpr := "一-龯ぁ-んァ-ヶー"
    alr := "A-Za-z"
    cleaned := RegExReplace(cleaned, "(?<=[" jpr alr "])\s+(?=[" jpr alr "])", "")

    ; 候補列挙（カタカナ優先 → 和英混在 → 英字のみ）
    cand := []
    noSp := RegExReplace(cleaned, "\s+", "")
    if RegExMatch(noSp, "([ァ-ヶー]{2,})", &m1)
        cand.Push(m1[1])
    if RegExMatch(noSp, "([A-Za-z" jpr "]{2,})", &m2)
        cand.Push(m2[1])
    if RegExMatch(noSp, "([A-Za-z]{2,})", &m3)
        cand.Push(m3[1])

    ; 最長候補を採用し、末尾の MR/数字を削る
    longest := ""
    for v in cand
        if StrLen(v) > StrLen(longest)
            longest := v
    longest := RegExReplace(longest, "(?:\d{2,}|MR)+$", "")

    return Trim(longest)
}
;-- 関数: OCR_ExtractRating(t)
;   目的: OCRに関する処理を行う。
;   引数/返り値: 定義参照
OCR_ExtractRating(t) {
    norm := OCR_Normalize(t)
    ; 数字（カンマ許容）と MR/LP（空白混入許容）を前後両パターンで検出
    patNum   := "(\d{2,5}(?:,\d{3})*)"         ; 例: 1,899 / 1899 / 99999
    patLabel := "(M\s*R|L\s*P)"                ; MR / M R / LP / L P

    if RegExMatch(norm, patNum "\s*" patLabel, &m) {
        val  := StrReplace(m[1], ",")
        kind := InStr(m[2], "L") ? "LP" : "MR"
        return { value: val, kind: kind }
    }
    if RegExMatch(norm, patLabel "\s*:?\s*" patNum, &m) {
        val  := StrReplace(m[2], ",")
        kind := InStr(m[1], "L") ? "LP" : "MR"
        return { value: val, kind: kind }
    }
    ; 見つからない場合は空
    return { value: "", kind: "" }
}

;-- 関数: OCR_GetSideFlags(text)
;   目的: OCRに関する処理を行う。
;   引数/返り値: 定義参照
OCR_GetSideFlags(text) {
    t := OCR_Normalize(text)
    ; W1N / WlN / VVIN / LOSE(0誤読) / 日本語「敗北」などを吸収
    hasWin  := RegExMatch(t, "i)\bW(?:I|1|L)N\b|VVIN|WIN\b")
    hasLose := RegExMatch(t, "i)\bL(?:O|0)SE\b|LOSE\b|敗北")
    return { win: !!hasWin, lose: !!hasLose }
}


;-- 関数: OCR_Normalize(t)
;   目的: OCRに関する処理を行う。
;   引数/返り値: 定義参照
OCR_Normalize(t) {
    if !t
        return ""
    t := StrUpper(t)
    ; 全角→半角の一部
    t := StrReplace(t, "Ｗ", "W")
    t := StrReplace(t, "Ｉ", "I")
    t := StrReplace(t, "Ｌ", "L")
    t := StrReplace(t, "Ｏ", "O")
    t := StrReplace(t, "０", "0")
    t := StrReplace(t, "Ｍ", "M")
    t := StrReplace(t, "Ｒ", "R")
    t := StrReplace(t, "Ｌ", "L")
    t := StrReplace(t, "Ｐ", "P")
    t := StrReplace(t, "Ｖ", "V")
    ; 全角数字を半角に（よく出る所だけ）
    t := StrReplace(t, "０","0"), t := StrReplace(t, "１","1")
    t := StrReplace(t, "２","2"), t := StrReplace(t, "３","3")
    t := StrReplace(t, "４","4"), t := StrReplace(t, "５","5")
    t := StrReplace(t, "６","6"), t := StrReplace(t, "７","7")
    t := StrReplace(t, "８","8"), t := StrReplace(t, "９","9")
    return t
}

;=============================================================================
; [OCR] 矩形を微調整して再検出（ジッター）。最良結果のOCRオブジェクトを返す
;   rect    : {x,y,w,h} 基準矩形（ピクセル）
;   opts    : OCR.FromRect に渡すオプション
;   winSel  : 画面境界のクランプ用（省略可、未使用でも動作）
;   confGood: これ以上なら即時確定とする目安
;   stepPx  : 平行移動の基本ピクセル
;   padPx   : 矩形の拡縮量（±padを試す）
;   rounds  : ステップ半径の最大数（1で±step、2で±2*stepまで）
;-----------------------------------------------------------------------------
OCR_TryWithJitter(rect, opts, winSel := "", confGood := 0.92, stepPx := 4, padPx := 2, rounds := 2) {
    best := ""
    bestScore := -1.0

    ; 候補矩形の列挙（中心→近傍→対角の順で優先）
    candRects := []

    ; 1) 元の矩形（最初に試す）
    candRects.Push(rect)

    ; 2) パッド（拡大/縮小）だけ
    for pad in [padPx, -padPx] {
        candRects.Push(OCR_RectInflate(rect, pad, pad))
    }

    ; 3) 平行移動（ラウンドごとに±n*step）
    Loop rounds {
        r := A_Index
        s := r * stepPx
        ; 水平/垂直
        candRects.Push(OCR_RectShift(rect, +s, 0))
        candRects.Push(OCR_RectShift(rect, -s, 0))
        candRects.Push(OCR_RectShift(rect, 0, +s))
        candRects.Push(OCR_RectShift(rect, 0, -s))
        ; 対角
        candRects.Push(OCR_RectShift(rect, +s, +s))
        candRects.Push(OCR_RectShift(rect, +s, -s))
        candRects.Push(OCR_RectShift(rect, -s, +s))
        candRects.Push(OCR_RectShift(rect, -s, -s))
        ; 移動+拡大
        candRects.Push(OCR_RectInflate(OCR_RectShift(rect, +s, 0), +padPx, +padPx))
        candRects.Push(OCR_RectInflate(OCR_RectShift(rect, -s, 0), +padPx, +padPx))
        candRects.Push(OCR_RectInflate(OCR_RectShift(rect, 0, +s), +padPx, +padPx))
        candRects.Push(OCR_RectInflate(OCR_RectShift(rect, 0, -s), +padPx, +padPx))
    }

    ; 重複除去 & 画面内にクランプ
    seen := Map()
    readyRects := []
    for cr in candRects {
        cr := OCR_RectClamp(cr, winSel)
        if (cr.w <= 2 || cr.h <= 2)  ; 小さすぎる候補は除外
            continue
        key := cr.x "|" cr.y "|" cr.w "|" cr.h
        if !seen.Has(key) {
            seen[key] := true
            readyRects.Push(cr)
        }
    }

    ; OCR実施（早期確定あり）
    for cr in readyRects {
        res := OCR.FromRect(cr.x, cr.y, cr.w, cr.h, opts)
        t  := OCR__GetText(res)
        cf := OCR__GetConf(res)

        ; スコア（信頼度があればそれを優先、無ければ非空テキストに重み）
        score := (IsNumber(cf) ? cf : (t != "" ? 0.60 + 0.01 * Min(StrLen(t), 20) : 0.0))

        if (score > bestScore && t != "") {
            best := res
            bestScore := score
            ; 十分に良ければ即確定
            if (IsNumber(cf) && cf >= confGood)
                break
        }
    }

    ; 何も取れなかった場合は元の結果を返すため、最後に最低限再試行
    if (best = "") {
        try best := OCR.FromRect(rect.x, rect.y, rect.w, rect.h, opts)
    }

    ; デバッグ表示（必要ならコメントアウト）
    ; try OCR_DebugTray(best, "OCR: jitter best")

    return best
}

;--- 以下は矩形・評価ヘルパ -----------------------------------------------

OCR_RectShift(r, dx, dy) {
    return { x: r.x + dx, y: r.y + dy, w: r.w, h: r.h }
}
OCR_RectInflate(r, padX, padY) {
    return { x: r.x - padX, y: r.y - padY, w: r.w + 2*padX, h: r.h + 2*padY }
}
OCR_RectClamp(r, winSel := "") {
    ; winSelを持っていれば画面境界を取得（持っていなければ0..∞で緩くクランプ）
    left := 0, top := 0, right := 2147483647, bottom := 2147483647
    try {
        if (IsObject(winSel) && ObjHasOwnProp(winSel, "hwnd")) {
            ; クライアント領域に合わせたい場合は DllCall 等で調整してください
            WinGetPos(&wx, &wy, &ww, &wh, winSel.hwnd)
            left := wx, top := wy, right := wx + ww, bottom := wy + wh
        }
    }
    ; clamp
    x := Max(left, Min(r.x, right - 1))
    y := Max(top,  Min(r.y, bottom - 1))
    w := Max(1, Min(r.w, right  - x))
    h := Max(1, Min(r.h, bottom - y))
    return { x:x, y:y, w:w, h:h }
}

OCR__GetText(res) {
    try {
        if ObjHasOwnProp(res, "text")
            return Trim(StrReplace(res.text, "`r`n", " "))
        if ObjHasOwnProp(res, "Text")
            return Trim(StrReplace(res.Text, "`r`n", " "))
    }
    return ""
}
OCR__GetConf(res) {
    try {
        if ObjHasOwnProp(res, "confidence")
            return res.confidence
        if ObjHasOwnProp(res, "Confidence")
            return res.Confidence
    }
    return ""
}


;=============================================================================
; [ブロック] 画面キャプチャ/画像検索
; 説明: スクリーンショット、画像検索、ピクセル検査。
;=============================================================================

;-- 関数: FindAnyImage(imgList, roi, tol, &outX?, &outY?)
;   目的: 画像を検出する。
;   引数/返り値: 定義参照
FindAnyImage(imgList, roi, tol, &outX?, &outY?) {
    for img in imgList {
        if ImageSearch(&x, &y, roi.x1, roi.y1, roi.x2, roi.y2, "*n" tol " " img) {
            outX := x, outY := y
            return true
        }
    }
    return false
}
;-- 関数: IsRegionMostlyBlack(roi, darkness := 32, grid := 8, brightAllowance := 5)
;   目的: 関連処理を行う。
;   引数/返り値: 定義参照
IsRegionMostlyBlack(roi, darkness := 32, grid := 8, brightAllowance := 5) {
    ; grid x grid の等間隔サンプルを取り、RGB 全てが threshold 以下なら“黒”
    bright := 0
    dx := Max(1, Floor((roi.x2 - roi.x1) / Max(1, grid - 1)))
    dy := Max(1, Floor((roi.y2 - roi.y1) / Max(1, grid - 1)))

    Loop grid {
        i := A_Index - 1
        x := roi.x1 + i*dx
        Loop grid {
            j := A_Index - 1
            y := roi.y1 + j*dy
            CoordMode("Pixel", "Client")  ; クライアント領域基準にする（推奨）
            col := PixelGetColor(x, y, "RGB")
            Log("Color at " x "," y " = " col)
            ; col は 0xRRGGBB
            r := (col >> 16) & 0xFF
            g := (col >> 8)  & 0xFF
            b :=  col        & 0xFF
            if (r > darkness || g > darkness || b > darkness) {
                bright += 1
                if (bright > brightAllowance)
                    return false
            }
        }
    }
    return true
}

;-- 関数: WaitImageDisappear(imgList, roi, tol, timeoutMs)
;   目的: 画像に関する処理を行う。
;   引数/返り値: 定義参照
WaitImageDisappear(imgList, roi, tol, timeoutMs) {
    endTick := A_TickCount + timeoutMs
    while A_TickCount < endTick {
        if !FindAnyImage(imgList, roi, tol, &x, &y)
            return true
        Sleep 80
    }
    return false
}


;=============================================================================
; [ブロック] 入力/オートメーション
; 説明: 入力送信やホットキーなど自動操作関連。
;=============================================================================

;-- 関数: Press(key, holdMs := 60)
;   目的: 関連処理を行う。
;   引数/返り値: 定義参照
Press(key, holdMs := 60) {
    Send "{" key " down}"
    Sleep holdMs
    Send "{" key " up}"
}
;-- 関数: SendEndNavigateTest()
;   目的: 終了の処理を終了する。
;   引数/返り値: 定義参照
SendEndNavigateTest() {
    global EndDownCount, EndConfirmCount, Key_Down, Key_Confirm
    EnsureFocusGame()
    Loop EndDownCount {
        Press(Key_Down, 60)
        Sleep 120
    }
    EnsureFocusGame()
    Loop EndConfirmCount {
        Press(Key_Confirm, 70)
        Sleep 180
    }
    TrayTip "送信", "S→F をテスト送信", 800
    Log("TEST: sent S→F")
}
;-- 関数: SendNextSelection()
;   目的: 関連処理を行う。
;   引数/返り値: 定義参照
SendNextSelection() {
    global NextDirection, Key_Down, Key_Up, NextRepeats, NextIntervalMs
    dir := StrLower(NextDirection)
    key := (dir = "up") ? Key_Up : Key_Down
    Loop Max(1, NextRepeats) {
        Press(key, 60)
        Sleep NextIntervalMs
    }
}


;=============================================================================
; [ブロック] 設定/保存・読込
; 説明: 設定ファイルの読込・保存、プロファイル管理。
;=============================================================================

;-- 関数: GetROI_Load_Default()
;   目的: 読みの処理を読み込む。
;   引数/返り値: 定義参照
GetROI_Load_Default() {
    ; 画面中央寄り 60% 領域（上下左右 20% を除外）
    sw := A_ScreenWidth, sh := A_ScreenHeight
    return {x1: Round(sw*0.20), y1: Round(sh*0.20), x2: Round(sw*0.80), y2: Round(sh*0.80)}
}

;-- 関数: LoadConfig(path)
;   目的: 設定を読み込む。
;   引数/返り値: 定義参照
LoadConfig(path) {
    global NextDirection, TotalMatches, MaxRunMinutes, RolloverMinutes, RolloverMode
    global ToleranceEnd, gUseFullROI, NextRepeats, NextIntervalMs
    global Key_StartRec, Key_StopRec, Key_ToggleRec, OBSWinSelector, GameWinSelector, AutoRefocusGame
    global Img_Ends, UseOBSRecording, UseOBSToggleForRollover, CheckOnStart_Game, CheckOnStart_OBS
    global CloseGameOnStop, GameExitTimeoutMs
    global LogEnabled, LogDir, AutoScrollLog
    global ResultSnapEnabled, ResultSnapDir
    global SaveOCREnabled, SaveOCRDir
    global SlackEnabled, SlackRouterUrl, SlackTimeoutMs

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
    CloseGameOnStop := (Integer(IniRead(path, "game", "CloseOnStop", CloseGameOnStop?1:0))=1)
    GameExitTimeoutMs := Integer(IniRead(path, "game", "ExitTimeoutMs", GameExitTimeoutMs))
    imgs := IniRead(path, "images", "EndImages", "")
    if (imgs != "") {
        Img_Ends := SplitList(imgs, ";")
    }
    LogEnabled := (Integer(IniRead(path, "log", "Enabled", LogEnabled?1:0))=1)
    LogDir     := IniRead(path, "log", "Dir", LogDir)
    ResultSnapEnabled := (Integer(IniRead(path, "log", "ResultSnapEnabled", ResultSnapEnabled?1:0))=1)
    ResultSnapDir := IniRead(path, "log", "ResultSnapDir", ResultSnapDir)
    SaveOCREnabled := (Integer(IniRead(path, "ocr", "SaveOCREnabled", SaveOCREnabled?1:0))=1)
    SaveOCRDir := IniRead(path, "ocr", "SaveOCRDir", SaveOCRDir)
    AutoScrollLog := (Integer(IniRead(path, "log", "AutoScroll", AutoScrollLog?1:0))=1)
    SlackEnabled   := (Integer(IniRead(path, "slack", "Enabled", SlackEnabled?1:0))=1)
    SlackRouterUrl := IniRead(path, "slack", "RouterUrl", SlackRouterUrl)
    SlackTimeoutMs := Integer(IniRead(path, "slack", "TimeoutMs", SlackTimeoutMs))
    chkSlackEnabled.Value := SlackEnabled
    edtSlackRouter.Value := SlackRouterUrl
    edtSlackTimeout.Value := SlackTimeoutMs
    UpdateSlackUIState()
}

;-- 関数: SaveConfig(path)
;   目的: 設定を保存する。
;   引数/返り値: 定義参照
SaveConfig(path) {
    global NextDirection, TotalMatches, MaxRunMinutes, RolloverMinutes, RolloverMode
    global ToleranceEnd, gUseFullROI, NextRepeats, NextIntervalMs
    global Key_StartRec, Key_StopRec, Key_ToggleRec, OBSWinSelector, GameWinSelector, AutoRefocusGame
    global Img_Ends, UseOBSRecording, UseOBSToggleForRollover, CheckOnStart_Game, CheckOnStart_OBS
    global CloseGameOnStop, GameExitTimeoutMs
    global LogEnabled, LogDir, AutoScrollLog
    global ResultSnapEnabled, ResultSnapDir
    global SaveOCREnabled, SaveOCRDir
    global SlackEnabled, SlackRouterUrl, SlackTimeoutMs

    SlackEnabled   := chkSlackEnabled.Value
    SlackRouterUrl := Trim(edtSlackRouter.Value)
    SlackTimeoutMs := Integer(edtSlackTimeout.Value)

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
    IniWrite(CloseGameOnStop?1:0, path, "game", "CloseOnStop")
    IniWrite(GameExitTimeoutMs,   path, "game", "ExitTimeoutMs")
    IniWrite(JoinList(Img_Ends, ";"),   path, "images", "EndImages")
    IniWrite(LogEnabled?1:0, path, "log", "Enabled")
    IniWrite(LogDir,         path, "log", "Dir")
    IniWrite(ResultSnapEnabled?1:0, path, "log", "ResultSnapEnabled")
    IniWrite(ResultSnapDir, path, "log", "ResultSnapDir")
    IniWrite(SaveOCREnabled?1:0, path, "ocr", "SaveOCREnabled")
    IniWrite(SaveOCRDir, path, "ocr", "SaveOCRDir")
    IniWrite(AutoScrollLog?1:0, path, "log", "AutoScroll")
    IniWrite(SlackEnabled?1:0, path, "slack", "Enabled")
    IniWrite(SlackRouterUrl,   path, "slack", "RouterUrl")
    IniWrite(SlackTimeoutMs,   path, "slack", "TimeoutMs")
}


;=============================================================================
; [ブロック] ログ/ファイル
; 説明: ログ出力やファイルローテーション処理。
;=============================================================================

;-- 関数: Log(msg)
;   目的: ログを記録する。
;   引数/返り値: 定義参照
Log(msg) {
    global LogEnabled, LogDir, logBox, AutoScrollLog
    global gLastLogText
    if IsSet(logBox) && logBox {
        try {
            logBox.Value .= FormatTime(A_Now, "HH:mm:ss") " - " msg "`r`n"
            static EM_SETSEL := 0x00B1, EM_SCROLLCARET := 0x00B7
            hwnd := logBox.Hwnd
            SendMessage EM_SETSEL, -1, -1, , "ahk_id " hwnd
            SendMessage EM_SCROLLCARET, 0, 0, , "ahk_id " hwnd
            if AutoScrollLog
                ScrollEditToBottom(hwnd)
        }
    }

    ; 最新ログを記録してステータス再描画
    gLastLogText := FormatTime(A_Now, "HH:mm:ss") " - " msg
    SetStatus(CurrentStatusText())

    ; （以下、ファイル出力はそのまま）
    if !LogEnabled
        return
    try {
        DirCreate(LogDir)
        stamp := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
        file := LogDir "\" FormatTime(A_Now, "yyyyMMdd") ".log"
        FileAppend(stamp " - " msg "`r`n", file, "UTF-8-RAW")
    }
}
;-- 関数: LogRect(tag, r)
;   目的: ログを記録する。
;   引数/返り値: 定義参照
LogRect(tag, r) {
    Log(Format("{1}: {2},{3},{4},{5}", tag, r.x, r.y, r.w, r.h))
}


;=============================================================================
; [ブロック] 安全停止/エラー処理
; 説明: 安全停止、エラー処理、例外時の後片付け。
;=============================================================================

;-- 関数: ExitHandler(*)
;   目的: 処理の処理を処理する。
;   引数/返り値: 定義参照
ExitHandler(*) {
    global AutoSaveOnExit, ConfigPath
    if AutoSaveOnExit {
        ApplyGuiToVars()
        SaveConfig(ConfigPath)
    }
}

;-- 関数: ForceStopAutomation()
;   目的: 停止の処理を停止する。
;   引数/返り値: 定義参照
ForceStopAutomation() {
    global gRunning, gPaused, gRecording, Key_StopRec, UseOBSRecording
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

;-- 関数: RequestSafeStop()
;   目的: 停止の処理を停止する。
;   引数/返り値: 定義参照
RequestSafeStop() {
    global gRunning, gRecording, gSafeStopRequested, UseOBSRecording, Key_StopRec
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

;-- 関数: SafeHighlight(res, ms := 800, color := "Yellow")
;   目的: 関連処理を行う。
;   引数/返り値: 定義参照
SafeHighlight(res, ms := 800, color := "Yellow") {
    try {
        if !(IsObject(res))
            return
        words := []
        ; Words プロパティの取得に失敗したらスキップ
        try words := res.Words
        catch {
            return
        }
        ; 0 件ならハイライトしない（ここで抜けるのが肝）
        if !(IsObject(words)) || (words.Length = 0)
            return

        ; ライブラリのシグネチャ差異に両対応
        try {
            res.Highlight(ms, color, words*)        ; ms, color, words*
        } catch {
            ; try res.Highlight(words*, ms, color)    ; words*, ms, color
            ; catch { ; どちらもだめなら諦める
            ; }
        }
    }
}
;-- 関数: ToIntSafe(txt, def)
;   目的: 関連処理を行う。
;   引数/返り値: 定義参照
ToIntSafe(txt, def) {
    if (txt = "")
        return def
    try {
        return Integer(txt)
    } catch as e {
        return def
    }
}


;=============================================================================
; [ブロック] 診断/デバッグ
; 説明: デバッグ出力や自己診断の補助。
;=============================================================================

;-- 関数: TestBlackWait(*)
;   目的: 関連処理を行う。
;   引数/返り値: 定義参照
TestBlackWait(*) {
    Log("黒画面待機テスト開始...")
    roiLoad := GetROI_Load_Default()
    try {
        WaitWhileBlackByRatio_Window(GameWinSelector
            , BlackDarknessThreshold, BlackMinBlackRatio, BlackGridX, BlackGridY
            , BlackCheckInterval, BlackMinWait, BlackStableAfter, BlackMaxWait
            , 0.30, 0.30)  ; 中央30%×30%をサンプリング
        Log("黒画面待機テスト完了")
    } catch Error as e {
        Log("黒画面待機テスト失敗: " e.Message)
    }
}

;-- 関数: OCR_TestButton()
;   目的: OCRに関する処理を行う。
;   引数/返り値: 定義参照
OCR_TestButton() {
    global SaveOCREnabled, GameWinSelector
    if !SaveOCREnabled {
        TrayTip "OCR", "SaveOCREnabled=false（基本設定でONにしてください）", 1500
        Log("OCR TEST: skipped (SaveOCREnabled=false)")
        return
    }
    ok := OCR_RecordCurrentMatch(GameWinSelector, showHighlight := false, test := false)
    if ok {
        TrayTip "OCR", "記録成功", 1200
        Log("OCR TEST: recorded to text")
    } else {
        TrayTip "OCR", "記録失敗（ウィンドウ/結果UI未検出）", 1500
        Log("OCR TEST: failed")
    }
}
;-- 関数: OCR_TestResultButton(winSel, showHighlight := false)
;   目的: OCRを読み取る。
;   引数/返り値: 定義参照
OCR_TestResultButton(winSel, showHighlight := false) {
    ocr := OCR_RecordCurrentMatch(winSel, showHighlight, test := true)
    title := "OCR: result"
    if (ocr) {
        Log("OCR: read result fields")
        summary := ""
        try {
            if (Type(ocr) = "Object" || Type(ocr) = "Map" || Type(ocr) = "Array") {
                txt := ""
                if ObjHasOwnProp(ocr, "text")
                    txt := ocr.text
                else if ObjHasOwnProp(ocr, "Text")
                    txt := ocr.Text
                if (IsSet(txt) && txt != "")
                    summary .= "Text: " SubStr(txt, 1, 140)

                if ObjHasOwnProp(ocr, "confidence")
                    conf := ocr.confidence
                else if ObjHasOwnProp(ocr, "Confidence")
                    conf := ocr.Confidence
                if (IsSet(conf))
                    summary .= (summary ? "`n" : "") "Conf: " Round(conf, 2)

                ; 座標/サイズ（あれば）
                bbox := ""
                try {
                    if ObjHasOwnProp(ocr, "x") && ObjHasOwnProp(ocr, "y")
                        bbox := ocr.x "," ocr.y
                    if ObjHasOwnProp(ocr, "w") && ObjHasOwnProp(ocr, "h")
                        bbox := (bbox ? bbox " " : "") ocr.w "x" ocr.h
                }
                if (bbox != "")
                    summary .= (summary ? "`n" : "") "Box: " bbox

                ; 何も拾えない場合はプロパティ名一覧
                if (summary = "") {
                    props := ""
                    try {
                        arr := ObjOwnProps(ocr)
                        for _, name in arr
                            props .= (props ? ", " : "") name
                    }
                    if (props != "")
                        summary := "Props: " props
                }
            } else {
                summary := String(ocr)
            }
        } catch Error as e {
            summary := "debug error: " e.Message
        }
        if (summary = "")
            summary := "(empty)"
        TrayTip(title, summary)

    } else {
        Log("OCR: no result found")
    }
}

;-- 関数: QuickDetectTest()
;   目的: UIを検出する。
;   引数/返り値: 定義参照
QuickDetectTest() {
    roi := (gUseFullROI ? GetROI_Full() : GetROI_End_Default(GameWinSelector))
    if FindAnyImage(Img_Ends, roi, ToleranceEnd, &x, &y) {
        TrayTip "検出OK", "座標: " x "," y "  ROI=" (gUseFullROI?"Full":"Default"), 1200
        Log("TEST: detect OK at " x "," y)
    } else {
        TrayTip "未検出", "ROI=" (gUseFullROI?"Full":"Default") " / tol=" ToleranceEnd, 1200
        Log("TEST: detect NG")
    }
}

;=============================================================================
; [ブロック] ユーティリティ
; 説明: 汎用的な共通ユーティリティ群。
;=============================================================================

;-- 関数: BlackRatio(roi, darkness := 32, gx := 8, gy := 6)
;   目的: 関連処理を行う。
;   引数/返り値: 定義参照
BlackRatio(roi, darkness := 32, gx := 8, gy := 6) {
    CoordMode("Pixel", "Screen")  ; 画面座標でサンプリング
    black := 0, total := 0
    dx := Max(1, Floor((roi.x2 - roi.x1) / Max(1, gx - 1)))
    dy := Max(1, Floor((roi.y2 - roi.y1) / Max(1, gy - 1)))
    Loop gx {
        i := A_Index - 1, x := roi.x1 + i*dx
        Loop gy {
            j := A_Index - 1, y := roi.y1 + j*dy
            col := PixelGetColor(x, y, "RGB")  ; v2: 戻り値
            r := (col >> 16) & 0xFF, g := (col >> 8) & 0xFF, b := col & 0xFF
            total += 1
            if (r <= darkness && g <= darkness && b <= darkness)
                black += 1
        }
    }
    return (total>0) ? (black/total) : 0.0
}

;-- 関数: CloseGameApp(graceFirst := true)
;   目的: 閉の処理を閉じる。
;   引数/返り値: 定義参照
CloseGameApp(graceFirst := true) {
    global GameWinSelector, GameExitTimeoutMs
    if !WinExist(GameWinSelector) {
        Log("CLOSE: GAME window not found (skip)")
        return
    }

    ; 1) まず WinClose（正常終了シグナル）
    try {
        WinClose GameWinSelector
        if WinWaitClose(GameWinSelector, , GameExitTimeoutMs) {
            TrayTip "GAME終了", "WinClose で終了しました", 1200
            Log("CLOSE: WinClose success")
            return
        }
    } catch as e {
        Log("CLOSE: WinClose error - " e.Message)
    }

    ; 2) Alt+F4（UI経由の終了）
    try {
        WinActivate GameWinSelector
        if WinWaitActive(GameWinSelector, , 800) {
            Send "!{F4}"
            if WinWaitClose(GameWinSelector, , GameExitTimeoutMs) {
                TrayTip "GAME終了", "Alt+F4 で終了しました", 1200
                Log("CLOSE: Alt+F4 success")
                return
            }
        }
    } catch as e {
        Log("CLOSE: Alt+F4 error - " e.Message)
    }

    ; 3) 最終手段：プロセス終了（強制）
    try {
        pid := WinGetPID(GameWinSelector)
        if pid {
            ProcessClose pid
            TrayTip "GAME終了(強制)", "ProcessClose で終了しました", 1500
            Log("CLOSE: ProcessClose success (pid=" pid ")")
        } else {
            Log("CLOSE: cannot get PID for ProcessClose")
        }
    } catch as e {
        TrayTip "GAME終了失敗", "手動で終了してください", 2000
        Log("CLOSE: ProcessClose failed - " e.Message)
    }
}

;-- 関数: EnsureFocusGame()
;   目的: 保証の処理を保証する。
;   引数/返り値: 定義参照
EnsureFocusGame() {
    global GameWinSelector
    if WinExist(GameWinSelector) {
        try WinRestore GameWinSelector
        WinActivate GameWinSelector
        if WinWaitActive(GameWinSelector, , 500) {
            Sleep 60
            return true
        }
    }
    TrayTip "警告", "GAMEのウィンドウが見つかりません / アクティブ化できません", 1200
    Log("FOCUS: GAME window not found or activate failed: " GameWinSelector)
    return false
}

;-- 関数: FixFrac(fr)
;   目的: 関連処理を行う。
;   引数/返り値: 定義参照
FixFrac(fr) {   
    fr.x1 := Clamp01(fr.x1), fr.y1 := Clamp01(fr.y1)
    fr.x2 := Clamp01(fr.x2), fr.y2 := Clamp01(fr.y2)
    if (fr.x2 <= fr.x1) fr.x2 := Min(1, fr.x1 + 0.02)  ; 最低幅2%
    if (fr.y2 <= fr.y1) fr.y2 := Min(1, fr.y1 + 0.02)  ; 最低高2%
    return fr
}

;-- 関数: FormatNameMR(name, mr)
;   目的: 整形の処理を整形する。
;   引数/返り値: 定義参照
FormatNameMR(name, mr) {
    name := Trim(name)
    return (mr != "" ? Format("{1} ({2} MR)", name, mr) : name)
}
;-- 関数: FormatNameRating(name, val, kind := "MR")
;   目的: 整形の処理を整形する。
;   引数/返り値: 定義参照
FormatNameRating(name, val, kind := "MR") {
    name := Trim(name), kind := (kind!="" ? kind : "MR")
    return (val != "" ? Format("{1} ({2} {3})", name, val, kind) : name)
}


;-- 関数: FracName(fr)
;   目的: 関連処理を行う。
;   引数/返り値: 定義参照
FracName(fr) {
    for k, v in [ROI_ReplayID, ROI_Time, ROI_L_Name, ROI_R_Name
                , ROI_L_Rating, ROI_R_Rating, ROI_L_Result, ROI_R_Result] {
        if (v is Map && v.x1=fr.x1 && v.y1=fr.y1 && v.x2=fr.x2 && v.y2=fr.y2)
            return ["ReplayID","Time","L_Name","R_Name","L_Rating","R_Rating","L_Res","R_Res"][A_Index]
    }
    return "?"
}

;-- 関数: GetClientRectScreen(winSel)
;   目的: 関連処理を行う。
;   引数/返り値: 定義参照
GetClientRectScreen(winSel) {
    hwnd := WinExist(winSel)
    if !hwnd
        throw Error("window not found: " winSel)
    rect := Buffer(16, 0)
    DllCall("User32.dll\GetClientRect", "ptr", hwnd, "ptr", rect)
    pt := Buffer(8, 0)
    DllCall("User32.dll\ClientToScreen", "ptr", hwnd, "ptr", pt)
    x := NumGet(pt, 0, "int"), y := NumGet(pt, 4, "int")
    w := NumGet(rect, 8, "int"), h := NumGet(rect,12, "int")
    return {x:x, y:y, w:w, h:h}
}

;-- 関数: GetROI_End_Default(winSel)
;   目的: 終了の処理を終了する。
;   引数/返り値: 定義参照
GetROI_End_Default(winSel) {
    if !WinExist(winSel)
        throw Error("Window not found: " winSel)
    ; クライアント領域の画面座標を取得
    WinGetClientPos(&cx, &cy, &cw, &ch, winSel)
    if (cw <= 0 || ch <= 0)
        throw Error("Invalid client size.")

    x1 := cx + Round(cw * 0.10)
    y1 := cy + Round(ch * 0.45)
    x2 := cx + Round(cw * 0.90)
    y2 := cy + Round(ch * 0.98)
    return {x1:x1, y1:y1, x2:x2, y2:y2}
}

;-- 関数: GetROI_Full()
;   目的: 関連処理を行う。
;   引数/返り値: 定義参照
GetROI_Full() {
    sw := A_ScreenWidth, sh := A_ScreenHeight
    return {x1: 0, y1: 0, x2: sw-1, y2: sh-1}
}

;-- 関数: IsBadName(n)
;   目的: 関連処理を行う。
;   引数/返り値: 定義参照
IsBadName(n) {
    n := Trim(n)
    if (n = "")
        return true
    if (StrLen(n) < 2)
        return true
    if RegExMatch(n, "i)\b(CLUB|TEAM|CREW|CLAN|GUILD)\b")
        return true
    if RegExMatch(n, "^[0-9]+$")
        return true
    return false
}

;-- 関数: JoinList(arr, sep:=";")
;   目的: 関連処理を行う。
;   引数/返り値: 定義参照
JoinList(arr, sep:=";") {
    out := ""
    for i, v in arr {
        out .= (i>1 ? sep : "") . v
    }
    return out
}
;-- 関数: MsToHMS(ms)
;   目的: 関連処理を行う。
;   引数/返り値: 定義参照
MsToHMS(ms) {
    if (ms < 0)  ; 念のため
        ms := 0
    total := Floor(ms / 1000)
    h := Floor(total / 3600)
    m := Floor(Mod(total, 3600) / 60)
    s := Mod(total, 60)
    return Format("{:02}:{:02}:{:02}", h, m, s)
}

;-- 関数: RectFromFrac(winSel, fr)
;   目的: 関連処理を行う。
;   引数/返り値: 定義参照
RectFromFrac(winSel, fr) {
    global ROI_PANEL
    base := GetClientRectScreen(winSel)           ; 画面座標のクライアント
    ; PANEL の実ピクセル
    px := base.x + Round(base.w * ROI_PANEL.x)
    py := base.y + Round(base.h * ROI_PANEL.y)
    pw := Round(base.w * ROI_PANEL.w)
    ph := Round(base.h * ROI_PANEL.h)

    ; 入力の安全化（文字列→数値、0～1にクランプ）
    _c := v => Max(0.0, Min(1.0, v+0.0))
    x1 := _c(fr.x1), y1 := _c(fr.y1), x2 := _c(fr.x2), y2 := _c(fr.y2)
    if (x2 <= x1)  ; 幅ゼロ防止
        x2 := Min(1.0, x1 + 0.02)
    if (y2 <= y1)  ; 高さゼロ防止
        y2 := Min(1.0, y1 + 0.02)

    ; ←ここがポイント：幅/高さは “差分”
    x := px + Round(pw * x1)
    y := py + Round(ph * y1)
    w := Max(1, Round(pw * (x2 - x1)))
    h := Max(1, Round(ph * (y2 - y1)))

    ; PANEL内に収める（はみ出し対策）
    w := Min(w, (px+pw) - x)
    h := Min(h, (py+ph) - y)
    return {x:x, y:y, w:w, h:h}
}

;-- 関数: RefocusGame(force := false)
;   目的: 関連処理を行う。
;   引数/返り値: 定義参照
RefocusGame(force := false) {
    global AutoRefocusGame
    if (!force && !AutoRefocusGame)
        return
    EnsureFocusGame()
}

;-- 関数: RestoreFocus(prevHwnd)
;   目的: 関連処理を行う。
;   引数/返り値: 定義参照
RestoreFocus(prevHwnd) {
    ; 旧互換：使わないが残しておく
    EnsureFocusGame()
}

;-- 関数: ScrollEditToBottom(hwnd)
;   目的: 関連処理を行う。
;   引数/返り値: 定義参照
ScrollEditToBottom(hwnd) {
    static WM_VSCROLL := 0x0115, SB_BOTTOM := 7
    static EM_SETSEL := 0x00B1, EM_SCROLLCARET := 0x00B7
    ; 1) キャレット移動 + CARETスクロール（効かない環境もあるため保険）
    SendMessage EM_SETSEL, -1, -1, , "ahk_id " hwnd
    SendMessage EM_SCROLLCARET, 0, 0, , "ahk_id " hwnd
    ; 2) さらにVスクロール命令で最下段へ（ReadOnly/非フォーカスでも効く）
    PostMessage WM_VSCROLL, SB_BOTTOM, 0, , "ahk_id " hwnd
}

;-- 関数: SplitList(txt, sep:=";")
;   目的: 関連処理を行う。
;   引数/返り値: 定義参照
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

;-- 関数: StrJoin(arr, sep:="`n")
;   目的: 関連処理を行う。
;   引数/返り値: 定義参照
StrJoin(arr, sep:="`n") {
    s := ""
    for i,v in arr
        s .= (i>1?sep:"") v
    return s
}

;-- 関数: ToggleFullROI()
;   目的: 切り替の処理を切り替える。
;   引数/返り値: 定義参照
ToggleFullROI() {
    global gUseFullROI
    gUseFullROI := !gUseFullROI
    chkROI.Value := gUseFullROI ? 1 : 0
    TrayTip "ROI切替", (gUseFullROI ? "全画面" : "中央～下"), 1000
    Log("ROI: " (gUseFullROI ? "Full" : "Default"))
}
;-- 関数: TogglePause()
;   目的: 切り替の処理を切り替える。
;   引数/返り値: 定義参照
TogglePause() {
    global gPaused
    gPaused := !gPaused
    TrayTip (gPaused ? "一時停止" : "再開"), "", 900
    Log(gPaused ? "PAUSE" : "RESUME")
}

;-- 関数: TryShiftNameUp(winSel, side, x1, x2, baseTop, bandH, h, opts, maxSteps := 4, stepFrac := 0.02)
;   目的: 関連処理を行う。
;   引数/返り値: 定義参照
TryShiftNameUp(winSel, side, x1, x2, baseTop, bandH, h, opts, maxSteps := 4, stepFrac := 0.02) {
    bestName := ""
    bestTop  := baseTop
    stepPx   := Round(h * stepFrac)

    Loop maxSteps {
        top := baseTop - (A_Index * stepPx)
        if (top < 0)
            break
        try {
            res := OCR.FromRect(x1, top, x2-x1, bandH, opts)
            raw := res.Text
            cand := OCR_ExtractNameSmart(raw)
            if !IsBadName(cand) {
                bestName := cand, bestTop := top
                ; 可視化（デバッグ）
                SafeHighlight(res, 600, (side="left"?"Fuchsia":"Orange"))
                Log(Format("DBG: {1}-shift#{2} top={3} pick='{4}'", side, A_Index, top, cand))
                break
            } else {
                Log(Format("DBG: {1}-shift#{2} top={3} raw='{4}' -> NG", side, A_Index, top, RegExReplace(raw,"\R"," ")))
            }
        } catch as e {
            Log("DBG: shift OCR error - " e.Message)
        }
    }
    return {name:bestName, top:bestTop}
}

;-- 関数: UpdatePauseBtn()
;   目的: 更新の処理を更新する。
;   引数/返り値: 定義参照
UpdatePauseBtn() {
    global gPaused, btnPause
    static last := ""
    newLabel := (gPaused ? "再開" : "一時停止")
    if (newLabel != last) {
        btnPause.Text := newLabel
        last := newLabel
    }
}

;===========================================================
; NirCmd を使ってアクティブウィンドウをPNG保存
; ファイル名: {録画ベース}_{経過}[_{label}].png
;   - 録画ベース: gCurrentTextPath の拡張子なし
;   - 経過: MsToHMS(A_TickCount - gRecStartTick)（:→-）
; 依存: MsToHMS(), gCurrentTextPath, gRecStartTick
;===========================================================

CaptureWithNirCmd(winSel, label := "", outDir := "") {
    hwnd := OCR__ResolveHwnd(winSel)
    if !hwnd
        return false

    ; NirCmd の場所（優先: tools\ 配下 → スクリプト直下）
    exe := A_ScriptDir "\tools\nircmd.exe"
    if !FileExist(exe)
        exe := A_ScriptDir "\nircmd.exe"
    if !FileExist(exe) {
        try Log("CaptureWithNirCmd: nircmd.exe not found")
        return false
    }

    path := OCR_BuildSnapPath(label, outDir)

    ; 直前のアクティブ窓を保存しておき、撮影後に戻す
    prev := WinExist("A")

    ; NirCmdの savescreenshotwin は「アクティブウィンドウ」を撮るため、対象を一時的にアクティブ化
    if (prev != hwnd) {
        WinActivate "ahk_id " hwnd
        WinWaitActive("ahk_id " hwnd, , 1)       ; 最大1秒待ち
        Sleep 120                                 ; 再描画の猶予
    }

    cmd := Format('"{1}" savescreenshotwin "{2}"', exe, path)
    rc  := RunWait(cmd, , "Hide")

    ; フォーカス復帰
    if (prev && prev != hwnd) {
        WinActivate "ahk_id " prev
    }

    if (rc = 0 && FileExist(path)) {
        try Log("Snapshot saved (NirCmd): " path)
        return true
    } else {
        try Log("CaptureWithNirCmd failed: rc=" rc " -> " path)
        return false
    }
}

; クライアント領域だけを NirCmd の savescreenshot で保存
CaptureWithNirCmd_Client(winSel, label := "", outDir := "") {
    hwnd := OCR__ResolveHwnd(winSel)
    if !hwnd
        return false
    exe := A_ScriptDir "\tools\nircmd.exe"
    if !FileExist(exe)
        exe := A_ScriptDir "\nircmd.exe"
    if !FileExist(exe) {
        try Log("CaptureWithNirCmd_Client: nircmd.exe not found")
        return false
    }

    path := OCR_BuildSnapPath(label, outDir)

    ; --- クライアント矩形をスクリーン座標に変換
    rc := Buffer(16, 0)                              ; RECT {left,top,right,bottom}
    DllCall("GetClientRect", "ptr", hwnd, "ptr", rc)
    left   := NumGet(rc, 0,  "int")
    top    := NumGet(rc, 4,  "int")
    right  := NumGet(rc, 8,  "int")
    bottom := NumGet(rc, 12, "int")
    pt := Buffer(8, 0)                               ; POINT for top-left
    NumPut("int", left, pt, 0), NumPut("int", top, pt, 4)
    DllCall("ClientToScreen", "ptr", hwnd, "ptr", pt)
    sx := NumGet(pt, 0, "int"), sy := NumGet(pt, 4, "int")
    w := right - left, h := bottom - top
    if (w <= 0 || h <= 0) {
        try Log("CaptureWithNirCmd_Client: invalid client rect")
        return false
    }

    ; --- クライアント矩形を指定して保存
    ;     savescreenshot x y width height "filename"
    cmd := Format('"{1}" savescreenshot {2} {3} {4} {5} "{6}"', exe, sx, sy, w, h, path)
    rc2 := RunWait(cmd, , "Hide")
    ok := (rc2 = 0 && FileExist(path))
    try Log(ok ? "Snapshot saved (client): " path : "CaptureWithNirCmd_Client failed rc=" rc2)
    return ok
}

;-----------------------------------------
; 共有ヘルパ
;-----------------------------------------
OCR_BuildSnapPath(label := "", outDir := "") {
    if (outDir = "")
        outDir := A_ScriptDir "\snapshots"
    if !DirExist(outDir)
        DirCreate(outDir)

    global gCurrentTextPath, gRecStartTick
    base := "rec"
    try {
        if (IsSet(gCurrentTextPath) && gCurrentTextPath != "") {
            SplitPath(gCurrentTextPath, , , , &noext)
            if (noext != "")
                base := noext
        }
    }
    elapsed := MsToHMS(Max(0, A_TickCount - gRecStartTick))
    elapsed := StrReplace(elapsed, ":", "-")
    if (label != "")
        base := base "_" label

    fn := outDir "\" OCR__SanitizeFilename(base "_" elapsed) ".png"
    return OCR__EnsureUniquePath(fn)
}

OCR__ResolveHwnd(winSel) {
    try {
        if IsInteger(winSel)
            return winSel
        if IsObject(winSel) && ObjHasOwnProp(winSel, "hwnd")
            return winSel.hwnd
    }
    try return WinExist("A")
    return 0
}

OCR__SanitizeFilename(s) {
    return RegExReplace(s, '[<>:"/\\|?*\x00-\x1F]', "")
}

OCR__EnsureUniquePath(path) {
    if !FileExist(path)
        return path
    SplitPath(path, &fn, &dir, &ext, &nameNoExt)
    i := 1
    loop {
        alt := dir "\" nameNoExt "-" i "." ext
        if !FileExist(alt)
            return alt
        i += 1
    }
}

;===========================================================
; PSCapture でスクリーンショット保存（AHK v2）
;   - ファイル名: {録画ベース}_{経過}[_{label}].png
;   - mode="client"  : クライアント領域のみ（枠/影なし）
;     mode="frame"   : ウィンドウ全体から影を自動トリミング（-aeroborder）
; 依存: MsToHMS(), gCurrentTextPath, gRecStartTick
;===========================================================

; 使い分けの入口（基本はこれを呼ぶ）
CaptureWithPSCapture(winSel, label := "", outDir := "", mode := "client") {
    if (StrLower(mode) = "frame")
        return CaptureWithPSCapture_FrameNoShadow(winSel, label, outDir)
    return CaptureWithPSCapture_Client(winSel, label, outDir)
}
; 1) クライアント領域のみ（枠・影なし）
CaptureWithPSCapture_Client(winSel, label := "", outDir := "") {
    hwnd := OCR__ResolveHwnd(winSel)
    if !hwnd
        return false
    path := OCR_BuildSnapPath(label, outDir)
    rc := PSCap.captureWindowClient(hwnd, path)
    ok   := (rc = 0 && FileExist(path))
    try Log(ok ? "PScapture client saved: " path : "PScapture client failed rc=" rc)
    return ok
}

; 2) 全体
CaptureWithPSCapture_FrameNoShadow(winSel, label := "", outDir := "", border := 200, flatten := true, bg := 255) {
    hwnd := OCR__ResolveHwnd(winSel)
    if !hwnd
        return false
    path := OCR_BuildSnapPath(label, outDir)
    rc := PSCap.capturePrimary(hwnd)
    ok := (rc = 0 && FileExist(path))
    try Log(ok ? "PScapture frame saved: " path : "PScapture frame failed rc=" rc)
    return ok
}

;===========================================================
; MiniCap でスクリーンショット保存（AHK v2）
;   - ファイル名: {録画ベース}_{経過}[_{label}].png
;   - mode="client"  : クライアント領域のみ（枠/影なし）
;     mode="frame"   : ウィンドウ全体から影を自動トリミング（-aeroborder）
; 依存: MsToHMS(), gCurrentTextPath, gRecStartTick
;===========================================================

; 使い分けの入口（基本はこれを呼ぶ）
CaptureWithMiniCap(winSel, label := "", outDir := "", mode := "client") {
    if (StrLower(mode) = "frame")
        return CaptureWithMiniCap_FrameNoShadow(winSel, label, outDir)
    return CaptureWithMiniCap_Client(winSel, label, outDir)
}

; 1) クライアント領域のみ（枠・影なし）
CaptureWithMiniCap_Client(winSel, label := "", outDir := "") {
    hwnd := OCR__ResolveHwnd(winSel)
    if !hwnd
        return false
    exe := MiniCap_ResolvePath()
    if !exe {
        try Log("MiniCap not found")
        return false
    }
    MiniCap_PrepareWindow(hwnd)  ; 最小化対策（必要ならコメントアウト）

    path := OCR_BuildSnapPath(label, outDir)
    cmd  := Format('"{1}" -capturehwnd {2} -client -nofocus -save "{3}" -exit -stderr', exe, hwnd, path)
    rc   := RunWait(cmd, , "Hide")
    ok   := (rc = 0 && FileExist(path))
    try Log(ok ? "MiniCap client saved: " path : "MiniCap client failed rc=" rc)
    return ok
}

; 2) タイトルバーは残して“影だけ”除去（-aeroborder）
;    border は大きめに（例: 200）で影トリムを強める
CaptureWithMiniCap_FrameNoShadow(winSel, label := "", outDir := "", border := 200, flatten := true, bg := 255) {
    hwnd := OCR__ResolveHwnd(winSel)
    if !hwnd
        return false
    exe := MiniCap_ResolvePath()
    if !exe {
        try Log("MiniCap not found")
        return false
    }
    MiniCap_PrepareWindow(hwnd)

    path := OCR_BuildSnapPath(label, outDir)
    opts := Format('-aeroborder {1}', border)
    if (flatten)
        opts .= Format(' -aeroflatten -aerocolor {1}', bg) ; 0..255（白=255, 黒=0）

    cmd := Format('"{1}" -capturehwnd {2} -nofocus {3} -save "{4}" -exit -stderr', exe, hwnd, opts, path)
    rc  := RunWait(cmd, , "Hide")
    ok  := (FileExist(path))
    try Log(ok ? "MiniCap frame saved: " path : "MiniCap frame failed rc=" rc)
    return ok
}

; --- MiniCap の場所解決 ---
MiniCap_ResolvePath() {
    exe := A_ScriptDir "\tools\MiniCap\MiniCap.exe"
    if FileExist(exe)
        return exe
    exe := A_ScriptDir "\tools\MiniCap.exe"
    if FileExist(exe)
        return exe
    exe := A_ScriptDir "\MiniCap.exe"
    if FileExist(exe)
        return exe
    return ""
}

; --- 最小化時は復帰（アクティブ化は行わない） ---
MiniCap_PrepareWindow(hwnd) {
    try {
        mm := WinGetMinMax("ahk_id " hwnd) ; 1=max, -1=min, 0=normal
        if (mm = -1)
            WinRestore "ahk_id " hwnd
        ; 必要に応じて: WinShow "ahk_id " hwnd
        Sleep 80
    }
}

; =========================
; Slack Notify (via slack-message-router)
; =========================

SlackNotify(text, level := "info") {
    global SlackEnabled, SlackRouterUrl, SlackTimeoutMs

    if (!SlackEnabled)
        return
    if (SlackRouterUrl = "")
        return

    text := StrReplace(text, "\", "\\")
    text := StrReplace(text, "`"", "`"`"" )
    text := StrReplace(text, "`r", "\r")
    text := StrReplace(text, "`n", "\n")

    body := "{`"text`":`"" text "`",`"level`":`"" level "`"}"

    try {
        req := ComObject("WinHttp.WinHttpRequest.5.1")
        req.Open("POST", SlackRouterUrl, false)
        req.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
        t := SlackTimeoutMs ? SlackTimeoutMs : 2000
        req.SetTimeouts(t, t, t, t)
        req.Send(body)
    } catch {
    }
}


