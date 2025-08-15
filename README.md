# Auto Replay

※動作保証しません　なにかあっても責任取りません　ジョークソフトです

AutoHotKeyを利用してとあるゲームのリプレイを自動再生するツール

OBSの録画機能と連携し自動で次のリプレイを再生します

手動での停止かマッチ回数条件か時間条件を満たす場合に終了します

## 使い方

- Ctrl+Alt+S  開始(リプレイ一覧画面からスタート)
- Ctrl+Alt+X  安全停止（次の終了UIまで待ってから録画停止）
- Ctrl+Alt+P  一時停止/再開
- Ctrl+Alt+W  その場で検出テスト
- Ctrl+Alt+R  ROI=全画面と既定の切替
- Ctrl+Alt+E  S→F テスト送信
- Ctrl+Alt+T  OBSキー送信テスト
- Ctrl+Alt+Shift+X  即時停止（強制）

## 設定

- OBS録画開始：Ctrl+F7（OBS側のホットキー設定と合わせる）
- OBS録画終了: Ctrl+F8（OBS側のホットキー設定と合わせる）

- MaxRunMinutes := 60 ; 0 の場合は時間判定を完全にスキップし、回数条件のみを参照
- TotalMatches := 20               ; 0=無限
