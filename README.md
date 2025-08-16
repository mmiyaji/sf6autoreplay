# Auto Replay

※動作保証しません　なにかあっても責任取りません　ジョークソフトです

AutoHotKeyを利用してとある格闘ゲームのリプレイを自動再生/OBS自動録画するツールです

- 終了UIを画像認識して次リプレイへ自動進行  
- 試合ごとに録画ファイルを安全に分割  
- GUI で設定・保存が可能（外部設定ファイル対応）  
- OBS の録画を使用するかどうかをオプション指定可能  
- ログタブとステータスバーで動作を確認可能  

---

## 📌 主な機能

- **リプレイ再生開始**: 「決定(Fキー)」を2回送信して自動再生  
- **終了UI検出**: 指定画像 (`assets/end_result1.png` など) を検出して次リプレイへ  
- **録画制御**:  
  - `Ctrl+F7` で録画開始  
  - `Ctrl+F8` で録画停止  
  - `Ctrl+F9` で録画ファイルを分割（再開付き）  
- **安全停止モード**:  
  次の終了UIが出るまで待ってから録画を終了  
- **ログ出力**: ファイル & GUI にリアルタイム表示  
- **GUIタブ形式**:  
  - 設定  
  - 制御  
  - ログ  

---

## 🛠️ 必要環境

- **Windows 10 / 11**
- **[AutoHotkey v2](https://www.autohotkey.com/)**
- **OBS Studio**（録画制御を使う場合）
- **とある格闘ゲーム (Steam版)**

---

## ⚙️ 設定方法

1. `assets/` フォルダに終了UI画像を配置  
   （例: `end_result1.png` / `end_result2.png` …）  
   自分の環境に合わせてキャプチャし直して画像を置き換えてください。解像度や表示言語の違いにより正しく認識しないためです。
2. スクリプトを起動すると GUI が開きます  
3. 必要に応じて設定を変更し、**保存**します  
   - 保存形式は `config.ini`  
   - 複数画像パスは **`;` 区切り**  

---

## 🎮 操作方法

| ホットキー          | 動作                                   |
|---------------------|----------------------------------------|
| `Ctrl+Alt+S`        | 自動リプレイ開始                       |
| `Ctrl+Alt+X`        | 安全停止（試合終了後に録画停止）       |
| `Ctrl+Alt+Shift+X`  | 即時停止（強制）                       |
| `Ctrl+Alt+P`        | 一時停止 / 再開                        |
| `Ctrl+Alt+W`        | 終了UI検出テスト                       |
| `Ctrl+Alt+R`        | ROI切替（全画面 or 既定領域）         |
| `Ctrl+Alt+E`        | S → F テスト送信                      |
| `Ctrl+Alt+T`        | OBSキー送信テスト                      |

---

## 📄 設定ファイル例 (`config.ini`)

```ini
[General]
TotalMatches=50
MatchHardTimeoutSec=300
Delay_BetweenItems=2000

[OBS]
EnableOBS=1
Key_StartRec=^{F7}
Key_StopRec=^{F8}
Key_SwitchRec=^{F9}

[Images]
EndUI=assets\end_result.png;assets\end_result2.png;assets\end_result3.png
ToleranceEnd=180
