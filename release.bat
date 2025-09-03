@echo off
setlocal

rem === パス推定（既定インストール）===
set "AHK2EXE=%ProgramFiles%\AutoHotkey\Compiler\Ahk2Exe.exe"
set "AHK_BASE=%ProgramFiles%\AutoHotkey\v2\AutoHotkey64.exe"

rem 見つからない場合は ProgramFiles(x86) 側もチェック
if not exist "%AHK2EXE%" (
    if exist "%ProgramFiles(x86)%\AutoHotkey\Compiler\Ahk2Exe.exe" set "AHK2EXE=%ProgramFiles(x86)%\AutoHotkey\Compiler\Ahk2Exe.exe"
)
if not exist "%AHK_BASE%" (
    if exist "%ProgramFiles(x86)%\AutoHotkey\v2\AutoHotkey64.exe" set "AHK_BASE=%ProgramFiles(x86)%\AutoHotkey\v2\AutoHotkey64.exe"
)

if not exist "%AHK2EXE%" (
    echo [ERROR] Ahk2Exe.exe が見つかりませんでした。AutoHotkey のインストールを確認してください。
    exit /b 1
)

if not exist "%AHK_BASE%" (
    echo [ERROR] AutoHotkey64.exe のベースEXEが見つかりませんでした。
    exit /b 1
)

rem === 入出力設定 ===
set "SRC=sf6autoreplay.ahk"   rem メインスクリプトに合わせて変更
set "ICON=icons\rec_icon.ico"
set "OUT=dist\sf6autoreplay.exe"

if not exist dist mkdir dist

echo Building: "%SRC%" -> "%OUT%"
"%AHK2EXE%" /in "%SRC%" /out "%OUT%" /base "%AHK_BASE%" /icon "%ICON%"
rem /mpress 1: UPX同梱の MPress 圧縮（サイズ削減）。不具合があれば外してください。

if errorlevel 1 (
    echo [FAIL] Build failed.
    exit /b 1
) else (
    echo [OK] Build success: %OUT%
)

endlocal
