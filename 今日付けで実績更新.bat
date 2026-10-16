@echo off
chcp 65001 > nul
cd /d "%~dp0"

set "SCRIPT=%~dp0今日付けで実績更新.py"
if not exist "%SCRIPT%" (
    echo [エラー] 今日付けで実績更新.py が見つかりません。
    echo BATファイルとPythonファイルを同じフォルダに置いてください。
    pause
    exit /b 1
)

where py > nul 2>&1
if errorlevel 1 (
    python "%SCRIPT%"
) else (
    py -3 "%SCRIPT%"
)

if errorlevel 1 (
    echo.
    echo [エラー] スクリプトの実行に失敗しました。
    pause
    exit /b 1
)
