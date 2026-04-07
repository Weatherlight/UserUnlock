@echo off
setlocal
title ユーザーアンロックスクリプト

:: バッチが存在するフォルダをカレントディレクトリにする
cd /d %~dp0

:: スクリプト名を設定
set SCRIPT_NAME=UserUnlock.ps1

echo ======================================================
echo   ユーザーアンロックスクリプト 起動バッチ
echo ======================================================
echo.

:: PowerShellの実行
:: %~dp0%SCRIPT_NAME% でバッチと同じ場所のスクリプトを指定
echo [情報] 処理を開始します...
:: PowerShellをSTAモードでバックグラウンド起動 (権限チェックはPS1側で行う)
start "" powershell -Sta -WindowStyle Hidden -ExecutionPolicy Bypass -NoProfile -File "%~dp0%SCRIPT_NAME%"

echo.
echo ======================================================
echo   すべての処理が終了しました。
echo ======================================================