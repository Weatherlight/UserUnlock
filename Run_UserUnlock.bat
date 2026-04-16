@echo off
setlocal
title ユーザーアンロックスクリプト

:: バッチが存在するフォルダをカレントディレクトリにする
cd /d %~dp0

:: スクリプト名を設定
set SCRIPT_NAME=UserUnlock.ps1

:: =====================================================
:: 資格情報の設定（任意）
:: 記載した場合はその値を PowerShell スクリプトへ渡す。
:: 省略（空欄）のままにした場合は実行時にダイアログで入力を促す。
:: =====================================================
set USERNAME=
set PLAINPASSWORD=
:: =====================================================

echo ======================================================
echo   ユーザーアンロックスクリプト 起動バッチ
echo ======================================================
echo.

:: PowerShellの実行
:: %~dp0%SCRIPT_NAME% でバッチと同じ場所のスクリプトを指定
echo [情報] 処理を開始します...

:: USERNAME / PLAINPASSWORD が設定されている場合のみ引数として追加する
set PS_ARGS=
if not "%USERNAME%"=="" set PS_ARGS=%PS_ARGS% -Username "%USERNAME%"
if not "%PLAINPASSWORD%"=="" set PS_ARGS=%PS_ARGS% -PlainPassword "%PLAINPASSWORD%"

:: PowerShellをSTAモードでバックグラウンド起動 (権限チェックはPS1側で行う)
start "" powershell -Sta -WindowStyle Hidden -ExecutionPolicy Bypass -NoProfile -File "%~dp0%SCRIPT_NAME%"%PS_ARGS%

echo.
echo ======================================================
echo   すべての処理が終了しました。
echo ======================================================