@echo off

REM 管理者権限で実行されているか確認
whoami /groups | find "S-1-16-12288" > nul
if errorlevel 1 (
    echo このスクリプトは管理者権限で実行する必要があります。
    pause
    exit /b
)

REM 現在のスリープ設定をログに出力
powercfg /query > sleep_log.txt
if not exist sleep_log.txt (
    echo powercfg コマンドの実行に失敗しました。
    pause
    exit /b
)

REM ログからAC電源とDC電源のスリープ設定値を取得
for /f "tokens=5 delims=: " %%a in ('findstr /i /c:"現在の AC 電源設定のインデックス" sleep_log.txt') do set SLEEP_AC_ORIGINAL=%%a
for /f "tokens=5 delims=: " %%a in ('findstr /i /c:"現在の DC 電源設定のインデックス" sleep_log.txt') do set SLEEP_DC_ORIGINAL=%%a

REM 値が取得できたか確認
if "%SLEEP_AC_ORIGINAL%"=="" (
    echo AC電源のスリープ設定値が取得できませんでした。
    type sleep_log.txt
    pause
    exit /b
)
if "%SLEEP_DC_ORIGINAL%"=="" (
    echo DC電源のスリープ設定値が取得できませんでした。
    type sleep_log.txt
    pause
    exit /b
)

REM スリープを無効化
powercfg /change standby-timeout-ac 0
powercfg /change standby-timeout-dc 0

echo スリープを無効化しました。終了するにはCtrl+Cを押してください。

:LOOP
REM 処理を継続
ping -n 2 127.0.0.1 > nul
GOTO LOOP

:EXIT
REM 元の設定を復元
powercfg /change standby-timeout-ac %SLEEP_AC_ORIGINAL%
powercfg /change standby-timeout-dc %SLEEP_DC_ORIGINAL%

echo スリープの設定を元に戻しました。
pause
