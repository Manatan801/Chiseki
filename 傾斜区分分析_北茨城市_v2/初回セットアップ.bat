@echo off
chcp 932 >nul 2>&1
setlocal EnableDelayedExpansion
pushd "%~dp0"
echo ============================================
echo  傾斜区分分析ツール 初回セットアップ
echo ============================================
echo.

REM Python実行コマンド検出
set PY_CMD=
for %%P in (py python python3) do (
    if "!PY_CMD!"=="" (
        %%P --version >nul 2>&1
        if !errorlevel! == 0 (
            set PY_CMD=%%P
            echo Python検出: %%P
        )
    )
)

if "!PY_CMD!"=="" (
    echo [エラー] Pythonが見つかりません。
    echo Pythonをインストールしてから再実行してください。
    echo https://www.python.org/downloads/
    pause
    popd
    exit /b 1
)

REM Pythonバージョン表示
!PY_CMD! --version

echo.
echo 必要なライブラリをインストールします...
echo.

if exist "%~dp0wheels" (
    echo オフラインインストール（wheelsフォルダ使用）...
    !PY_CMD! -m pip install --no-index --find-links "%~dp0wheels" -r "%~dp0requirements.txt"
) else (
    echo オンラインインストール（インターネット接続が必要）...
    !PY_CMD! -m pip install -r "%~dp0requirements.txt"
)

if !errorlevel! neq 0 (
    echo.
    echo [エラー] インストールに失敗しました。
    echo 管理者として実行するか、手動でインストールしてください。
    pause
    popd
    exit /b 1
)

echo.
echo インストール完了！
echo.
echo 「傾斜区分分析_起動.bat」をダブルクリックして使用開始してください。
popd
pause
