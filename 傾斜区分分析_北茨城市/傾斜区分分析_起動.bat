@echo off
chcp 932 >nul 2>&1
setlocal EnableDelayedExpansion
pushd "%~dp0"

REM Python実行コマンド検出
set PY_CMD=
for %%P in (py python python3) do (
    if "!PY_CMD!"=="" (
        %%P --version >nul 2>&1
        if !errorlevel! == 0 (
            set PY_CMD=%%P
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

REM DEMデータの存在確認
if not exist "%~dp0data\kitaibaraki_dem.npz" (
    echo [エラー] 地形データが見つかりません。
    echo data\kitaibaraki_dem.npz が必要です。
    echo.
    echo 初回セットアップ.bat を先に実行してください。
    pause
    popd
    exit /b 1
)

REM 必須ライブラリ確認
!PY_CMD! -c "import numpy, matplotlib" >nul 2>&1
if !errorlevel! neq 0 (
    echo 必要なライブラリをインストールします...
    if exist "%~dp0wheels" (
        !PY_CMD! -m pip install --no-index --find-links "%~dp0wheels" -r "%~dp0requirements.txt"
    ) else (
        !PY_CMD! -m pip install -r "%~dp0requirements.txt"
    )
    if !errorlevel! neq 0 (
        echo.
        echo [エラー] ライブラリのインストールに失敗しました。
        echo 初回セットアップ.bat を管理者として実行してください。
        pause
        popd
        exit /b 1
    )
)

echo.
echo 傾斜区分分析ツールを起動中...
echo ブラウザが自動で開きます。
echo このウィンドウを閉じると終了します。
echo.
!PY_CMD! "%~dp0terrain_web.py"

popd
pause
