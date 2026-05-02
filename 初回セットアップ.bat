@echo off
cd /d "%~dp0"
echo 必要ライブラリをオフラインでインストールします。
echo wheels フォルダ内のファイルを使用します。
echo.

if not exist wheels (
  echo wheels フォルダが見つかりません。
  echo 配布フォルダを丸ごとコピーしたか確認してください。
  echo.
  pause
  exit /b 1
)

python3 -m pip install --no-index --find-links "%~dp0wheels" -r requirements.txt
if not errorlevel 1 goto done

py -3.12 -m pip install --no-index --find-links "%~dp0wheels" -r requirements.txt
if not errorlevel 1 goto done

py -3 -m pip install --no-index --find-links "%~dp0wheels" -r requirements.txt
if not errorlevel 1 goto done

python -m pip install --no-index --find-links "%~dp0wheels" -r requirements.txt
if not errorlevel 1 goto done

echo.
echo セットアップに失敗しました。
echo Python 3.12 64bit がインストールされているか確認してください。
echo.
pause
exit /b 1

:done
echo.
echo セットアップが完了しました。
echo 次回からは DXF帳票ツール起動.bat をダブルクリックしてください。
echo.
pause
