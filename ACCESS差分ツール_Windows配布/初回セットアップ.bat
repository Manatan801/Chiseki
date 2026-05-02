@echo off
cd /d "%~dp0"
echo 必要ライブラリをインストールします。
echo.

if exist wheels (
  echo wheels フォルダがあるため、オフラインインストールを試します。
  python3 -m pip install --no-index --find-links "%~dp0wheels" -r requirements.txt
  if not errorlevel 1 goto done

  py -3.12 -m pip install --no-index --find-links "%~dp0wheels" -r requirements.txt
  if not errorlevel 1 goto done

  py -3 -m pip install --no-index --find-links "%~dp0wheels" -r requirements.txt
  if not errorlevel 1 goto done

  python -m pip install --no-index --find-links "%~dp0wheels" -r requirements.txt
  if not errorlevel 1 goto done
)

echo.
echo オフラインインストールに失敗したため、通常インストールを試します。
python3 -m pip install -r requirements.txt
if not errorlevel 1 goto done

py -3.12 -m pip install -r requirements.txt
if not errorlevel 1 goto done

py -3 -m pip install -r requirements.txt
if not errorlevel 1 goto done

python -m pip install -r requirements.txt
if not errorlevel 1 goto done

echo.
echo セットアップに失敗しました。
echo Python 3.12 64bit がインストールされているか確認してください。
echo Access MDBを読むには Microsoft Access Database Engine 64bit も必要です。
echo.
pause
exit /b 1

:done
echo.
echo セットアップが完了しました。
echo Access MDBを読むには Microsoft Access Database Engine 64bit が必要です。
echo 次回からは ACCESS差分ツール起動.bat をダブルクリックしてください。
echo.
pause
