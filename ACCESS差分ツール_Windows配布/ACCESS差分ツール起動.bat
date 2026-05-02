@echo off
cd /d "%~dp0"
python3 access_diff_web.py
if not errorlevel 1 goto :eof

py -3 access_diff_web.py
if not errorlevel 1 goto :eof

python access_diff_web.py
if not errorlevel 1 goto :eof

echo.
echo Python 3 で起動できませんでした。
echo Python 3.10以上をインストールし、もう一度実行してください。
echo.
pause
