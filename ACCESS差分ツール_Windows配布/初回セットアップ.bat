@echo off
setlocal
cd /d "%~dp0"

echo Installing required packages.
echo.

set "PY_CMD="

where py >nul 2>nul
if errorlevel 1 goto try_python
py -3.12 -c "import sys; raise SystemExit(0 if sys.maxsize > 2**32 else 1)" >nul 2>nul
if not errorlevel 1 set "PY_CMD=py -3.12"
goto python_checked

:try_python
where python >nul 2>nul
if errorlevel 1 goto python_checked
python -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 12) and sys.maxsize > 2**32 else 1)" >nul 2>nul
if not errorlevel 1 set "PY_CMD=python"

:python_checked
if not defined PY_CMD (
  echo.
  echo Python 3.12 64-bit was not found.
  echo Install Python 3.12 64-bit and run this file again.
  echo.
  pause
  exit /b 1
)

if exist wheels (
  echo Using local wheels folder for offline install.
  %PY_CMD% -m pip install --no-index --find-links "%~dp0wheels" -r requirements.txt
  if not errorlevel 1 goto done
)

echo.
echo Offline install failed. Trying normal pip install.
%PY_CMD% -m pip install -r requirements.txt
if not errorlevel 1 goto done

echo.
echo Setup failed.
echo Check Python 3.12 64-bit.
echo Microsoft Access Database Engine 64-bit is also required for MDB files.
echo.
pause
exit /b 1

:done
echo.
echo Setup completed.
echo Microsoft Access Database Engine 64-bit is required for MDB files.
echo Run ACCESS tool launcher next time.
echo.
pause
