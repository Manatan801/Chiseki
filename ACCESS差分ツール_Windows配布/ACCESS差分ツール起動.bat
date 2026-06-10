@echo off
setlocal
cd /d "%~dp0"

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
  echo Install Python 3.12 64-bit, then run setup again.
  echo.
  pause
  exit /b 1
)

%PY_CMD% access_diff_web.py
set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" (
  echo.
  echo The tool stopped with error code %EXIT_CODE%.
  echo.
  pause
)
exit /b %EXIT_CODE%
