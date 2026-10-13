@echo off
setlocal
title Price Planning v2
pushd "%~dp0"
if errorlevel 1 exit /b 1

rem Same Python resolution as run.bat: approved local Python, the Windows launcher, then PATH.
set "PYCMD="
if not exist "python\python.exe" goto check_launcher
"python\python.exe" -c "import sys; sys.exit(not (sys.version_info >= (3, 10)))" >nul 2>nul
if not errorlevel 1 set PYCMD="python\python.exe"
if defined PYCMD goto install

:check_launcher
where py >nul 2>nul
if errorlevel 1 goto check_path
py -3 -c "import sys; sys.exit(not (sys.version_info >= (3, 10)))" >nul 2>nul
if not errorlevel 1 set "PYCMD=py -3"
if defined PYCMD goto install

:check_path
where python >nul 2>nul
if errorlevel 1 goto missing_python
python -c "import sys; sys.exit(not (sys.version_info >= (3, 10)))" >nul 2>nul
if not errorlevel 1 set "PYCMD=python"
if defined PYCMD goto install
goto missing_python

:install
set "VENV_PY="
for /f "delims=" %%P in ('%PYCMD% ops\install_env.py') do set "VENV_PY=%%P"
if not defined VENV_PY ( echo Dependency install failed. & set "planning_exit=1" & goto leave )
if /i "%~1"=="--backup" goto backup
if /i "%~1"=="--restore" goto backup
"%VENV_PY%" -u "run_v2.py" %*
goto finished

:backup
"%VENV_PY%" -m pp.backup %*
goto finished

:missing_python
echo Python 3.10 or newer is required. Python 3.12 is recommended.
echo Use your company's approved Python installation, then run this file again.
echo The app itself does not need an administrator installation or internet access.
set "planning_exit=1"
goto leave

:finished
set "planning_exit=%errorlevel%"
if not "%planning_exit%"=="0" echo Failed. Read the message above and run again after correcting it.

:leave
if "%~1"=="" pause
popd
exit /b %planning_exit%
