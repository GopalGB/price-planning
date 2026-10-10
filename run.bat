@echo off
setlocal
title Catalyst Brands Price Planning
pushd "%~dp0"
if errorlevel 1 exit /b 1

rem Use an approved local Python, the Windows Python launcher, or Python on PATH.
if not exist "python\python.exe" goto check_launcher
"python\python.exe" -c "import sys; sys.exit(not (sys.version_info >= (3, 10)))" >nul 2>nul
if not errorlevel 1 goto local_python

:check_launcher
where py >nul 2>nul
if errorlevel 1 goto check_path
py -3 -c "import sys; sys.exit(not (sys.version_info >= (3, 10)))" >nul 2>nul
if not errorlevel 1 goto python_launcher

:check_path
where python >nul 2>nul
if errorlevel 1 goto missing_python
python -c "import sys; sys.exit(not (sys.version_info >= (3, 10)))" >nul 2>nul
if not errorlevel 1 goto path_python
goto missing_python

:local_python
set "VENV_PY="
for /f "delims=" %%P in ('"python\python.exe" ops\install_env.py') do set "VENV_PY=%%P"
if not defined VENV_PY ( echo Dependency install failed. & set "planning_exit=1" & goto leave )
"%VENV_PY%" -u "run.py" --demo %*
goto finished

:python_launcher
set "VENV_PY="
for /f "delims=" %%P in ('py -3 ops\install_env.py') do set "VENV_PY=%%P"
if not defined VENV_PY ( echo Dependency install failed. & set "planning_exit=1" & goto leave )
"%VENV_PY%" -u "run.py" --demo %*
goto finished

:path_python
set "VENV_PY="
for /f "delims=" %%P in ('python ops\install_env.py') do set "VENV_PY=%%P"
if not defined VENV_PY ( echo Dependency install failed. & set "planning_exit=1" & goto leave )
"%VENV_PY%" -u "run.py" --demo %*
goto finished

:missing_python
echo Python 3.10 or newer is required. Python 3.12 is recommended.
echo Use your company's approved Python installation, then run this file again.
echo The app itself does not need an administrator installation or internet access.
set "planning_exit=1"
goto leave

:finished
set "planning_exit=%errorlevel%"
if not "%planning_exit%"=="0" echo Startup failed. Read the message above and run again after correcting it.

:leave
if "%~1"=="" pause
popd
exit /b %planning_exit%
