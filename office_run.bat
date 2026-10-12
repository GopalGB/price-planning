@echo off
setlocal EnableExtensions
title Price Planning
rem One-step office start. Layout: a folder holding the template .xlsm and the app\ clone.
rem   app\office_run.bat          set up on first use, then start the app
rem   app\office_run.bat check    run the four Excel checks again, then start the app
pushd "%~dp0"
if errorlevel 1 exit /b 1
for %%I in ("%~dp0..") do set "HOME_DIR=%%~fI"

rem 1. Template: the .xlsm placed next to the app folder (the first one by name).
set "XLSM="
for %%F in ("%HOME_DIR%\*.xlsm") do if not defined XLSM set "XLSM=%%~fF"
if not defined XLSM goto no_template
echo Template: %XLSM%
set "PP_TEMPLATE_PATH=%XLSM%"
set "PP_ENGINE_CACHE_DIR=%HOME_DIR%\engine-cache"

rem 2. Settings on first use: Excel engine, open to the office network.
if exist settings.json goto have_settings
> settings.json echo {"bind": "0.0.0.0", "port": 8000, "demo_mode": true, "engine": "excel"}
:have_settings

rem 3. Python 3.12, 3.11 or 3.10 (the bundled packages are built for these).
call :find_python
if defined PYEXE goto have_python
echo Getting Python 3.12 for this user, one time...
rem The python.org NuGet package is the full Python as a plain zip: no installer to run, no admin
curl.exe --ssl-no-revoke -L -o "%TEMP%\py312.zip" https://www.nuget.org/api/v2/package/python/3.12.10
if errorlevel 1 goto python_failed
if not exist "%LOCALAPPDATA%\PricePlanning\python312" mkdir "%LOCALAPPDATA%\PricePlanning\python312"
tar -xf "%TEMP%\py312.zip" -C "%LOCALAPPDATA%\PricePlanning\python312"
call :find_python
if not defined PYEXE goto python_failed
:have_python
echo Python: %PYEXE%

rem 4. Packages, offline from vendor\wheels (only when requirements.lock changed).
set "VENV_PY="
for /f "delims=" %%P in ('call "%PYEXE%" ops\install_env.py') do set "VENV_PY=%%P"
if not defined VENV_PY goto install_failed

rem 5. The four Excel checks: first run only, or "office_run.bat check".
rem Python's COM lookup reads this key; "Invalid class string" there means it is missing here
reg query "HKCR\Excel.Application\CLSID" >nul 2>&1
if errorlevel 1 goto no_excel
if /i "%~1"=="check" goto checks
if exist "%HOME_DIR%\checks-done.txt" goto start_app
:checks
echo.
echo Running the four Excel checks. The first builds caches: 20 to 40 minutes the first time.
set "FAILED=0"
for %%A in ("" "--macros" "--refresh" "--slicers") do (
  echo ===== CHECK %%~A =====
  "%VENV_PY%" ops\excel_smoke.py %%~A
  if errorlevel 1 set /a FAILED+=1
)
> "%HOME_DIR%\checks-done.txt" echo Excel checks run %DATE% %TIME%: %FAILED% of 4 did not pass
if not "%FAILED%"=="0" echo %FAILED% of 4 checks did not PASS. Send the lines above to Claude. Starting the app anyway.
goto start_app

:no_excel
echo.
echo Microsoft Excel is not installed on this computer, so the app opens VIEW-ONLY:
echo every tab with Excel's saved numbers, no editing, macros or refresh.
echo For the full app, run this on a computer with desktop Excel, or install Excel here.
echo.
set "PP_TEMPLATE_RESULTS=cached"
:start_app
echo.
echo Starting the app. Share the "Share with your team" link below. Ctrl+C stops it.
"%VENV_PY%" -u run_v2.py
goto leave

:find_python
set "PYEXE="
if exist "%LOCALAPPDATA%\PricePlanning\python312\tools\python.exe" set "PYEXE=%LOCALAPPDATA%\PricePlanning\python312\tools\python.exe"
for %%V in (312 311 310) do if not defined PYEXE if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe"
for %%V in (3.12 3.11 3.10) do if not defined PYEXE for /f "delims=" %%P in ('py -%%V -c "import sys; print(sys.executable)" 2^>nul') do set "PYEXE=%%P"
exit /b 0

:no_template
echo No .xlsm found in "%HOME_DIR%". Put the template there and run this again.
goto fail
:python_failed
echo Could not get Python 3.12. Check the internet connection, then run this again.
goto fail
:install_failed
echo Installing the packages failed. Send the lines above to Claude.
:fail
popd
pause
exit /b 1
:leave
popd
