@echo off
setlocal EnableExtensions
title Price Planning pull
rem One-step office pull. Run it in the folder that holds the template .xlsm: it gets Git for
rem this user once (no admin), then puts the app in app\ or updates it there.
set "HOME_DIR=%~dp0"
if exist "%~dp0office_run.bat" for %%I in ("%~dp0..") do set "HOME_DIR=%%~fI\"
set "GIT=%LOCALAPPDATA%\MinGit\cmd\git.exe"
if exist "%GIT%" goto have_git
echo Getting Git for this user, one time...
curl.exe --ssl-no-revoke -L -o "%TEMP%\mingit.zip" https://github.com/git-for-windows/git/releases/download/v2.56.0.windows.2/MinGit-2.56.0.2-64-bit.zip
if errorlevel 1 goto git_failed
if not exist "%LOCALAPPDATA%\MinGit" mkdir "%LOCALAPPDATA%\MinGit"
tar -xf "%TEMP%\mingit.zip" -C "%LOCALAPPDATA%\MinGit"
if not exist "%GIT%" goto git_failed
:have_git
rem Windows certificate store, like curl.exe here; the office proxy breaks revocation checks
rem The app folder may have been set up by another Windows user; git refuses such folders otherwise
set GITCMD="%GIT%" -c http.sslBackend=schannel -c http.schannelCheckRevoke=false -c core.autocrlf=false -c safe.directory=*
if exist "%HOME_DIR%app\.git" goto update
%GITCMD% clone --branch feat/v2-integration https://github.com/GopalGB/price-planning.git "%HOME_DIR%app"
if errorlevel 1 goto pull_failed
goto done
:update
rem The app folder holds only the app's own files: a local edit there would block every update.
rem One block: cmd reads it whole before the pull replaces this file under it.
(
  %GITCMD% -C "%HOME_DIR%app" checkout -- .
  %GITCMD% -C "%HOME_DIR%app" pull --ff-only || goto pull_failed
  echo.
  echo App is up to date in %HOME_DIR%app
  echo Start it with:  app\office_run.bat
  exit /b 0
)
:done
echo.
echo App is up to date in %HOME_DIR%app
echo Start it with:  app\office_run.bat
exit /b 0
:git_failed
echo Could not get Git. Check the internet connection, then run this again.
pause
exit /b 1
:pull_failed
echo Could not update the app. Read the lines above for the reason, then run this again.
pause
exit /b 1
