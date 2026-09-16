@echo off
rem ===========================================================
rem  File sync checker - interactive menu launcher
rem  Double-click this file to start the menu.
rem
rem  Messages here are ASCII on purpose: this file must stay
rem  readable in any console code page.
rem ===========================================================

rem Move to this file's folder. pushd (not cd) so that a UNC
rem path on a shared drive also works.
pushd "%~dp0" || (
    echo [ERROR] Cannot enter the folder of this script.
    pause
    exit /b 1
)

rem Prefer the py launcher; fall back to python on PATH.
set PYTHON=
py -3 --version >nul 2>&1 && set PYTHON=py -3
if not defined PYTHON (
    python --version >nul 2>&1 && set PYTHON=python
)
if not defined PYTHON (
    echo [ERROR] Python 3 was not found.
    echo         Install Python 3.9 or later from https://www.python.org/
    echo         and make sure it is added to PATH.
    popd
    pause
    exit /b 1
)

%PYTHON% menu.py %*
set RC=%ERRORLEVEL%

popd
if not "%RC%"=="0" echo.& echo [exit code %RC%]
pause
exit /b %RC%
