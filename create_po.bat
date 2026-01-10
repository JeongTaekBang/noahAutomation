@echo off
chcp 65001 >nul
title NOAH PO Generator

set PYTHON_PATH=%USERPROFILE%\anaconda3\envs\po-automate\python.exe

echo ========================================
echo    NOAH Purchase Order Generator
echo ========================================
echo.

:input
set /p ORDER_NO="Enter RCK Order No. (e.g. ND-0005): "

if "%ORDER_NO%"=="" (
    echo [Error] Please enter Order No.
    goto input
)

echo.
echo Generating PO...
echo.

"%PYTHON_PATH%" "%~dp0create_po.py" %ORDER_NO%

echo.
echo ----------------------------------------
set /p CONTINUE="Generate another PO? (Y/N): "
if /i "%CONTINUE%"=="Y" goto input

echo.
echo Done.
pause
