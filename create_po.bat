@echo off
chcp 65001 >nul
title NOAH PO Generator

set PYTHON_PATH=%USERPROFILE%\anaconda3\envs\po-automate\python.exe

:menu
cls
echo ========================================
echo    NOAH Purchase Order Generator
echo ========================================
echo.
echo   [1] 발주서 생성 (PO 생성)
echo   [2] 발주 이력 조회
echo   [3] 발주 이력 Excel 내보내기
echo   [0] 종료
echo.
echo ========================================
echo.

set /p CHOICE="선택 (0-3): "

if "%CHOICE%"=="1" goto create_po
if "%CHOICE%"=="2" goto view_history
if "%CHOICE%"=="3" goto export_history
if "%CHOICE%"=="0" goto end
echo [오류] 올바른 번호를 입력하세요.
pause
goto menu

:create_po
echo.
echo ----------------------------------------
echo   발주서 생성
echo ----------------------------------------
echo.

:input
set /p ORDER_NO="RCK Order No. 입력 (예: ND-0005): "

if "%ORDER_NO%"=="" (
    echo [오류] Order No.를 입력하세요.
    goto input
)

echo.
echo 발주서 생성 중...
echo.

"%PYTHON_PATH%" "%~dp0create_po.py" %ORDER_NO%

echo.
echo ----------------------------------------
set /p CONTINUE="다른 발주서를 생성하시겠습니까? (Y/N): "
if /i "%CONTINUE%"=="Y" goto input
goto menu

:view_history
echo.
echo ----------------------------------------
echo   발주 이력 조회
echo ----------------------------------------
echo.

"%PYTHON_PATH%" "%~dp0create_po.py" --history

echo.
pause
goto menu

:export_history
echo.
echo ----------------------------------------
echo   발주 이력 Excel 내보내기
echo ----------------------------------------
echo.

"%PYTHON_PATH%" "%~dp0create_po.py" --history --export

echo.
pause
goto menu

:end
echo.
echo 프로그램을 종료합니다.
pause
