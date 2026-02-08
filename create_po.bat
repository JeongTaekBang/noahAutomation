@echo off
chcp 65001 >nul
title NOAH PO Generator

REM 사용자별 설정 파일 로드 (local_config.bat)
if exist "%~dp0local_config.bat" (
    call "%~dp0local_config.bat"
) else (
    echo.
    echo [경고] local_config.bat 파일이 없습니다.
    echo        local_config.example.bat 을 복사해서 local_config.bat 으로 만드세요.
    echo        그리고 본인의 Python 경로를 설정하세요.
    echo.
    pause
    exit /b 1
)

:menu
cls
echo ========================================
echo    NOAH Document Generator
echo ========================================
echo.
echo   [국내]
echo   [1] 발주서 생성 (PO)
echo   [2] 거래명세표 생성 (DN/선수금)
echo.
echo   [해외]
echo   [3] Proforma Invoice 생성 (PI)
echo.
echo   [기타]
echo   [8] 발주 이력 조회
echo   [9] 발주 이력 Excel 내보내기
echo   [0] 종료
echo.
echo ========================================
echo.

set /p CHOICE="선택 (0-9): "

if "%CHOICE%"=="1" goto create_po
if "%CHOICE%"=="2" goto create_ts
if "%CHOICE%"=="3" goto create_pi
if "%CHOICE%"=="8" goto view_history
if "%CHOICE%"=="9" goto export_history
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

:create_ts
echo.
echo ----------------------------------------
echo   거래명세표 생성 (국내 전용)
echo ----------------------------------------
echo.
echo   [1] 단건 거래명세표 (DN 1건)
echo   [2] 월합 거래명세표 (여러 DN을 한 장으로)
echo   [0] 메뉴로 돌아가기
echo.

set /p TS_MODE="선택: "

if "%TS_MODE%"=="1" goto ts_single
if "%TS_MODE%"=="2" goto ts_merge
if "%TS_MODE%"=="0" goto menu
echo [오류] 올바른 번호를 입력하세요.
pause
goto create_ts

:ts_single
echo.
echo   - 납품: DN_ID (예: DND-2026-0001)
echo   - 선수금: 선수금_ID (예: ADV_2026-0001)
echo.

:ts_input
set /p TS_DOC_ID="ID 입력: "

if "%TS_DOC_ID%"=="" (
    echo [오류] ID를 입력하세요.
    goto ts_input
)

echo.
echo 거래명세표 생성 중...
echo.

"%PYTHON_PATH%" "%~dp0create_ts.py" %TS_DOC_ID%

echo.
echo ----------------------------------------
set /p TS_CONTINUE="다른 거래명세표를 생성하시겠습니까? (Y/N): "
if /i "%TS_CONTINUE%"=="Y" goto ts_input
goto menu

:ts_merge
echo.
echo ----------------------------------------
echo   월합 거래명세표 (여러 DN을 한 장으로)
echo ----------------------------------------
echo.
echo   DN_ID 목록을 세로로 붙여넣기 하세요.
echo   (빈 줄 입력하면 생성 시작)
echo.

"%PYTHON_PATH%" "%~dp0create_ts.py" --interactive --merge

echo.
pause
goto menu

:create_pi
echo.
echo ----------------------------------------
echo   Proforma Invoice 생성 (해외)
echo ----------------------------------------
echo.
echo   SO_ID 입력 (예: SOO-2026-0001)
echo.

:pi_input
set /p PI_SO_ID="SO_ID 입력: "

if "%PI_SO_ID%"=="" (
    echo [오류] SO_ID를 입력하세요.
    goto pi_input
)

echo.
echo Proforma Invoice 생성 중...
echo.

"%PYTHON_PATH%" "%~dp0create_pi.py" %PI_SO_ID%

echo.
echo ----------------------------------------
set /p PI_CONTINUE="다른 Proforma Invoice를 생성하시겠습니까? (Y/N): "
if /i "%PI_CONTINUE%"=="Y" goto pi_input
goto menu

:end
echo.
echo 프로그램을 종료합니다.
pause
