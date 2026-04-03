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
echo   [4] Final Invoice 생성 (대금 청구)
echo   [5] Order Confirmation 생성 (OC)
echo   [6] Commercial Invoice 생성 (CI)
echo   [7] Packing List 생성 (PL)
echo.
echo   [데이터]
echo   [8] DB Sync (Excel → SQLite)
echo   [9] Order Book Close (월 마감)
echo.
echo   [분석]
echo   [D] 대시보드
echo   [R] PO 매입대사 (Reconciliation)
echo   [S] SO 매출대사 (Sales Reconciliation)
echo   [I] Industry Code 대사
echo.
echo   [기타]
echo   [H] 발주 이력 조회
echo   [0] 종료
echo.
echo ========================================
echo.

set /p CHOICE="선택: "

if "%CHOICE%"=="1" goto create_po
if "%CHOICE%"=="2" goto create_ts
if "%CHOICE%"=="3" goto create_pi
if "%CHOICE%"=="4" goto create_fi
if "%CHOICE%"=="5" goto create_oc
if "%CHOICE%"=="6" goto create_ci
if "%CHOICE%"=="7" goto create_pl
if "%CHOICE%"=="8" goto sync_db
if "%CHOICE%"=="9" goto close_period
if /i "%CHOICE%"=="D" goto dashboard
if /i "%CHOICE%"=="R" goto reconcile
if /i "%CHOICE%"=="S" goto reconcile_so
if /i "%CHOICE%"=="I" goto reconcile_ind

if /i "%CHOICE%"=="H" goto view_history
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

:create_fi
echo.
echo ----------------------------------------
echo   Final Invoice 생성 (대금 청구)
echo ----------------------------------------
echo.
echo   DN_ID 입력 (예: DNO-2026-0001)
echo.

:fi_input
set /p FI_DN_ID="DN_ID 입력: "

if "%FI_DN_ID%"=="" (
    echo [오류] DN_ID를 입력하세요.
    goto fi_input
)

echo.
echo Final Invoice 생성 중...
echo.

"%PYTHON_PATH%" "%~dp0create_fi.py" %FI_DN_ID%

echo.
echo ----------------------------------------
set /p FI_CONTINUE="다른 Final Invoice를 생성하시겠습니까? (Y/N): "
if /i "%FI_CONTINUE%"=="Y" goto fi_input
goto menu

:create_oc
echo.
echo ----------------------------------------
echo   Order Confirmation 생성 (해외)
echo ----------------------------------------
echo.
echo   SO_ID 입력 (예: SOO-2026-0001)
echo.

:oc_input
set /p OC_SO_ID="SO_ID 입력: "

if "%OC_SO_ID%"=="" (
    echo [오류] SO_ID를 입력하세요.
    goto oc_input
)

echo.
echo Order Confirmation 생성 중...
echo.

"%PYTHON_PATH%" "%~dp0create_oc.py" %OC_SO_ID%

echo.
echo ----------------------------------------
set /p OC_CONTINUE="다른 Order Confirmation을 생성하시겠습니까? (Y/N): "
if /i "%OC_CONTINUE%"=="Y" goto oc_input
goto menu

:create_pl
echo.
echo ----------------------------------------
echo   Packing List 생성 (해외)
echo ----------------------------------------
echo.
echo   DN_ID 입력 (예: DNO-2026-0001)
echo.

:pl_input
set /p PL_DN_ID="DN_ID 입력: "

if "%PL_DN_ID%"=="" (
    echo [오류] DN_ID를 입력하세요.
    goto pl_input
)

echo.
echo Packing List 생성 중...
echo.

"%PYTHON_PATH%" "%~dp0create_pl.py" %PL_DN_ID%

echo.
echo ----------------------------------------
set /p PL_CONTINUE="다른 Packing List를 생성하시겠습니까? (Y/N): "
if /i "%PL_CONTINUE%"=="Y" goto pl_input
goto menu

:create_ci
echo.
echo ----------------------------------------
echo   Commercial Invoice 생성 (해외)
echo ----------------------------------------
echo.
echo   DN_ID 입력 (예: DNO-2026-0001)
echo.

:ci_input
set /p CI_DN_ID="DN_ID 입력: "

if "%CI_DN_ID%"=="" (
    echo [오류] DN_ID를 입력하세요.
    goto ci_input
)

echo.
echo Commercial Invoice 생성 중...
echo.

"%PYTHON_PATH%" "%~dp0create_ci.py" %CI_DN_ID%

echo.
echo ----------------------------------------
set /p CI_CONTINUE="다른 Commercial Invoice를 생성하시겠습니까? (Y/N): "
if /i "%CI_CONTINUE%"=="Y" goto ci_input
goto menu

:sync_db
echo.
echo ----------------------------------------
echo   Excel → SQLite DB 동기화
echo ----------------------------------------
echo.

"%PYTHON_PATH%" "%~dp0sync_db.py" --changes

echo.
pause
goto menu

:close_period
echo.
echo ----------------------------------------
echo   Order Book 월 마감 (스냅샷)
echo ----------------------------------------
echo.
echo   [1] 월 마감
echo   [2] 마감 취소 (최신만)
echo   [3] 마감 현황 조회
echo   [4] 현재 상태
echo   [0] 메뉴로 돌아가기
echo.

set /p CP_MODE="선택: "

if "%CP_MODE%"=="1" goto cp_close
if "%CP_MODE%"=="2" goto cp_undo
if "%CP_MODE%"=="3" goto cp_list
if "%CP_MODE%"=="4" goto cp_status
if "%CP_MODE%"=="0" goto menu
echo [오류] 올바른 번호를 입력하세요.
pause
goto close_period

:cp_close
echo.
set /p CP_PERIOD="마감할 Period 입력 (예: 2026-01): "
if "%CP_PERIOD%"=="" (
    echo [오류] Period를 입력하세요.
    goto cp_close
)
set /p CP_NOTE="비고 (없으면 Enter): "

echo.
echo 마감 처리 중...
echo.

if "%CP_NOTE%"=="" (
    "%PYTHON_PATH%" "%~dp0close_period.py" %CP_PERIOD%
) else (
    "%PYTHON_PATH%" "%~dp0close_period.py" %CP_PERIOD% --note "%CP_NOTE%"
)

echo.
pause
goto menu

:cp_undo
echo.
set /p CP_UNDO_PERIOD="취소할 Period 입력 (예: 2026-01): "
if "%CP_UNDO_PERIOD%"=="" (
    echo [오류] Period를 입력하세요.
    goto cp_undo
)

echo.
"%PYTHON_PATH%" "%~dp0close_period.py" --undo %CP_UNDO_PERIOD%

echo.
pause
goto menu

:cp_list
echo.
"%PYTHON_PATH%" "%~dp0close_period.py" --list

echo.
pause
goto menu

:cp_status
echo.
"%PYTHON_PATH%" "%~dp0close_period.py" --status

echo.
pause
goto menu

:reconcile
echo.
echo ----------------------------------------
echo   PO 매입대사 (Reconciliation)
echo ----------------------------------------
echo.

:recon_input
set /p RECON_PERIOD="대사 월 입력 (예: P03): "

if "%RECON_PERIOD%"=="" (
    echo [오류] 월 코드를 입력하세요.
    goto recon_input
)

echo.
echo 매입대사 실행 중...
echo.

"%PYTHON_PATH%" "%~dp0reconcile_po.py" %RECON_PERIOD%

echo.
pause
goto menu

:reconcile_so
echo.
echo ----------------------------------------
echo   SO 매출대사 (Sales Reconciliation)
echo ----------------------------------------
echo.

:recon_so_input
set /p RECON_SO_PERIOD="대사 월 입력 (예: P03): "

if "%RECON_SO_PERIOD%"=="" (
    echo [오류] 월 코드를 입력하세요.
    goto recon_so_input
)

echo.
echo 매출대사 실행 중...
echo.

"%PYTHON_PATH%" "%~dp0reconcile_so.py" %RECON_SO_PERIOD%

echo.
pause
goto menu

:reconcile_ind
echo.
echo ----------------------------------------
echo   Industry Code 대사
echo ----------------------------------------
echo.

:recon_ind_input
set /p RECON_IND_PERIOD="대사 월 입력 (예: P03): "

if "%RECON_IND_PERIOD%"=="" (
    echo [오류] 월 코드를 입력하세요.
    goto recon_ind_input
)

echo.
echo Industry Code 대사 실행 중...
echo.

"%PYTHON_PATH%" "%~dp0reconcile_ind.py" %RECON_IND_PERIOD%

echo.
pause
goto menu

:dashboard
echo.
echo ----------------------------------------
echo   대시보드 (내 PC)
echo ----------------------------------------
echo.
echo   브라우저에서 대시보드가 열립니다.
echo   종료: Ctrl+C
echo.

"%PYTHON_PATH%" -m streamlit run "%~dp0dashboard.py"

echo.
pause
goto menu

:end
echo.
echo 프로그램을 종료합니다.
pause
