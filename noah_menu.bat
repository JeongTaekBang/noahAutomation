@echo off
chcp 65001 >nul
title NOAH 문서 생성기

REM ============================================================
REM  NOAH 문서 생성기 — 문서 메뉴 (국내 / 해외)
REM
REM  동료 PC용 진입점. 더블클릭하면 메뉴가 뜬다.
REM    - Python: 배포판이면 동봉된 python\ , 개발 PC면 local_config.bat
REM    - 데이터 파일(NOAH_SO_PO_DN.xlsx): 첫 실행 때 setup_data_path.py가
REM      OneDrive에서 찾아 noah_config.ini에 적는다 (GUI와 같은 ini)
REM
REM  create_po.bat(전체 메뉴)은 문서 항목을 여기에 위임한다:
REM      call noah_menu.bat po      ← 그 항목만 실행하고 돌아온다
REM  문서 생성 블록은 이 파일 한 곳에만 둔다 — 옵션을 더할 때 두 메뉴가 갈리지 않게.
REM ============================================================

setlocal
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

if defined PYTHON_PATH goto py_ready
if exist "%~dp0python\python.exe" (
    set "PYTHON_PATH=%~dp0python\python.exe"
    goto py_ready
)
if exist "%~dp0local_config.bat" (
    call "%~dp0local_config.bat"
    goto py_ready
)
echo.
echo [오류] Python을 찾지 못했습니다.
echo        배포판이면 python\ 폴더가, 개발 PC면 local_config.bat 이 있어야 합니다.
pause
exit /b 1

:py_ready
REM 데이터 파일이 정해져 있지 않으면 여기서 찾고 확인받는다 (이미 있으면 바로 지나간다)
"%PYTHON_PATH%" "%~dp0setup_data_path.py"
if errorlevel 1 (
    echo.
    echo 데이터 파일이 지정되지 않아 종료합니다.
    pause
    exit /b 1
)

REM 인자가 있으면 그 항목만 실행하고 돌아온다 (create_po.bat 위임용)
set "ONESHOT="
if not "%~1"=="" (
    set "ONESHOT=1"
    set "CHOICE=%~1"
    goto dispatch
)

:menu
cls
echo ========================================
echo    NOAH 문서 생성기
echo ========================================
"%PYTHON_PATH%" "%~dp0setup_data_path.py" --print
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
echo   [P] 데이터 파일 경로 변경
echo   [0] 종료
echo.
echo ========================================
echo.

set "CHOICE="
set /p CHOICE="선택: "

:dispatch
if "%CHOICE%"=="1" goto create_po
if "%CHOICE%"=="2" goto create_ts
if "%CHOICE%"=="3" goto create_pi
if "%CHOICE%"=="4" goto create_fi
if "%CHOICE%"=="5" goto create_oc
if "%CHOICE%"=="6" goto create_ci
if "%CHOICE%"=="7" goto create_pl
if /i "%CHOICE%"=="po" goto create_po
if /i "%CHOICE%"=="ts" goto create_ts
if /i "%CHOICE%"=="pi" goto create_pi
if /i "%CHOICE%"=="fi" goto create_fi
if /i "%CHOICE%"=="oc" goto create_oc
if /i "%CHOICE%"=="ci" goto create_ci
if /i "%CHOICE%"=="pl" goto create_pl
if /i "%CHOICE%"=="P" goto change_path
if "%CHOICE%"=="0" goto end
echo [오류] 올바른 번호를 입력하세요.
pause
goto back

:back
REM 위임 호출이면 한 항목으로 끝, 아니면 메뉴로
if defined ONESHOT exit /b 0
goto menu

:change_path
echo.
echo ----------------------------------------
echo   데이터 파일 경로 변경
echo ----------------------------------------

"%PYTHON_PATH%" "%~dp0setup_data_path.py" --change

echo.
pause
goto back

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
goto back

:create_ts
echo.
echo ----------------------------------------
echo   거래명세표 생성 (국내 전용)
echo ----------------------------------------
echo.
echo   [1] 단건 거래명세표 (DN 1건)
echo   [2] 월합 거래명세표 (여러 DN을 한 장으로)
echo   [3] 하루치 묶음 발송 (날짜로 골라 거래처별 메일 1통)
echo   [4] 묶음 발송 (DN 목록 붙여넣기 - 거래처별 메일 1통)
echo   [0] 메뉴로 돌아가기
echo.
echo   [2]는 '문서'를 한 장으로 합치고, [3][4]는 문서는 DN별 1장 그대로 두고
echo   '메일'만 거래처별 한 통(첨부 여러 개)으로 묶습니다.
echo.

set /p TS_MODE="선택: "

if "%TS_MODE%"=="1" goto ts_single
if "%TS_MODE%"=="2" goto ts_merge
if "%TS_MODE%"=="3" goto ts_batch
if "%TS_MODE%"=="4" goto ts_one_mail
if "%TS_MODE%"=="0" goto back
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
goto back

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
goto back

:ts_batch
echo.
echo ----------------------------------------
echo   하루치 묶음 발송 (문서는 DN별 1장, 메일은 거래처별 1통)
echo ----------------------------------------
echo.
echo   그날 출고분을 자동으로 찾아 거래처별로 묶어 보냅니다.
echo   (예: 8/6 씨앤케이 8건 - 문서 8장, 메일 1통에 첨부 8개)
echo.
echo   거래처를 지정하면 발송 전 y/N을 한 번 묻고,
echo   비워 두면 그날 전체를 거래처별로 나눠 확인 없이 메일 초안을 띄웁니다.
echo   (초안까지만 열립니다 - 보내기는 메일 창에서 직접)
echo.

set "TS_DATE="
set /p TS_DATE="출고일 (예: 2026-08-06 또는 8/6): "
if not defined TS_DATE goto ts_batch_no_date

REM 거래처명에 '(주)'가 흔하다 — if 괄호블록 안에서 전개되면 괄호 짝이 깨져 배치가 죽으므로
REM 'if not defined' + 라벨 분기를 쓴다.
set "TS_CUSTOMER="
set /p TS_CUSTOMER="거래처 (Enter=그날 전체): "

echo.
echo 거래명세표 생성 중...
echo.

if not defined TS_CUSTOMER goto ts_batch_all

"%PYTHON_PATH%" "%~dp0create_ts.py" --date "%TS_DATE%" --customer "%TS_CUSTOMER%"
goto ts_batch_done

:ts_batch_all
REM 거래처를 안 고른 경우는 거래처마다 y/N을 묻게 되므로(그날 5곳이면 5번) 확인을 건너뛴다.
REM --mail은 '초안 열기'까지다 — 자동 발송(--send)이 아니라 보내기는 사람이 누른다.
"%PYTHON_PATH%" "%~dp0create_ts.py" --date "%TS_DATE%" --mail

:ts_batch_done
echo.
pause
goto back

:ts_batch_no_date
echo [오류] 출고일을 입력하세요.
pause
goto ts_batch

:ts_one_mail
echo.
echo ----------------------------------------
echo   묶음 발송 (문서는 DN별 1장, 메일은 거래처별 1통)
echo ----------------------------------------
echo.
echo   DN_ID 목록을 세로로 붙여넣기 하세요.
echo   (빈 줄 입력하면 생성 시작)
echo.

"%PYTHON_PATH%" "%~dp0create_ts.py" --interactive --one-mail

echo.
pause
goto back

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
goto back

:create_fi
echo.
echo ----------------------------------------
echo   Final Invoice 생성 (대금 청구)
echo ----------------------------------------
echo.
echo   [1] DN_ID 기준 생성
echo   [2] 발주번호 기준 생성 (복수 DN 통합)
echo   [0] 메뉴로 돌아가기
echo.

set /p FI_MODE="선택: "

if "%FI_MODE%"=="1" goto fi_by_dn
if "%FI_MODE%"=="2" goto fi_by_po
if "%FI_MODE%"=="0" goto back
echo [오류] 올바른 번호를 입력하세요.
pause
goto create_fi

:fi_by_dn
echo.
echo   DN_ID 입력 (예: DNO-2026-0001)
echo.

:fi_dn_input
set /p FI_DN_ID="DN_ID 입력: "

if "%FI_DN_ID%"=="" (
    echo [오류] DN_ID를 입력하세요.
    goto fi_dn_input
)

echo.
echo Final Invoice 생성 중...
echo.

"%PYTHON_PATH%" "%~dp0create_fi.py" %FI_DN_ID%

echo.
echo ----------------------------------------
set /p FI_DN_CONT="다른 Final Invoice를 생성하시겠습니까? (Y/N): "
if /i "%FI_DN_CONT%"=="Y" goto fi_dn_input
goto back

:fi_by_po
echo.
echo   발주번호 입력 (예: 26KPO00144)
echo   (빈 입력 시 사용 가능한 PO 목록 표시)
echo.

:fi_po_input
set /p FI_RCK_PO="RCK PO 입력: "

if "%FI_RCK_PO%"=="" (
    "%PYTHON_PATH%" "%~dp0create_fi.py" --po
    echo.
    goto fi_po_input
)

echo.
echo Final Invoice 생성 중 (RCK PO: %FI_RCK_PO%)...
echo.

"%PYTHON_PATH%" "%~dp0create_fi.py" --po %FI_RCK_PO%

echo.
echo ----------------------------------------
set /p FI_PO_CONT="다른 발주번호로 생성하시겠습니까? (Y/N): "
if /i "%FI_PO_CONT%"=="Y" goto fi_po_input
goto back

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
goto back

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
goto back

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
goto back

:end
echo.
echo 프로그램을 종료합니다.
exit /b 0
