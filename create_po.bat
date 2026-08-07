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
echo   [N] DN 출고기록 자동 입력 (출고리스트 → DN_국내)
echo.
echo   [분석]
echo   [D] 대시보드
echo   [R] PO 매입대사 (Reconciliation)
echo   [S] SO 매출대사 (Sales Reconciliation)
echo   [I] Industry Code 대사
echo.
echo   [기타]
echo   [C] 거래처 납기현황 조회 (미출고 회신용)
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
if /i "%CHOICE%"=="N" goto create_dn
if /i "%CHOICE%"=="D" goto dashboard
if /i "%CHOICE%"=="R" goto reconcile
if /i "%CHOICE%"=="S" goto reconcile_so
if /i "%CHOICE%"=="I" goto reconcile_ind

if /i "%CHOICE%"=="C" goto delivery_status
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

:delivery_status
echo.
echo ----------------------------------------
echo   거래처 납기현황 조회
echo ----------------------------------------
echo.
echo   사업자등록번호(하이픈 무관) 또는 거래처명 일부를 입력하세요.
echo   그냥 Enter를 누르면 미출고가 있는 거래처 목록을 보여줍니다.
echo.
echo   생성 후 수신자를 보여주고 이메일 발송 여부를 물어봅니다.
echo.

set "DS_CUSTOMER="
set /p DS_CUSTOMER="거래처: "

echo.
REM 괄호 블록 대신 라벨로 분기한다. 거래처명에 '(주)'가 흔한데,
REM if 괄호블록 안에서 DS_CUSTOMER가 전개되면 괄호 짝이 깨져 배치가 죽는다.
REM 'if not defined'는 값을 전개하지 않아 특수문자가 섞여도 안전하다.
if not defined DS_CUSTOMER goto ds_list

"%PYTHON_PATH%" "%~dp0delivery_status.py" "%DS_CUSTOMER%"
goto ds_done

:ds_list
"%PYTHON_PATH%" "%~dp0delivery_status.py" --list

:ds_done
echo.
pause
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
REM 납기현황(:delivery_status)과 같은 방식으로 'if not defined' + 라벨 분기를 쓴다.
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
goto menu

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
echo   [1] DN_ID 기준 생성
echo   [2] 발주번호 기준 생성 (복수 DN 통합)
echo   [0] 메뉴로 돌아가기
echo.

set /p FI_MODE="선택: "

if "%FI_MODE%"=="1" goto fi_by_dn
if "%FI_MODE%"=="2" goto fi_by_po
if "%FI_MODE%"=="0" goto menu
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
goto menu

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

:create_dn
echo.
echo ----------------------------------------
echo   DN 출고기록 자동 입력
echo ----------------------------------------
echo.
echo   공장 출고리스트(2026리스트_RCK_Pxx.xlsx)를 읽어
echo   NOAH_SO_PO_DN.xlsx의 DN_국내에 출고 내역을 추가합니다.
echo.
echo   [1] 미리보기만 (워크북은 건드리지 않음)
echo   [2] 미리보기 후 확인하고 추가
echo   [0] 메뉴로 돌아가기
echo.

set /p DN_MODE="선택: "

if "%DN_MODE%"=="1" goto dn_dry
if "%DN_MODE%"=="2" goto dn_write
if "%DN_MODE%"=="0" goto menu
echo [오류] 올바른 번호를 입력하세요.
pause
goto create_dn

:dn_dry
echo.
:dn_dry_input
set /p DN_PERIOD="기간 코드 입력 (예: P08): "
if "%DN_PERIOD%"=="" (
    echo [오류] 기간 코드를 입력하세요.
    goto dn_dry_input
)

echo.
"%PYTHON_PATH%" "%~dp0create_dn.py" %DN_PERIOD% --dry-run

echo.
pause
goto menu

:dn_write
echo.
echo   추가할 행을 먼저 보여주고, 진행 여부를 다시 묻습니다.
echo   쓰기 전 워크북 사본이 generated_dn\backup 에 저장됩니다.
echo.
:dn_write_input
set /p DN_PERIOD="기간 코드 입력 (예: P08): "
if "%DN_PERIOD%"=="" (
    echo [오류] 기간 코드를 입력하세요.
    goto dn_write_input
)

echo.
"%PYTHON_PATH%" "%~dp0create_dn.py" %DN_PERIOD%

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
echo   [1] Industry Code 채움 + Sector 검증 (전체)
echo   [2] Sector 검증만
echo   [0] 메뉴로 돌아가기
echo.

set /p IND_MODE="선택: "

if "%IND_MODE%"=="1" goto ind_full
if "%IND_MODE%"=="2" goto ind_sector
if "%IND_MODE%"=="0" goto menu
echo [오류] 올바른 번호를 입력하세요.
pause
goto reconcile_ind

:ind_full
echo.
:ind_full_input
set /p RECON_IND_PERIOD="대사 월 입력 (예: P03): "

if "%RECON_IND_PERIOD%"=="" (
    echo [오류] 월 코드를 입력하세요.
    goto ind_full_input
)

echo.
echo Industry Code 대사 실행 중...
echo.

"%PYTHON_PATH%" "%~dp0reconcile_ind.py" %RECON_IND_PERIOD%

echo.
pause
goto menu

:ind_sector
echo.
echo Sector 검증 실행 중...
echo.

"%PYTHON_PATH%" "%~dp0reconcile_ind.py" --sector-only

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
