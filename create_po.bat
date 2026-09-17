@echo off
chcp 65001 >nul
title NOAH PO Generator

REM 문서 7종([국내]/[해외])의 입력·실행 블록은 noah_menu.bat 에 있다 — 여기서는
REM `call noah_menu.bat <키>` 로 위임한다. 동료 PC 배포판도 같은 파일을 쓰므로
REM 문서 옵션을 더할 때는 noah_menu.bat 한 곳만 고치면 두 메뉴에 같이 반영된다.

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
call "%~dp0noah_menu.bat" po
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
call "%~dp0noah_menu.bat" ts
goto menu

:create_pi
call "%~dp0noah_menu.bat" pi
goto menu

:create_fi
call "%~dp0noah_menu.bat" fi
goto menu

:create_oc
call "%~dp0noah_menu.bat" oc
goto menu

:create_ci
call "%~dp0noah_menu.bat" ci
goto menu

:create_pl
call "%~dp0noah_menu.bat" pl
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
