"""
NOAH 문서 생성기 — 사내 배포판 빌드
=====================================

Python 런타임까지 통째로 담은 무설치 배포 zip을 만든다.
받는 사람은 압축을 풀고 `설치.bat`을 누르면 끝이다 (Python 설치 불필요).

Usage:
    python cli_dist/build_portable_gui.py

출력:
    cli_dist/NOAH_문서생성기_배포.zip
      ├── 설치.bat        ← 더블클릭 (LocalAppData로 복사 + 바탕화면 바로가기)
      ├── README.txt
      └── app/            ← 본체
          ├── NOAH 문서생성기.vbs   ← 평소 실행 (콘솔 없음)
          ├── NOAH 문서생성기.bat   ← 문제 진단용 (콘솔 표시)
          ├── 제거.bat              ← 설치 폴더 + 바탕화면 바로가기 삭제
          ├── BUILD_INFO.txt        ← 빌드 버전·구성 (GUI 타이틀·문의 대응용)
          ├── noah_gui.py, create_*.py, delivery_status.py
          ├── po_generator/, templates/
          └── python/               ← CPython 3.11 + tkinter + pandas/xlwings

빌드를 임시 폴더에서 하는 이유 두 가지
  1. 프로젝트 폴더가 OneDrive 안에 있다. 런타임 파일 수천 개를 여기에 만들면
     전부 동기화된다. 프로젝트에는 완성된 zip 하나만 남긴다.
  2. Windows MAX_PATH(260자). 프로젝트 경로가 이미 길어서 그 아래에
     python/Lib/site-packages/... 를 풀면 pip가 파일을 못 만든다 (실측 확인).

런타임으로 python.org 임베디드 배포본을 쓰지 않는 이유
  임베디드 배포본에는 tkinter가 없다(_tkinter.pyd, tcl/, Lib/tkinter 전부 부재).
  NuGet CPython 패키지에도 없다(1773개 엔트리 중 tk 관련 0건). 그래서 tkinter와
  pythonw.exe를 모두 포함하는 python-build-standalone 배포본을 쓴다.

런타임 내려받기·핀 대조·트리밍·zip 만들기는 대시보드 배포판과 똑같아서
`build_common.py`(프로젝트 루트)가 갖고 있다. 여기 남는 것은 이 배포판만의 것
— 무엇을 담고(APP_FILES), 어떻게 실행하고(런처), 무엇을 증명하는가(verify).

주의: 이 모듈의 최상위는 **상수와 함수 정의만** 있어야 한다.
  tests/test_noah_gui.py가 이 파일을 경로로 로드해 APP_FILES·parse_pins 등을
  검사한다 — 모듈 로드가 곧 빌드 시작이면 테스트가 빌드를 돌려버린다.
  (예외는 build_common을 찾기 위한 sys.path 한 줄뿐이다.)
"""

from __future__ import annotations

import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
PROJECT_ROOT = HERE.parent

# 공통부(build_common)는 프로젝트 루트에 있다. 이 파일은
# `python cli_dist/build_portable_gui.py`로 실행되므로 sys.path[0]이 cli_dist/ 다
# — 루트를 넣어 줘야 import가 된다. tests는 루트가 이미 sys.path에 있어
# 이 줄이 없어도 되지만, 중복 삽입은 무해하다.
sys.path.insert(0, str(PROJECT_ROOT))

from build_common import (  # noqa: E402 — 위 sys.path 배선 뒤여야 한다
    Stepper,
    canon,  # noqa: F401 — tests가 이 모듈의 속성으로 접근한다 (재export)
    compute_version,
    install_packages,
    install_runtime,
    make_zip,
    parse_pins,  # noqa: F401 — 위와 같음
    prepare,
    trim_runtime,
    verify_env,
    write_build_info,
    write_text,
)

ZIP_PATH = HERE / "NOAH_문서생성기_배포.zip"

# 빌드 작업 폴더 — 짧은 경로 + OneDrive 밖 (위 주석 참고)
BUILD_ROOT = Path(tempfile.gettempdir()) / "noah_build"
APP_DIR = BUILD_ROOT / "app"
PYTHON_DIR = APP_DIR / "python"

# === 런타임 ===
# 검증된 빌드를 고정한다. 배포판은 재현 가능해야 한다.
# 갱신하려면 https://github.com/astral-sh/python-build-standalone/releases 에서
# 새 태그의 cpython-3.11.*-x86_64-pc-windows-msvc-install_only.tar.gz 를 찾아 바꾼다.
RUNTIME_TAG = "20260718"
RUNTIME_FILE = "cpython-3.11.15%2B20260718-x86_64-pc-windows-msvc-install_only.tar.gz"
RUNTIME_URL = (
    "https://github.com/astral-sh/python-build-standalone/releases/download/"
    f"{RUNTIME_TAG}/{RUNTIME_FILE}"
)
RUNTIME_PY_VERSION = re.search(r"cpython-([\d.]+)", RUNTIME_FILE).group(1)  # type: ignore[union-attr]

# === 배포에 포함할 앱 파일 ===
# GUI가 실행하는 CLI는 전부 여기 있어야 한다. 빠지면 그 버튼만 배포판에서 조용히 실패하므로
# verify()가 noah_gui.DOC_TYPES와 대조해 zip 만들기 전에 잡는다.
APP_FILES = (
    "noah_gui.py",
    "create_po.py",
    "create_ts.py",
    "create_pi.py",
    "create_fi.py",
    "create_oc.py",
    "create_ci.py",
    "create_pl.py",
    "delivery_status.py",
)
APP_DIRS = ("po_generator", "templates")

# 배포에서 반드시 빼야 하는 것
#   user_settings.py — 빌드한 사람의 OneDrive 경로가 박혀 있다. 이게 들어가면
#                      받는 사람 PC에서 ini보다 우선해서 남의 경로를 바라본다.
EXCLUDE_FILES = {"user_settings.py"}

# === 런타임 트리밍 ===
# 기준: repo 전체 grep으로 import 0건 확인된 것만 지운다 (2026-07-30 실측 ~60MB 추가 절감).
# 지운 뒤 verify()가 전체 import + CLI 8종 --help로 안전을 증명한다.
TRIM_GLOBS = (
    "**/*.pdb",
    # setuptools를 지우면 이 .pth도 반드시 함께 — 고아 .pth는 site가 매 실행마다
    # `__import__('_distutils_hack')` 실패를 stderr에 찍는다
    "Lib/site-packages/distutils-precedence.pth",
    # 지운 패키지의 dist-info가 남으면 importlib.metadata가 유령 패키지를 보고한다
    "Lib/site-packages/pip-*.dist-info",
    "Lib/site-packages/setuptools-*.dist-info",
    "Lib/site-packages/PyWin32.chm",      # 도움말 파일 (2.6MB)
    "Lib/site-packages/numpy/**/tests",   # 테스트 16.9MB — pandas/tests와 같은 이유
    "tcl/*.lib",                          # C 링크 스텁 — 런타임 불필요
)
TRIM_DIRS = (
    "Lib/test",
    "Lib/idlelib",
    "Lib/ensurepip",                      # pip 부트스트랩 — 배포판은 pip을 쓰지 않는다
    "Lib/site-packages/pandas/tests",
    "Lib/site-packages/pip",              # 12.7MB — 런타임에 패키지 설치할 일 없음
    "Lib/site-packages/setuptools",       # 9MB — ↑ distutils-precedence.pth와 세트
    "Lib/site-packages/_distutils_hack",
    "Lib/site-packages/pythonwin",        # Pythonwin IDE (mfc140u.dll 등 11MB) — 미사용
    "Lib/site-packages/adodbapi",         # pywin32 부속 DB 어댑터 — 미사용
    "Lib/site-packages/isapi",            # pywin32 부속 IIS 확장 — 미사용
    "Lib/site-packages/numpy/distutils",  # numpy 빌드 도구 3종 — 런타임 미사용
    "Lib/site-packages/numpy/f2py",
    "Lib/site-packages/numpy/_pyinstaller",
    "Scripts",                            # pip이 만든 콘솔 스텁뿐 — pip 제거로 전부 고아
    "tcl/tix8.4.3",                       # Tix 위젯 확장 — tkinter 기본 위젯만 쓴다
    "include",                            # C 헤더 — 런타임 불필요
)
# 유지하는 것 (지우면 안 되는 이유):
#   Lib/site-packages/pywin32.pth·pywin32_system32·win32/lib — pywin32 DLL 로딩 경로.
#     .pth가 지워진 pythonwin을 참조하지만 site는 없는 경로를 조용히 건너뛴다.
#   Lib/site-packages/win32comext — win32com 내부에서 지연 로딩될 수 있어 보수적으로 유지 (4MB)
#   python/Lib/__pycache__ — 표준 라이브러리 기동 속도
#   Lib/distutils — 3.11 표준 라이브러리 (site-packages 쪽 hack만 제거 대상)

# 설치 대상 — LocalAppData는 OneDrive 동기화 대상이 아니다.
# 회사 PC는 바탕화면·문서가 전부 OneDrive로 리디렉션(KFM)되어 있어서,
# 사용자가 압축을 어디에 풀든 그대로 두면 런타임 전체가 동기화된다.
INSTALL_DIR_NAME = "NOAH_DocGen"  # 경로는 ASCII로 (batch에서 안전)
SHORTCUT_NAME = "NOAH 문서 생성기.lnk"

STEP_TOTAL = 7
step = Stepper(STEP_TOTAL)


# === 빌드 단계 (이 배포판 고유) =============================================

def copy_app() -> None:
    step("앱 파일 복사")
    for name in APP_FILES:
        shutil.copy2(PROJECT_ROOT / name, APP_DIR / name)
    for name in APP_DIRS:
        shutil.copytree(
            PROJECT_ROOT / name, APP_DIR / name,
            ignore=shutil.ignore_patterns("__pycache__", "*.pyc", "Old", *EXCLUDE_FILES),
        )

    leaked = [p.name for p in APP_DIR.rglob("*") if p.name in EXCLUDE_FILES]
    if leaked:
        raise RuntimeError(f"배포에 들어가면 안 되는 파일이 포함됨: {leaked}")

    print(f"  파일 {len(APP_FILES)}개 + 폴더 {len(APP_DIRS)}개")


def create_launchers(version: str) -> None:
    step("런처 · 설치/제거 스크립트")

    # 바탕화면 바로가기 생성/삭제 헬퍼 — 배치의 PowerShell이 아니라 **동봉한 python**으로.
    #
    # WScript.Shell(WshShortcut)은 경로를 시스템 ANSI 코드페이지로 변환한다.
    # 이 회사 PC들은 로캘이 en-US(CP1252)라 한글 경로("바탕 화면", 바로가기 이름)가
    # `?`로 깨져 E_INVALIDARG로 실패한다 (2026-07-30 설치 왕복 테스트 실측 —
    # PowerShell -EncodedCommand로 인자 전달을 고쳐도 COM 내부에서 다시 깨진다).
    # pywin32의 IShellLinkW + SHGetFolderPath는 전 구간 유니코드라 로캘과 무관하고,
    # 바탕화면 경로도 KFM(OneDrive 리디렉션)을 따라간다 (%USERPROFILE%\Desktop 금지).
    write_text(APP_DIR / "_make_shortcut.py", f'''"""바탕화면 바로가기 생성/삭제 — 설치.bat/제거.bat 전용 (사람이 직접 실행할 일 없음)

WScript.Shell은 시스템 ANSI 코드페이지(en-US PC는 CP1252)로 경로를 변환해
한글이 ?로 깨진다. IShellLinkW는 유니코드라 로캘과 무관하다 (빌드 스크립트 주석 참조).

Usage:
    python _make_shortcut.py create | remove
"""
import sys
from pathlib import Path

import pythoncom
from win32com.shell import shell, shellcon

APP_DIR = Path(__file__).resolve().parent
SHORTCUT_NAME = {SHORTCUT_NAME!r}
TARGET = APP_DIR / "NOAH 문서생성기.vbs"


def desktop() -> Path:
    """바탕화면 — KFM(OneDrive 리디렉션) 반영"""
    return Path(shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOPDIRECTORY, None, 0))


def create() -> None:
    link = pythoncom.CoCreateInstance(
        shell.CLSID_ShellLink, None, pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IShellLink)
    link.SetPath(str(TARGET))
    link.SetWorkingDirectory(str(APP_DIR))
    link.QueryInterface(pythoncom.IID_IPersistFile).Save(str(desktop() / SHORTCUT_NAME), 0)


def remove() -> None:
    try:
        (desktop() / SHORTCUT_NAME).unlink()
    except FileNotFoundError:
        pass


if __name__ == "__main__":
    if len(sys.argv) != 2 or sys.argv[1] not in ("create", "remove"):
        print(__doc__)
        sys.exit(2)
    create() if sys.argv[1] == "create" else remove()
''')

    # 평소 실행 — pythonw로 콘솔 없이
    write_text(APP_DIR / "NOAH 문서생성기.vbs", '''Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
appDir = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = appDir
sh.Run """" & appDir & "\\python\\pythonw.exe"" ""noah_gui.py""", 0, False
''')

    # 진단용 — 콘솔에 오류가 보이게
    write_text(APP_DIR / "NOAH 문서생성기.bat", '''@echo off
chcp 65001 >nul
title NOAH 문서 생성기 (진단 모드)
cd /d "%~dp0"
python\\python.exe noah_gui.py
echo.
echo 창이 닫히면 위 메시지를 확인하세요.
pause
''')

    # 제거 — 설치 폴더와 바탕화면 바로가기 삭제.
    #   자기 자신이 삭제 대상 폴더 안에 있으므로 %TEMP%로 복사한 뒤 그쪽을 실행한다.
    #   start(비동기)여야 한다 — call은 부모 cmd(CWD가 설치 폴더 안)를 살려둬 rd가 실패한다.
    #   taskkill은 쓰지 않는다 — 개발 PC의 무관한 python까지 잡는다. 잠겨서 못 지우면 안내만.
    write_text(APP_DIR / "제거.bat", f'''@echo off
chcp 65001 >nul
title NOAH 문서 생성기 제거
set "TARGET=%LOCALAPPDATA%\\{INSTALL_DIR_NAME}"

if /i "%~1"=="GO" goto :run

echo.
echo  NOAH 문서 생성기 제거
echo  ----------------------------------------
echo  다음 폴더를 삭제합니다: %TARGET%
echo  (만든 문서와 NOAH_SO_PO_DN.xlsx는 삭제되지 않습니다)
echo.
choice /c YN /n /m "  제거하시겠습니까? [Y/N] "
if errorlevel 2 exit /b 0

copy /y "%~f0" "%TEMP%\\noah_uninstall.bat" >nul
start "" "%TEMP%\\noah_uninstall.bat" GO
exit

:run
cd /d "%TEMP%"
ping -n 2 127.0.0.1 >nul
rem 바로가기 먼저 — 삭제에 폴더 안의 python이 필요하다 (실패해도 계속)
"%TARGET%\\python\\python.exe" -B "%TARGET%\\_make_shortcut.py" remove >nul 2>&1
rd /s /q "%TARGET%" 2>nul
if exist "%TARGET%" (
    echo  [오류] 일부 파일을 지우지 못했습니다.
    echo  NOAH 문서 생성기가 실행 중이면 닫고 다시 실행하세요.
    pause
    exit /b 1
)
echo.
echo  제거되었습니다.
pause
exit /b 0
''')

    # 설치 — LocalAppData로 복사 + 바탕화면 바로가기
    #   /XF noah_config.ini : 업데이트 시 사용자가 지정한 데이터 파일 경로를 지운다면
    #                          매번 마법사를 다시 띄우게 된다. 설정은 보존한다.
    #   /R:1 /W:1 : 기본값(재시도 100만×30초)은 앱이 켜져 있으면 무한 대기처럼 보인다.
    #               실행 중 잠김은 재시도로 안 풀리므로 빨리 실패하고 안내한다.
    write_text(BUILD_ROOT / "설치.bat", f'''@echo off
chcp 65001 >nul
title NOAH 문서 생성기 설치
set "TARGET=%LOCALAPPDATA%\\{INSTALL_DIR_NAME}"

echo.
echo  NOAH 문서 생성기 설치  (v{version})
echo  ----------------------------------------
echo  설치 위치: %TARGET%
echo.
echo  복사 중입니다. 1~2분 걸립니다...
echo.

robocopy "%~dp0app" "%TARGET%" /MIR /XF noah_config.ini /R:1 /W:1 /NFL /NDL /NJH /NJS /NP >nul
if %ERRORLEVEL% GEQ 8 (
    echo  [오류] 파일 복사에 실패했습니다.
    echo  NOAH 문서 생성기가 실행 중이면 닫고 다시 실행하세요.
    pause
    exit /b 1
)

"%TARGET%\\python\\python.exe" -B "%TARGET%\\_make_shortcut.py" create
if %ERRORLEVEL% NEQ 0 (
    echo  [경고] 바탕화면 바로가기를 만들지 못했습니다.
    echo  설치 폴더의 "NOAH 문서생성기.vbs" 로 직접 실행할 수 있습니다.
)

echo.
echo  설치가 끝났습니다.
echo  바탕화면의 [NOAH 문서 생성기] 아이콘으로 실행하세요.
echo.
pause
''')

    write_text(BUILD_ROOT / "README.txt", f'''NOAH 문서 생성기 — 설치 안내  (v{version})
================================

[설치]
  1. 이 폴더의 "설치.bat" 을 더블클릭하세요.
  2. 바탕화면에 [NOAH 문서 생성기] 아이콘이 생깁니다.
  3. 아이콘을 실행하면 처음 한 번 NOAH_SO_PO_DN.xlsx 위치를 묻습니다.

[준비물]
  - Excel 데스크톱 (문서 생성에 Excel을 사용합니다)
  - NOAH_SO_PO_DN.xlsx 가 OneDrive로 "동기화된" 상태
    ※ 웹 링크(https://...)는 사용할 수 없습니다.
       OneDrive 웹에서 파일을 열고 [내 파일에 바로가기 추가]를 누른 뒤,
       동기화가 끝나면 프로그램에서 [파일 찾기]로 선택하세요.

[만들 수 있는 문서]
  발주서(PO) · 거래명세표(TS) · Proforma Invoice(PI)
  Final Invoice(FI) · Order Confirmation(OC)
  Commercial Invoice(CI) · Packing List(PL)

[납기현황 회신]
  거래처가 "언제 나오냐"고 물을 때 보낼 미출고 현황표를 만듭니다.
  문서 종류에서 [납기현황]을 고르고 사업자등록번호(또는 거래처명 일부)를 넣으세요.
  - 사업자번호를 모르면 [거래처 목록]을 골라 [생성] — 미출고가 있는 거래처가 나옵니다.
  - [메일 초안 만들기]를 켜면 본문에 표가 들어간 회신 메일이 열립니다 (보내기는 직접 확인).

[버전 확인]
  프로그램 창 제목에 버전이 표시됩니다 (예: v{version}).
  문의하실 때 이 버전을 알려주시면 빠르게 확인할 수 있습니다.

[업데이트]
  새 zip을 받아 "설치.bat" 을 다시 실행하면 됩니다.
  지정해 둔 데이터 파일 경로는 그대로 유지됩니다.
  ※ 프로그램이 켜져 있으면 업데이트가 실패합니다. 닫고 실행하세요.

[제거]
  설치 폴더(%LOCALAPPDATA%\\{INSTALL_DIR_NAME})의 "제거.bat" 을 실행하세요.
  프로그램과 설정이 삭제됩니다. 만든 문서와 NOAH_SO_PO_DN.xlsx는 삭제되지 않습니다.

[문제가 생기면]
  설치 폴더(%LOCALAPPDATA%\\{INSTALL_DIR_NAME})의
  "NOAH 문서생성기.bat" 을 실행하면 검은 창에 오류 메시지가 표시됩니다.
  그 내용을 담당자에게 전달해 주세요.
''')
    print("  설치.bat / 제거.bat / README.txt / 런처 2종")


def verify() -> None:
    """배포 전 스모크 — 여기서 걸러야 받는 사람이 안 겪는다"""
    print("\n  [검증] 런타임 import...")
    result = subprocess.run(
        [str(PYTHON_DIR / "python.exe"), "-B", "-c",
         # win32com.shell(win32comext)은 _make_shortcut.py가 쓴다 — 트리밍되면 여기서 걸리게
         "import tkinter, pandas, openpyxl, xlwings, win32com.client, pythoncom;"
         "from win32com.shell import shell, shellcon;"
         "print('  OK  tkinter', tkinter.TkVersion, '| pandas', pandas.__version__,"
         "'| xlwings', xlwings.__version__)"],
        env=verify_env(),
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    print(result.stdout.rstrip() or result.stderr.rstrip())
    if result.returncode != 0:
        raise RuntimeError("런타임 검증 실패 — 위 오류를 확인하세요.")

    # GUI 로드 + "버튼이 부르는 CLI가 실제로 복사됐는지" + BUILD_INFO 존재를 한 번에 본다.
    # APP_FILES에 넣는 걸 잊으면 여기서 걸린다 — 안 걸리면 받는 사람이 그 버튼에서 겪는다.
    # 판정 자체는 noah_gui가 소유한다 (missing_scripts/read_build_info — 테스트와 동일 판정).
    print("  [검증] GUI 문서 종류 ↔ 배포 파일...")
    result = subprocess.run(
        [str(PYTHON_DIR / "python.exe"), "-B", "-c",
         "import sys\n"
         "import noah_gui\n"
         "missing = noah_gui.missing_scripts()\n"
         "if missing:\n"
         "    print('  누락:', missing)\n"
         "    sys.exit(1)\n"
         "if not noah_gui.read_build_info():\n"
         "    print('  누락: BUILD_INFO.txt')\n"
         "    sys.exit(1)\n"
         "print('  OK  문서 종류', len(noah_gui.DOC_TYPES), '종 — CLI·BUILD_INFO 포함')\n"],
        cwd=str(APP_DIR),
        env=verify_env(),
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    print(result.stdout.rstrip() or result.stderr.rstrip())
    if result.returncode != 0:
        raise RuntimeError("GUI 검증 실패 — APP_FILES에 빠진 CLI가 있는지 확인하세요.")

    # CLI 전수 스모크 — 트리밍이 import를 깨지 않았다는 안전망.
    # --help는 argparse가 도움말을 찍고 0으로 끝나므로, 모듈 import 전체가 검증된다.
    clis = [name for name in APP_FILES if name != "noah_gui.py"]
    print(f"  [검증] CLI --help 스모크 ({len(clis)}종, ~30초)...")
    failures: list[str] = []
    for name in clis:
        try:
            result = subprocess.run(
                [str(PYTHON_DIR / "python.exe"), "-B", name, "--help"],
                cwd=str(APP_DIR), env=verify_env(), timeout=120,
                capture_output=True, text=True, encoding="utf-8", errors="replace",
            )
        except subprocess.TimeoutExpired:
            failures.append(f"{name}: 120초 초과 (멈춤)")
            continue
        if result.returncode != 0:
            tail = (result.stderr or result.stdout).strip().splitlines()[-5:]
            failures.append(f"{name} (코드 {result.returncode})\n      " + "\n      ".join(tail))
    if failures:
        for failure in failures:
            print(f"  실패: {failure}")
        raise RuntimeError("CLI 스모크 실패 — 트리밍이 과했거나 의존성이 빠졌습니다.")
    print(f"  OK  {len(clis)}종 전부 실행 가능")


def build() -> int:
    version = compute_version()
    print(f"  빌드 버전: {version}")

    step("빌드 폴더 준비")
    prepare(BUILD_ROOT, APP_DIR)

    step("Python 런타임 (tkinter 포함)")
    install_runtime(RUNTIME_URL, RUNTIME_TAG, BUILD_ROOT, PYTHON_DIR)

    step("패키지 설치 + 버전 대조 (2~3분)")
    packages = install_packages(PYTHON_DIR, HERE / "requirements.txt")

    step("런타임 트리밍")
    trim_runtime(PYTHON_DIR, TRIM_GLOBS, TRIM_DIRS)

    copy_app()
    write_build_info(APP_DIR, version, packages, RUNTIME_PY_VERSION, RUNTIME_TAG)
    create_launchers(version)
    verify()

    step("배포 zip 생성")
    make_zip(ZIP_PATH, BUILD_ROOT, APP_DIR, PYTHON_DIR)

    print(f"\n{'=' * 56}\n  빌드 완료  (v{version})\n{'=' * 56}")
    print("  배포 방법:")
    print(f"    1. {ZIP_PATH.name} 을 전달")
    print("    2. 받는 사람은 압축을 풀고 '설치.bat' 더블클릭")
    print()
    print("  ※ 이 zip에는 회사 로고·서명이 든 템플릿이 들어 있습니다. 사내 전달용입니다.")
    return 0


if __name__ == "__main__":
    sys.exit(build())
