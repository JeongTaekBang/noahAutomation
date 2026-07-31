"""NOAH 대시보드 — 사내 배포판 빌드
=====================================

Python 런타임까지 담은 무설치 배포 zip을 만든다. 받는 사람은 압축을 풀고
`설치.bat`을 누르면 끝이다 (Python 설치 불필요).

Usage:
    python dashboard_dist/build_portable_dashboard.py

출력:
    dashboard_dist/NOAH_대시보드_배포.zip
      ├── 설치.bat        ← 더블클릭 (LocalAppData로 복사 + 바탕화면 바로가기)
      ├── README.txt
      └── app/
          ├── NOAH 대시보드.vbs   ← 평소 실행 (콘솔 없음)
          ├── NOAH 대시보드.bat   ← 문제 진단용 (콘솔 표시)
          ├── 설정.bat            ← noah_config.ini 열기 (DB 경로 지정)
          ├── 제거.bat
          ├── BUILD_INFO.txt
          ├── _run_dashboard.py   ← 빈 포트 찾기 → streamlit 기동 → 브라우저 열기
          ├── dashboard.py, po_generator/, sql/
          └── python/             ← CPython 3.11 + streamlit/plotly/pandas

런타임 내려받기·핀 3자 대조·트리밍·zip 만들기는 문서생성기 배포판과 똑같아서
`build_common.py`(프로젝트 루트)가 갖고 있다. 여기 있는 것은 대시보드만의 것이다.

**dashboard.py를 고치지 않는다.** 예전 빌드(build_dist.py)는 `po_generator` import를
문자열 치환으로 걷어내 standalone 파일을 만들었는데, import가 한 줄 바뀌자 패턴이
안 맞아 빌드가 통째로 죽었고 배포본이 4개월 낡았다. 지금은 `po_generator/`를 그대로
동봉해 dashboard.py를 무수정으로 돌린다 — import가 늘어도 빌드는 안 깨진다.

주의: 이 모듈의 최상위는 **상수와 함수 정의만** 있어야 한다 (tests가 경로로 로드한다).
  예외는 build_common을 찾기 위한 sys.path 한 줄뿐이다.
"""

from __future__ import annotations

import contextlib
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
PROJECT_ROOT = HERE.parent

# 공통부는 프로젝트 루트에 있다. 이 파일은 `python dashboard_dist/...`로 실행되므로
# sys.path[0]이 dashboard_dist/ 다 — 루트를 넣어 줘야 import가 된다.
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

ZIP_PATH = HERE / "NOAH_대시보드_배포.zip"

# 빌드 작업 폴더 — 짧은 경로 + OneDrive 밖.
# 문서생성기와 다른 이름이어야 한다 (두 빌드를 연달아 돌려도 서로 지우지 않게).
BUILD_ROOT = Path(tempfile.gettempdir()) / "noah_dash_build"
APP_DIR = BUILD_ROOT / "app"
PYTHON_DIR = APP_DIR / "python"

# === 런타임 ===
# 문서생성기와 같은 배포본을 쓴다 — 캐시(%TEMP%)를 공유해 두 번째 빌드는 내려받기가 없다.
RUNTIME_TAG = "20260718"
RUNTIME_FILE = "cpython-3.11.15%2B20260718-x86_64-pc-windows-msvc-install_only.tar.gz"
RUNTIME_URL = (
    "https://github.com/astral-sh/python-build-standalone/releases/download/"
    f"{RUNTIME_TAG}/{RUNTIME_FILE}"
)
RUNTIME_PY_VERSION = re.search(r"cpython-([\d.]+)", RUNTIME_FILE).group(1)  # type: ignore[union-attr]

# === 배포에 포함할 앱 파일 ===
APP_FILES = ("dashboard.py",)
# po_generator — dashboard.py가 config.DB_FILE과 db_schema 3개를 쓴다.
#   패키지 __init__을 비워 둔 덕에 이 경로는 표준 라이브러리만 의존한다
#   (openpyxl·xlwings가 딸려오지 않는다 — po_generator/__init__.py 주석 참조).
APP_DIRS = ("po_generator",)

# 배포에서 반드시 빼야 하는 것
#   user_settings.py — 빌드한 사람의 OneDrive 경로가 박혀 있다. 이게 들어가면
#                      받는 사람 PC에서 ini보다 우선해서 남의 경로를 바라본다.
EXCLUDE_FILES = {"user_settings.py"}

INSTALL_DIR_NAME = "NOAH_Dashboard"  # 경로는 ASCII로 (batch에서 안전)
SHORTCUT_NAME = "NOAH 대시보드.lnk"

# === 런타임 트리밍 ===
# 기준: 실측으로 "안 쓰인다"를 확인한 것만 지운다. 지운 뒤 verify()가 증명한다.
#
# pyarrow (설치 직후 84MB) — `import pyarrow`는 lib/ipc/types/util만,
#   streamlit은 여기에 _compute를 더 얹는다 (실측). parquet/flight/dataset/
#   substrait/acero/gandiva는 어느 경로에서도 로드되지 않는다.
#   → arrow.dll · arrow_compute.dll · arrow_python.dll · lib/_compute .pyd 는 남긴다.
# plotly (58MB) — dashboard.py는 st.plotly_chart만 쓴다 (write_html·plotly.offline 0건).
#   package_data의 plotly.min.js는 오프라인 HTML 출력용이고, 브라우저 렌더링은
#   streamlit이 자기 번들로 한다. labextension은 JupyterLab 확장이라 무관.
# pydeck (14MB) — st.map/st.pydeck_chart를 안 쓴다. nbextension은 Jupyter 자산.
# tkinter/tcl — 대시보드는 GUI 창을 만들지 않는다 (안내 대화상자는 ctypes MessageBoxW).
TRIM_GLOBS = (
    "**/*.pdb",
    "Lib/site-packages/distutils-precedence.pth",
    "Lib/site-packages/pip-*.dist-info",
    "Lib/site-packages/setuptools-*.dist-info",
    "Lib/site-packages/PyWin32.chm",
    "Lib/site-packages/numpy/**/tests",
    # pyarrow — 링크 스텁과 미사용 컴포넌트
    "Lib/site-packages/pyarrow/*.lib",
    "Lib/site-packages/pyarrow/arrow_flight*.dll",
    "Lib/site-packages/pyarrow/arrow_substrait*.dll",
    "Lib/site-packages/pyarrow/arrow_dataset*.dll",
    "Lib/site-packages/pyarrow/arrow_acero*.dll",
    "Lib/site-packages/pyarrow/parquet*.dll",
    "Lib/site-packages/pyarrow/gandiva*.dll",
    "Lib/site-packages/pyarrow/_flight*.pyd",
    "Lib/site-packages/pyarrow/_substrait*.pyd",
    "Lib/site-packages/pyarrow/_dataset*.pyd",
    "Lib/site-packages/pyarrow/_acero*.pyd",
    "Lib/site-packages/pyarrow/_parquet*.pyd",
    "Lib/site-packages/pyarrow/_gandiva*.pyd",
    # tkinter 확장 모듈 / Tcl-Tk DLL
    "_tkinter.pyd",
    "DLLs/_tkinter.pyd",
    "tcl86t.dll",
    "tk86t.dll",
)
TRIM_DIRS = (
    "Lib/test",
    "Lib/idlelib",
    "Lib/ensurepip",
    "Lib/tkinter",
    "tcl",
    "Lib/site-packages/pandas/tests",
    "Lib/site-packages/pip",
    "Lib/site-packages/setuptools",
    "Lib/site-packages/_distutils_hack",
    "Lib/site-packages/pythonwin",
    "Lib/site-packages/adodbapi",
    "Lib/site-packages/isapi",
    "Lib/site-packages/numpy/distutils",
    "Lib/site-packages/numpy/f2py",
    "Lib/site-packages/numpy/_pyinstaller",
    "Scripts",
    "include",
    # pyarrow — 테스트·C++ 헤더·소스
    "Lib/site-packages/pyarrow/tests",
    "Lib/site-packages/pyarrow/include",
    "Lib/site-packages/pyarrow/src",
    "Lib/site-packages/pyarrow/parquet",
    # plotly / pydeck — Jupyter 자산과 오프라인 HTML용 번들
    "Lib/site-packages/plotly/labextension",
    "Lib/site-packages/plotly/package_data",
    "Lib/site-packages/pydeck/nbextension",
)
# 유지하는 것:
#   pywin32.pth · pywin32_system32 · win32 · win32com(ext) — 설치.bat의 바로가기가 쓴다.
#   streamlit/static — 프런트엔드 번들 그 자체다. 지우면 빈 화면이 된다.

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
    # sql/ — dashboard.py가 order_book.sql을 읽는다. 폴더째 담되 xlsx는 뺀다
    # (개별 파일을 나열하면 쿼리가 늘 때 조용히 빠진다).
    shutil.copytree(
        PROJECT_ROOT / "sql", APP_DIR / "sql",
        ignore=shutil.ignore_patterns("*.xlsx", "__pycache__"),
    )

    leaked = [p.name for p in APP_DIR.rglob("*") if p.name in EXCLUDE_FILES]
    if leaked:
        raise RuntimeError(f"배포에 들어가면 안 되는 파일이 포함됨: {leaked}")

    sql_count = len(list((APP_DIR / "sql").glob("*.sql")))
    print(f"  dashboard.py + po_generator/ + sql/ ({sql_count}개)")


def create_launchers(version: str) -> None:
    step("런처 · 설치/제거 스크립트")

    # 바탕화면 바로가기 — WScript.Shell은 시스템 ANSI 코드페이지로 경로를 변환해
    # 한글이 ?로 깨진다(회사 PC 로캘 en-US). IShellLinkW는 유니코드라 로캘과 무관하고
    # 바탕화면 경로도 KFM(OneDrive 리디렉션)을 따라간다. 경위는 cli_dist 쪽 주석에 있다.
    write_text(APP_DIR / "_make_shortcut.py", f'''"""바탕화면 바로가기 생성/삭제 — 설치.bat/제거.bat 전용

Usage:
    python _make_shortcut.py create | remove
"""
import sys
from pathlib import Path

import pythoncom
from win32com.shell import shell, shellcon

APP_DIR = Path(__file__).resolve().parent
SHORTCUT_NAME = {SHORTCUT_NAME!r}
TARGET = APP_DIR / "NOAH 대시보드.vbs"


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

    # 대시보드 기동기.
    #
    # 포트를 8501로 박지 않는 이유: 이미 쓰고 있으면 streamlit이 그냥 실패한다.
    #   빈 포트를 OS에게 받아서(bind 0) 넘긴다.
    # 127.0.0.1에만 바인딩하는 이유: 회사 PC는 관리자 권한이 없어 방화벽 인바운드
    #   규칙을 못 만든다(실측: Access is denied). 0.0.0.0에 열면 승인할 수 없는
    #   방화벽 프롬프트만 뜨고 외부 접속은 어차피 안 된다. 내 PC 전용임을 분명히 한다.
    # 대화상자를 ctypes MessageBoxW로 띄우는 이유: tkinter를 트리밍했고(대시보드는
    #   GUI 창이 없다), pythonw로 도는 프로세스는 print가 아무 데도 안 보인다.
    write_text(APP_DIR / "_run_dashboard.py", '''"""NOAH 대시보드 기동 — 빈 포트 확보 → streamlit → 브라우저

콘솔 없이(pythonw) 도는 경로라 사용자에게 보일 말은 MessageBoxW로 띄운다.
"""
from __future__ import annotations

import ctypes
import socket
import subprocess
import sys
import time
import urllib.error
import urllib.request
import webbrowser
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
PYTHONW = APP_DIR / "python" / "pythonw.exe"


def alert(text: str, title: str = "NOAH 대시보드") -> None:
    ctypes.windll.user32.MessageBoxW(0, text, title, 0x10)  # MB_ICONERROR


def free_port() -> int:
    """OS에게 빈 포트를 받아 즉시 반납 — streamlit이 그 번호로 다시 연다"""
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return int(s.getsockname()[1])


def main() -> int:
    sys.path.insert(0, str(APP_DIR))
    from po_generator.config import DB_FILE

    if not Path(DB_FILE).exists():
        alert(
            "데이터 파일(noah_data.db)을 찾지 못했습니다.\\n\\n"
            f"찾은 위치:\\n{DB_FILE}\\n\\n"
            "OneDrive 동기화가 끝났는지 확인하시고, 위치가 다르다면\\n"
            "설치 폴더의 [설정.bat] 으로 경로를 지정하세요."
        )
        return 1

    port = free_port()
    proc = subprocess.Popen(
        [str(PYTHONW), "-m", "streamlit", "run", "dashboard.py",
         "--server.port", str(port),
         "--server.address", "127.0.0.1",
         "--server.headless", "true",
         "--browser.gatherUsageStats", "false"],
        cwd=str(APP_DIR),
    )

    url = f"http://127.0.0.1:{port}"
    for _ in range(120):  # 최대 60초
        if proc.poll() is not None:
            alert("대시보드를 시작하지 못했습니다.\\n\\n"
                  "설치 폴더의 [NOAH 대시보드.bat] 을 실행하면\\n"
                  "검은 창에 오류 내용이 표시됩니다.")
            return 1
        try:
            with urllib.request.urlopen(f"{url}/_stcore/health", timeout=1):
                break
        except (urllib.error.URLError, OSError):
            time.sleep(0.5)
    else:
        proc.terminate()
        alert("대시보드가 60초 안에 응답하지 않았습니다.")
        return 1

    webbrowser.open(url)
    try:
        proc.wait()
    except KeyboardInterrupt:
        proc.terminate()
    return 0


if __name__ == "__main__":
    sys.exit(main())
''')

    # 평소 실행 — pythonw로 콘솔 없이
    write_text(APP_DIR / "NOAH 대시보드.vbs", '''Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
appDir = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = appDir
sh.Run """" & appDir & "\\python\\pythonw.exe"" ""_run_dashboard.py""", 0, False
''')

    # 진단용 — 콘솔에 오류가 보이게
    write_text(APP_DIR / "NOAH 대시보드.bat", '''@echo off
chcp 65001 >nul
title NOAH 대시보드 (진단 모드)
cd /d "%~dp0"
python\\python.exe _run_dashboard.py
echo.
echo 창이 닫히면 위 메시지를 확인하세요.
pause
''')

    # DB 경로 지정 — 없으면 만들어서 연다 (빈 메모장이 뜨는 것보다 낫다)
    write_text(APP_DIR / "설정.bat", '''@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist noah_config.ini (
    echo [paths]> noah_config.ini
    echo ; noah_data.db 가 들어 있는 폴더 경로를 적으세요.>> noah_config.ini
    echo ; 예: data_folder = C:\\Users\\홍길동\\OneDrive - Rotork plc\\...\\NOAH ACTUATION>> noah_config.ini
    echo data_folder = >> noah_config.ini
)
notepad noah_config.ini
''')

    write_text(APP_DIR / "제거.bat", f'''@echo off
chcp 65001 >nul
title NOAH 대시보드 제거
set "TARGET=%LOCALAPPDATA%\\{INSTALL_DIR_NAME}"

if /i "%~1"=="GO" goto :run

echo.
echo  NOAH 대시보드 제거
echo  ----------------------------------------
echo  다음 폴더를 삭제합니다: %TARGET%
echo  (noah_data.db 등 데이터 파일은 삭제되지 않습니다)
echo.
choice /c YN /n /m "  제거하시겠습니까? [Y/N] "
if errorlevel 2 exit /b 0

copy /y "%~f0" "%TEMP%\\noah_dash_uninstall.bat" >nul
start "" "%TEMP%\\noah_dash_uninstall.bat" GO
exit

:run
cd /d "%TEMP%"
ping -n 2 127.0.0.1 >nul
rem 바로가기 먼저 — 삭제에 폴더 안의 python이 필요하다 (실패해도 계속)
"%TARGET%\\python\\python.exe" -B "%TARGET%\\_make_shortcut.py" remove >nul 2>&1
rd /s /q "%TARGET%" 2>nul
if exist "%TARGET%" (
    echo  [오류] 일부 파일을 지우지 못했습니다.
    echo  대시보드가 실행 중이면 브라우저 탭과 창을 닫고 다시 실행하세요.
    pause
    exit /b 1
)
echo.
echo  제거되었습니다.
pause
exit /b 0
''')

    # 설치 — LocalAppData로 복사 + 바탕화면 바로가기
    #   /XF noah_config.ini : 업데이트 때 사용자가 지정한 DB 경로를 보존한다.
    write_text(BUILD_ROOT / "설치.bat", f'''@echo off
chcp 65001 >nul
title NOAH 대시보드 설치
set "TARGET=%LOCALAPPDATA%\\{INSTALL_DIR_NAME}"

echo.
echo  NOAH 대시보드 설치  (v{version})
echo  ----------------------------------------
echo  설치 위치: %TARGET%
echo.
echo  복사 중입니다. 1~2분 걸립니다...
echo.

robocopy "%~dp0app" "%TARGET%" /MIR /XF noah_config.ini /R:1 /W:1 /NFL /NDL /NJH /NJS /NP >nul
if %ERRORLEVEL% GEQ 8 (
    echo  [오류] 파일 복사에 실패했습니다.
    echo  대시보드가 실행 중이면 닫고 다시 실행하세요.
    pause
    exit /b 1
)

"%TARGET%\\python\\python.exe" -B "%TARGET%\\_make_shortcut.py" create
if %ERRORLEVEL% NEQ 0 (
    echo  [경고] 바탕화면 바로가기를 만들지 못했습니다.
    echo  설치 폴더의 "NOAH 대시보드.vbs" 로 직접 실행할 수 있습니다.
)

echo.
echo  설치가 끝났습니다.
echo  바탕화면의 [NOAH 대시보드] 아이콘으로 실행하세요.
echo.
pause
''')

    write_text(BUILD_ROOT / "README.txt", f'''NOAH 대시보드 — 설치 안내  (v{version})
================================

[설치]
  1. 이 폴더의 "설치.bat" 을 더블클릭하세요.
  2. 바탕화면에 [NOAH 대시보드] 아이콘이 생깁니다.
  3. 아이콘을 실행하면 기본 브라우저에 대시보드가 열립니다.

[준비물]
  noah_data.db 가 OneDrive로 "동기화된" 상태여야 합니다.
  담당자가 공유한 NOAH ACTUATION 폴더를 [내 파일에 바로가기 추가]로 받아
  동기화가 끝나면 준비 완료입니다.

  ※ 파일을 못 찾는다는 안내가 뜨면 설치 폴더의 "설정.bat" 을 실행해
     noah_data.db 가 들어 있는 폴더 경로를 적어 주세요.

[데이터는 언제 갱신되나요]
  담당자가 갱신하면 OneDrive가 각자 PC로 내려보냅니다.
  대시보드는 켤 때마다 그 시점의 파일을 읽고, 화면 위쪽에
  "데이터 기준" 시각을 표시합니다. 그 시각이 오늘이 아니면
  아직 동기화 중이거나 담당자가 아직 갱신하지 않은 것입니다.

[종료]
  브라우저 탭을 닫아도 서버는 남습니다.
  작업 표시줄에서 완전히 끝내려면 로그아웃하거나 PC를 재시작하세요.
  (설치 폴더의 "NOAH 대시보드.bat" 으로 켰다면 그 검은 창을 닫으면 됩니다.)

[버전 확인]
  설치 폴더의 BUILD_INFO.txt 첫 줄에 버전이 있습니다 (v{version}).
  문의하실 때 이 버전을 알려주시면 빠르게 확인할 수 있습니다.

[업데이트]
  새 zip을 받아 "설치.bat" 을 다시 실행하면 됩니다.
  지정해 둔 DB 경로(noah_config.ini)는 그대로 유지됩니다.

[제거]
  설치 폴더(%LOCALAPPDATA%\\{INSTALL_DIR_NAME})의 "제거.bat" 을 실행하세요.

[문제가 생기면]
  설치 폴더의 "NOAH 대시보드.bat" 을 실행하면 검은 창에 오류 메시지가
  표시됩니다. 그 내용을 담당자에게 전달해 주세요.
''')
    print("  설치.bat / 제거.bat / 설정.bat / README.txt / 런처 2종 / 기동기")


def verify() -> None:
    """배포 전 스모크 — 여기서 걸러야 받는 사람이 안 겪는다"""
    print("\n  [검증] 런타임 import + pyarrow 데이터 경로...")
    # pyarrow를 크게 잘라냈다. streamlit이 표를 그릴 때 쓰는 경로
    # (pandas → Arrow Table → IPC 직렬화, + _compute)를 실제로 태워서 증명한다.
    # 단순 import만으로는 잘라낸 컴포넌트가 지연 로딩될 때 터지는 것을 못 잡는다.
    result = subprocess.run(
        [str(PYTHON_DIR / "python.exe"), "-B", "-c",
         "import streamlit, plotly, pandas as pd, pyarrow as pa, pyarrow.compute as pc\n"
         "from po_generator.config import DB_FILE\n"
         "from po_generator.db_schema import get_sync_metadata, SYNC_LOG_CHANGE_TYPES\n"
         "df = pd.DataFrame({'a': [1, 2, 3], 'b': ['x', 'y', 'z']})\n"
         "t = pa.Table.from_pandas(df)\n"
         "assert pc.sum(t['a']).as_py() == 6\n"
         "sink = pa.BufferOutputStream()\n"
         "with pa.ipc.new_stream(sink, t.schema) as w:\n"
         "    w.write_table(t)\n"
         "assert len(sink.getvalue()) > 0\n"
         "print('  OK  streamlit', streamlit.__version__, '| plotly', plotly.__version__,\n"
         "      '| pandas', pd.__version__, '| pyarrow', pa.__version__, 'IPC 왕복')\n"],
        cwd=str(APP_DIR),
        env=verify_env(),
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    print(result.stdout.rstrip() or result.stderr.rstrip())
    if result.returncode != 0:
        raise RuntimeError("런타임 검증 실패 — 트리밍이 과했는지 확인하세요.")

    # 진짜 검증: streamlit을 실제로 띄워 dashboard.py를 **실제 데이터로** 실행시킨다.
    # import만으로는 "서버가 뜨는가 / 쿼리가 도는가"를 못 본다.
    print("  [검증] streamlit 기동 + dashboard.py 실행 (~40초)...")
    with _verify_data_config() as note:
        if note:
            print(f"  {note}")
        _verify_serves()


@contextlib.contextmanager
def _verify_data_config():
    """검증 동안만 APP_DIR에 noah_config.ini를 두고 실제 DB **사본**을 가리킨다

    사본을 쓰는 이유: dashboard.py는 열 때 `ensure_so_change_ack_table()`로 DDL을
    실행한다. 빌드가 운영 DB에 쓰기를 하는 일은 없어야 한다.
    ini를 지우는 이유: 빌드한 사람의 경로가 배포 zip에 실려 나가면, 받는 사람이
    남의 PC 경로를 바라본다 (user_settings.py를 EXCLUDE하는 것과 같은 이유).
    """
    ini = APP_DIR / "noah_config.ini"
    data_dir = BUILD_ROOT / "_verify_data"
    note = ""
    try:
        from po_generator.config import DB_FILE as real_db
        if Path(real_db).exists():
            data_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(real_db, data_dir / "noah_data.db")
            write_text(ini, f"[paths]\ndata_folder = {data_dir}\n")
            note = f"실제 DB 사본으로 검증 ({Path(real_db).stat().st_size / 1024 / 1024:.0f} MB)"
        else:
            note = "[주의] noah_data.db가 없어 빈 DB로 검증합니다 — 쿼리 경로는 미검증"
        yield note
    finally:
        ini.unlink(missing_ok=True)
        shutil.rmtree(data_dir, ignore_errors=True)


def _verify_serves() -> None:
    """streamlit을 띄워 health와 스크립트 실행을 확인하고 반드시 종료시킨다"""
    import socket
    import time
    import urllib.error
    import urllib.request

    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        port = int(s.getsockname()[1])

    proc = subprocess.Popen(
        [str(PYTHON_DIR / "python.exe"), "-B", "-m", "streamlit", "run", "dashboard.py",
         "--server.port", str(port), "--server.address", "127.0.0.1",
         "--server.headless", "true", "--browser.gatherUsageStats", "false"],
        cwd=str(APP_DIR), env=verify_env(),
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        text=True, encoding="utf-8", errors="replace",
    )
    try:
        healthy = False
        for _ in range(80):  # 최대 40초
            if proc.poll() is not None:
                break
            try:
                with urllib.request.urlopen(f"http://127.0.0.1:{port}/_stcore/health",
                                            timeout=1) as resp:
                    if resp.status == 200:
                        healthy = True
                        break
            except (urllib.error.URLError, OSError):
                time.sleep(0.5)

        if not healthy:
            raise RuntimeError("streamlit 서버가 응답하지 않습니다 — 위 로그를 확인하세요.")

        # health는 스크립트 오류와 무관하게 200이다. dashboard.py가 실제로 돌았는지는
        # 페이지를 한 번 받아 보고, 서버 로그에 Traceback이 찍혔는지로 판정한다.
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/", timeout=10) as resp:
            if resp.status != 200:
                raise RuntimeError(f"대시보드 페이지 응답 {resp.status}")
        time.sleep(3)  # 스크립트 실행이 로그에 남을 시간
        print(f"  OK  포트 {port} health 200 + 페이지 200")
    finally:
        proc.terminate()
        try:
            out = proc.communicate(timeout=15)[0] or ""
        except subprocess.TimeoutExpired:
            proc.kill()
            out = proc.communicate()[0] or ""

    if "Traceback" in out:
        tail = "\n      ".join(out.strip().splitlines()[-15:])
        raise RuntimeError(f"dashboard.py 실행 중 예외가 발생했습니다:\n      {tail}")


def build() -> int:
    version = compute_version()
    print(f"  빌드 버전: {version}")

    step("빌드 폴더 준비")
    prepare(BUILD_ROOT, APP_DIR)

    step("Python 런타임")
    install_runtime(RUNTIME_URL, RUNTIME_TAG, BUILD_ROOT, PYTHON_DIR)

    step("패키지 설치 + 버전 대조 (3~5분)")
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
    print("    3. noah_data.db 가 들어 있는 OneDrive 폴더를 공유해 둘 것")
    print()
    return 0


if __name__ == "__main__":
    sys.exit(build())
