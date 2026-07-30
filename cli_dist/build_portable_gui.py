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

주의: 이 모듈의 최상위는 **상수와 함수 정의만** 있어야 한다.
  tests/test_noah_gui.py가 이 파일을 경로로 로드해 APP_FILES·parse_pins 등을
  검사한다 — 모듈 로드가 곧 빌드 시작이면 테스트가 빌드를 돌려버린다.
"""

from __future__ import annotations

import datetime as dt
import importlib.metadata
import itertools
import json
import os
import re
import shutil
import subprocess
import sys
import tarfile
import tempfile
import urllib.request
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
PROJECT_ROOT = HERE.parent

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

# 런타임에 pip이 필요 없어 트리밍으로 제거하지만, 핀 대조 시점(설치 직후)에는
# 아직 존재한다 — "핀에 없는 패키지" 검사에서만 눈감아 준다.
IGNORE_PACKAGES = frozenset({"pip", "setuptools"})

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
_STEP_NO = itertools.count(1)


def step(msg: str) -> None:
    print(f"\n{'=' * 56}\n  {next(_STEP_NO)}/{STEP_TOTAL}  {msg}\n{'=' * 56}")


def write_text(path: Path, content: str) -> None:
    """UTF-8 (BOM 없음) — chcp 65001과 조합해 기존 create_po.bat과 동일한 방식"""
    path.write_text(content, encoding="utf-8", newline="\r\n")


def download(url: str, dest: Path) -> None:
    print(f"  다운로드: {url.rsplit('/', 1)[-1]}")
    req = urllib.request.Request(url, headers={"User-Agent": "noah-build"})
    with urllib.request.urlopen(req, timeout=300) as resp, dest.open("wb") as f:
        shutil.copyfileobj(resp, f)
    print(f"  {dest.stat().st_size / 1024 / 1024:.1f} MB")


# === 버전 / 핀 대조 헬퍼 ====================================================

def canon(name: str) -> str:
    """패키지명 정규화 (PEP 503) — et_xmlfile/et-xmlfile 같은 표기 차이를 흡수한다"""
    return re.sub(r"[-_.]+", "-", name).lower()


def parse_pins(text: str) -> dict[str, str]:
    """requirements 텍스트 → {정규화된 이름: 버전}

    `이름==버전` 형식만 허용한다. `>=` 같은 범위가 끼어들면 "여기서 테스트한
    그대로"라는 재현성 약속이 조용히 깨지므로 파싱 단계에서 막는다.
    """
    pins: dict[str, str] = {}
    for raw in text.splitlines():
        line = raw.split("#", 1)[0].strip()
        if not line:
            continue
        m = re.fullmatch(r"([A-Za-z0-9._-]+)\s*==\s*([A-Za-z0-9.+!_-]+)", line)
        if m is None:
            raise ValueError(f"핀 고정(이름==버전)이 아닌 요구사항: {line!r}")
        pins[canon(m.group(1))] = m.group(2)
    return pins


def compute_version() -> str:
    """빌드 버전 — 날짜 + git 커밋 (예: 2026.07.30+326b87e)

    CHANGELOG이 날짜 기반이라 버전도 날짜를 앞세운다. 커밋 해시가 있어야
    "어느 빌드 쓰세요?" 문의에서 코드 상태를 특정할 수 있다.
    커밋 안 된 변경이 있으면 `.dirty` — 깨끗한 트리에서 빌드하라는 신호다.
    """
    date = dt.date.today().strftime("%Y.%m.%d")

    def _git(*args: str) -> str:
        return subprocess.run(
            ["git", *args],
            cwd=str(PROJECT_ROOT), capture_output=True, text=True,
            errors="replace", check=True,
        ).stdout.strip()

    try:
        sha = _git("rev-parse", "--short", "HEAD")
        dirty = ".dirty" if _git("status", "--porcelain") else ""
    except (OSError, subprocess.CalledProcessError):
        return f"{date}+unknown"
    return f"{date}+{sha}{dirty}"


def verify_env() -> dict[str, str]:
    """검증용 자식 프로세스 환경

    파이프로 받는 자식 stdout은 기본이 콘솔 코드페이지(한국어 Windows면 cp949)라
    아래 `encoding="utf-8"` 디코딩과 어긋나 한글 메시지가 깨진다.
    통과/실패를 읽을 수 없으면 검증이 검증이 아니다.
    """
    return {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUTF8": "1"}


def _runtime_packages() -> dict[str, str]:
    """배포 런타임에 실제 설치된 패키지 {정규화된 이름: 버전}"""
    result = subprocess.run(
        [str(PYTHON_DIR / "python.exe"), "-B", "-c",
         "import importlib.metadata, json\n"
         "print(json.dumps({d.metadata['Name']: d.version"
         " for d in importlib.metadata.distributions() if d.metadata['Name']}))\n"],
        env=verify_env(),
        capture_output=True, text=True, encoding="utf-8", errors="replace", check=True,
    )
    return {canon(name): ver for name, ver in json.loads(result.stdout).items()}


def _dev_env_versions(names: set[str]) -> dict[str, str]:
    """빌드를 실행한 개발 env(po-automate)의 설치 버전 — 없는 패키지는 빠진다"""
    versions: dict[str, str] = {}
    for name in names:
        try:
            versions[name] = importlib.metadata.version(name)
        except importlib.metadata.PackageNotFoundError:
            continue
    return versions


# === 빌드 단계 =============================================================

def prepare() -> None:
    step("빌드 폴더 준비")
    if BUILD_ROOT.exists():
        shutil.rmtree(BUILD_ROOT)
    APP_DIR.mkdir(parents=True)
    print(f"  {BUILD_ROOT}")


def install_runtime() -> None:
    step("Python 런타임 (tkinter 포함)")
    cache = Path(tempfile.gettempdir()) / f"noah_runtime_{RUNTIME_TAG}.tar.gz"
    if not cache.exists():
        download(RUNTIME_URL, cache)
    else:
        print(f"  캐시 사용: {cache.name}")

    with tarfile.open(cache) as tar:
        tar.extractall(BUILD_ROOT)  # noqa: S202 — 신뢰된 공식 배포본
    # 아카이브 최상위가 python/ 이므로 app/ 아래로 옮긴다
    shutil.move(str(BUILD_ROOT / "python"), str(PYTHON_DIR))
    print(f"  {PYTHON_DIR}")


def install_packages() -> dict[str, str]:
    """패키지 설치 + 3자 대조 (핀 = 배포 런타임 = 개발 env)

    transitive까지 requirements.txt에 고정돼 있고, 여기서 어긋남을 전부 모아
    실패시킨다. 통과하면 배포판은 개발 PC에서 테스트한 버전 조합 그대로다.

    Returns:
        배포 런타임의 {패키지: 버전} — BUILD_INFO.txt에 재사용
    """
    step("패키지 설치 + 버전 대조 (2~3분)")
    subprocess.run(
        [str(PYTHON_DIR / "python.exe"), "-m", "pip", "install",
         "-r", str(HERE / "requirements.txt"),
         "--no-warn-script-location", "-q"],
        check=True,
    )

    pins = parse_pins((HERE / "requirements.txt").read_text(encoding="utf-8"))
    installed = _runtime_packages()
    dev = _dev_env_versions(set(pins))

    problems: list[str] = []
    # 핀 3자 대조 — 같은 검사를 배포 런타임/개발 env 두 대상에 적용
    for label, actual in (("배포 런타임", installed), ("개발 env", dev)):
        for name, want in sorted(pins.items()):
            got = actual.get(name)
            if got != want:
                problems.append(f"{label} {name}: 핀 {want} vs 설치 {got or '없음'}")
    for name in sorted(set(installed) - set(pins) - IGNORE_PACKAGES):
        problems.append(
            f"핀에 없는 패키지 설치됨: {name} {installed[name]}"
            " — requirements.txt에 핀을 추가하세요 (의존성이 늘었다는 신호)"
        )

    if problems:
        detail = "\n".join(f"  - {p}" for p in problems)
        raise RuntimeError(
            "패키지 버전이 핀과 다릅니다 — 배포판은 개발 PC에서 테스트한 그대로여야 합니다.\n"
            f"{detail}\n"
            "  해결: 개발 env(po-automate)에서 `pip install 이름==버전`으로 맞추거나,\n"
            "        새 버전으로 테스트를 통과시킨 뒤 cli_dist/requirements.txt 핀을 갱신하세요."
        )
    print(f"  OK  핀 {len(pins)}개 = 배포 런타임 = 개발 env")
    return installed


def trim_runtime() -> None:
    step("런타임 트리밍")
    removed = 0

    def _rm(path: Path) -> None:
        """파일/디렉터리 제거 (크기 집계). 이미 없는 경로는 no-op —
        앞선 glob 패턴이 부모째 지웠을 수 있어 존재 확인을 여기서 흡수한다."""
        nonlocal removed
        if path.is_file():
            removed += path.stat().st_size
            path.unlink()
        elif path.is_dir():
            removed += sum(f.stat().st_size for f in path.rglob("*") if f.is_file())
            shutil.rmtree(path)

    for pattern in TRIM_GLOBS:
        for path in PYTHON_DIR.glob(pattern):
            _rm(path)
    for rel in TRIM_DIRS:
        _rm(PYTHON_DIR / rel)
    print(f"  {removed / 1024 / 1024:.0f} MB 제거")


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


def write_build_info(version: str, packages: dict[str, str]) -> None:
    """app/BUILD_INFO.txt — 배포판의 정체

    기계가 읽는 건 `version:` 줄 하나뿐이다 (noah_gui.read_build_info → GUI 타이틀).
    나머지는 문의 대응용 — "어느 빌드 쓰세요?"에 사용자가 이 파일을 열어 답한다.
    """
    lines = [
        f"version: {version}",
        f"built: {dt.datetime.now():%Y-%m-%d %H:%M}",
        f"runtime: python {RUNTIME_PY_VERSION} (python-build-standalone {RUNTIME_TAG})",
        "packages:",
    ]
    lines += [
        f"  {name} {ver}"
        for name, ver in sorted(packages.items())
        if name not in IGNORE_PACKAGES  # 트리밍으로 제거되는 것은 싣지 않는다
    ]
    write_text(APP_DIR / "BUILD_INFO.txt", "\n".join(lines) + "\n")
    print(f"  BUILD_INFO.txt  v{version}")


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


def make_zip() -> None:
    step("배포 zip 생성")
    if ZIP_PATH.exists():
        ZIP_PATH.unlink()
    ZIP_PATH.parent.mkdir(parents=True, exist_ok=True)

    # 앱 쪽 __pycache__는 넣지 않는다 (검증 단계가 만들 수 있고, 첫 실행 때 다시 생긴다).
    # python/ 안의 것은 표준 라이브러리 로딩을 빠르게 하므로 그대로 둔다.
    for cache in (APP_DIR.rglob("__pycache__")):
        if PYTHON_DIR not in cache.parents:
            shutil.rmtree(cache, ignore_errors=True)

    files = [p for p in BUILD_ROOT.rglob("*") if p.is_file()]
    with zipfile.ZipFile(ZIP_PATH, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for i, path in enumerate(files, 1):
            zf.write(path, path.relative_to(BUILD_ROOT))
            if i % 2000 == 0:
                print(f"  {i}/{len(files)}...")

    raw = sum(p.stat().st_size for p in files)
    print(f"\n  {ZIP_PATH}")
    print(f"  {ZIP_PATH.stat().st_size / 1024 / 1024:.0f} MB "
          f"(압축 전 {raw / 1024 / 1024:.0f} MB, 파일 {len(files):,}개)")


def build() -> int:
    version = compute_version()
    print(f"  빌드 버전: {version}")

    prepare()
    install_runtime()
    packages = install_packages()
    trim_runtime()
    copy_app()
    write_build_info(version, packages)
    create_launchers(version)
    verify()
    make_zip()

    print(f"\n{'=' * 56}\n  빌드 완료  (v{version})\n{'=' * 56}")
    print("  배포 방법:")
    print(f"    1. {ZIP_PATH.name} 을 전달")
    print("    2. 받는 사람은 압축을 풀고 '설치.bat' 더블클릭")
    print()
    print("  ※ 이 zip에는 회사 로고·서명이 든 템플릿이 들어 있습니다. 사내 전달용입니다.")
    return 0


if __name__ == "__main__":
    sys.exit(build())
