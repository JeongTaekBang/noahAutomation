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
      └── app/            ← 본체
          ├── NOAH 문서생성기.vbs   ← 평소 실행 (콘솔 없음)
          ├── NOAH 문서생성기.bat   ← 문제 진단용 (콘솔 표시)
          ├── noah_gui.py, create_*.py, po_generator/, templates/
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
"""

from __future__ import annotations

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

# === 배포에 포함할 앱 파일 ===
APP_FILES = (
    "noah_gui.py",
    "create_po.py",
    "create_ts.py",
    "create_pi.py",
    "create_fi.py",
    "create_oc.py",
    "create_ci.py",
    "create_pl.py",
)
APP_DIRS = ("po_generator", "templates")

# 배포에서 반드시 빼야 하는 것
#   user_settings.py — 빌드한 사람의 OneDrive 경로가 박혀 있다. 이게 들어가면
#                      받는 사람 PC에서 ini보다 우선해서 남의 경로를 바라본다.
EXCLUDE_FILES = {"user_settings.py"}

# 런타임 트리밍 (배포 크기 302MB → 190MB)
TRIM_GLOBS = ("**/*.pdb",)
TRIM_DIRS = (
    "Lib/test",
    "Lib/idlelib",
    "Lib/site-packages/pandas/tests",
)

# 설치 대상 — LocalAppData는 OneDrive 동기화 대상이 아니다.
# 회사 PC는 바탕화면·문서가 전부 OneDrive로 리디렉션(KFM)되어 있어서,
# 사용자가 압축을 어디에 풀든 그대로 두면 런타임 전체가 동기화된다.
INSTALL_DIR_NAME = "NOAH_DocGen"  # 경로는 ASCII로 (batch에서 안전)
SHORTCUT_NAME = "NOAH 문서 생성기.lnk"


def step(msg: str) -> None:
    print(f"\n{'=' * 56}\n  {msg}\n{'=' * 56}")


def write_text(path: Path, content: str) -> None:
    """UTF-8 (BOM 없음) — chcp 65001과 조합해 기존 create_po.bat과 동일한 방식"""
    path.write_text(content, encoding="utf-8", newline="\r\n")


def download(url: str, dest: Path) -> None:
    print(f"  다운로드: {url.rsplit('/', 1)[-1]}")
    req = urllib.request.Request(url, headers={"User-Agent": "noah-build"})
    with urllib.request.urlopen(req, timeout=300) as resp, dest.open("wb") as f:
        shutil.copyfileobj(resp, f)
    print(f"  {dest.stat().st_size / 1024 / 1024:.1f} MB")


# === 빌드 단계 =============================================================

def prepare() -> None:
    step("1/7  빌드 폴더 준비")
    if BUILD_ROOT.exists():
        shutil.rmtree(BUILD_ROOT)
    APP_DIR.mkdir(parents=True)
    print(f"  {BUILD_ROOT}")


def install_runtime() -> None:
    step("2/7  Python 런타임 (tkinter 포함)")
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


def install_packages() -> None:
    step("3/7  패키지 설치 (2~3분)")
    subprocess.run(
        [str(PYTHON_DIR / "python.exe"), "-m", "pip", "install",
         "-r", str(HERE / "requirements.txt"),
         "--no-warn-script-location", "-q"],
        check=True,
    )
    print("  완료")


def trim_runtime() -> None:
    step("4/7  런타임 트리밍")
    removed = 0
    for pattern in TRIM_GLOBS:
        for path in PYTHON_DIR.glob(pattern):
            if path.is_file():
                removed += path.stat().st_size
                path.unlink()
    for rel in TRIM_DIRS:
        target = PYTHON_DIR / rel
        if target.exists():
            removed += sum(f.stat().st_size for f in target.rglob("*") if f.is_file())
            shutil.rmtree(target)
    print(f"  {removed / 1024 / 1024:.0f} MB 제거")


def copy_app() -> None:
    step("5/7  앱 파일 복사")
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


def create_launchers() -> None:
    step("6/7  런처 · 설치 스크립트")

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

    # 설치 — LocalAppData로 복사 + 바탕화면 바로가기
    #   /XF noah_config.ini : 업데이트 시 사용자가 지정한 데이터 파일 경로를 지운다면
    #                          매번 마법사를 다시 띄우게 된다. 설정은 보존한다.
    write_text(BUILD_ROOT / "설치.bat", f'''@echo off
chcp 65001 >nul
title NOAH 문서 생성기 설치
set "TARGET=%LOCALAPPDATA%\\{INSTALL_DIR_NAME}"

echo.
echo  NOAH 문서 생성기 설치
echo  ----------------------------------------
echo  설치 위치: %TARGET%
echo.
echo  복사 중입니다. 1~2분 걸립니다...
echo.

robocopy "%~dp0app" "%TARGET%" /MIR /XF noah_config.ini /NFL /NDL /NJH /NJS /NP >nul
if %ERRORLEVEL% GEQ 8 (
    echo  [오류] 파일 복사에 실패했습니다.
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$d=[Environment]::GetFolderPath('Desktop');" ^
  "$t=Join-Path $env:LOCALAPPDATA '{INSTALL_DIR_NAME}';" ^
  "$s=(New-Object -ComObject WScript.Shell).CreateShortcut((Join-Path $d '{SHORTCUT_NAME}'));" ^
  "$s.TargetPath=Join-Path $t 'NOAH 문서생성기.vbs';" ^
  "$s.WorkingDirectory=$t; $s.Save()"

echo.
echo  설치가 끝났습니다.
echo  바탕화면의 [NOAH 문서 생성기] 아이콘으로 실행하세요.
echo.
pause
''')

    write_text(BUILD_ROOT / "README.txt", f'''NOAH 문서 생성기 — 설치 안내
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

[업데이트]
  새 zip을 받아 "설치.bat" 을 다시 실행하면 됩니다.
  지정해 둔 데이터 파일 경로는 그대로 유지됩니다.

[문제가 생기면]
  설치 폴더(%LOCALAPPDATA%\\{INSTALL_DIR_NAME})의
  "NOAH 문서생성기.bat" 을 실행하면 검은 창에 오류 메시지가 표시됩니다.
  그 내용을 담당자에게 전달해 주세요.
''')
    print("  설치.bat / README.txt / 런처 2종")


def verify() -> None:
    """배포 전 스모크 — 여기서 걸러야 받는 사람이 안 겪는다"""
    print("\n  [검증] 런타임 import...")
    result = subprocess.run(
        [str(PYTHON_DIR / "python.exe"), "-B", "-c",
         "import tkinter, pandas, openpyxl, xlwings, win32com.client, pythoncom;"
         "print('  OK  tkinter', tkinter.TkVersion, '| pandas', pandas.__version__,"
         "'| xlwings', xlwings.__version__)"],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    print(result.stdout.rstrip() or result.stderr.rstrip())
    if result.returncode != 0:
        raise RuntimeError("런타임 검증 실패 — 위 오류를 확인하세요.")

    print("  [검증] GUI 모듈 로드...")
    result = subprocess.run(
        [str(PYTHON_DIR / "python.exe"), "-B", "-c",
         "import noah_gui; print('  OK  문서 종류', len(noah_gui.DOC_TYPES), '종')"],
        cwd=str(APP_DIR),
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    print(result.stdout.rstrip() or result.stderr.rstrip())
    if result.returncode != 0:
        raise RuntimeError("GUI 모듈 검증 실패")


def make_zip() -> None:
    step("7/7  배포 zip 생성")
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
    prepare()
    install_runtime()
    install_packages()
    trim_runtime()
    copy_app()
    create_launchers()
    verify()
    make_zip()

    print(f"\n{'=' * 56}\n  빌드 완료\n{'=' * 56}")
    print("  배포 방법:")
    print(f"    1. {ZIP_PATH.name} 을 전달")
    print("    2. 받는 사람은 압축을 풀고 '설치.bat' 더블클릭")
    print()
    print("  ※ 이 zip에는 회사 로고·서명이 든 템플릿이 들어 있습니다. 사내 전달용입니다.")
    return 0


if __name__ == "__main__":
    sys.exit(build())
