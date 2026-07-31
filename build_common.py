"""사내 배포판 빌드 공통부
==========================

`cli_dist/build_portable_gui.py`(문서 생성기)와 `dashboard_dist/build_portable_dashboard.py`
(대시보드)가 공유한다. 두 배포판은 **런타임을 담는 방식이 같고 담는 내용만 다르다** —
런타임 내려받기·핀 3자 대조·트리밍·zip 만들기는 한 글자도 다를 이유가 없다.

복사해 두지 않는 이유는 `po_generator/mail_cli.py`와 같다. 복사본은 갈라지고, 그 갈라짐은
"한쪽 배포판만 조용히 낡은 pandas로 나간다" 같은 형태로 늦게 드러난다. 여기서 한 번 고치면
두 배포판이 같이 고쳐진다.

각 배포판이 따로 소유하는 것 (여기 두지 않는다):
  - 무엇을 담는가       APP_FILES / APP_DIRS
  - 어떻게 실행하는가   런처·설치/제거 스크립트
  - 무엇을 증명하는가   verify() — 앱마다 "돌아간다"의 정의가 다르다

주의: 이 모듈의 최상위는 **상수와 함수 정의만** 있어야 한다.
  tests가 빌드 스크립트를 경로로 로드하는데, 그때 이 모듈도 함께 임포트된다.
  최상위에서 뭔가 실행하면 테스트가 빌드를 돌려버린다.
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
import tarfile
import tempfile
import urllib.request
import zipfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent

# 런타임에 pip이 필요 없어 트리밍으로 제거하지만, 핀 대조 시점(설치 직후)에는
# 아직 존재한다 — "핀에 없는 패키지" 검사에서만 눈감아 준다.
IGNORE_PACKAGES = frozenset({"pip", "setuptools"})


class Stepper:
    """`1/7  런타임 설치` 같은 단계 헤더 — 배포판마다 단계 수가 달라 인스턴스로 만든다"""

    def __init__(self, total: int) -> None:
        self.total = total
        self._counter = itertools.count(1)

    def __call__(self, msg: str) -> None:
        print(f"\n{'=' * 56}\n  {next(self._counter)}/{self.total}  {msg}\n{'=' * 56}")


def write_text(path: Path, content: str) -> None:
    """UTF-8 (BOM 없음) — chcp 65001과 조합해 기존 create_po.bat과 동일한 방식"""
    path.write_text(content, encoding="utf-8", newline="\r\n")


def download(url: str, dest: Path) -> None:
    print(f"  다운로드: {url.rsplit('/', 1)[-1]}")
    req = urllib.request.Request(url, headers={"User-Agent": "noah-build"})
    with urllib.request.urlopen(req, timeout=300) as resp, dest.open("wb") as f:
        shutil.copyfileobj(resp, f)
    print(f"  {dest.stat().st_size / 1024 / 1024:.1f} MB")


# === 버전 / 핀 대조 ========================================================

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


def runtime_packages(python_dir: Path) -> dict[str, str]:
    """배포 런타임에 실제 설치된 패키지 {정규화된 이름: 버전}"""
    result = subprocess.run(
        [str(python_dir / "python.exe"), "-B", "-c",
         "import importlib.metadata, json\n"
         "print(json.dumps({d.metadata['Name']: d.version"
         " for d in importlib.metadata.distributions() if d.metadata['Name']}))\n"],
        env=verify_env(),
        capture_output=True, text=True, encoding="utf-8", errors="replace", check=True,
    )
    return {canon(name): ver for name, ver in json.loads(result.stdout).items()}


def dev_env_versions(names: set[str]) -> dict[str, str]:
    """빌드를 실행한 개발 env(po-automate)의 설치 버전 — 없는 패키지는 빠진다"""
    versions: dict[str, str] = {}
    for name in names:
        try:
            versions[name] = importlib.metadata.version(name)
        except importlib.metadata.PackageNotFoundError:
            continue
    return versions


# === 빌드 단계 =============================================================

def prepare(build_root: Path, app_dir: Path) -> None:
    if build_root.exists():
        shutil.rmtree(build_root)
    app_dir.mkdir(parents=True)
    print(f"  {build_root}")


def install_runtime(runtime_url: str, runtime_tag: str, build_root: Path, python_dir: Path) -> None:
    """python-build-standalone 배포본을 내려받아 python_dir에 푼다

    같은 태그를 두 배포판이 함께 쓰므로 캐시는 태그 단위로 %TEMP%에 둔다 —
    대시보드를 빌드한 뒤 문서생성기를 빌드하면 내려받기가 생략된다.
    """
    cache = Path(tempfile.gettempdir()) / f"noah_runtime_{runtime_tag}.tar.gz"
    if not cache.exists():
        download(runtime_url, cache)
    else:
        print(f"  캐시 사용: {cache.name}")

    with tarfile.open(cache) as tar:
        tar.extractall(build_root)  # noqa: S202 — 신뢰된 공식 배포본
    # 아카이브 최상위가 python/ 이므로 app/ 아래로 옮긴다
    shutil.move(str(build_root / "python"), str(python_dir))
    print(f"  {python_dir}")


def install_packages(python_dir: Path, req_path: Path) -> dict[str, str]:
    """패키지 설치 + 3자 대조 (핀 = 배포 런타임 = 개발 env)

    transitive까지 requirements.txt에 고정돼 있고, 여기서 어긋남을 전부 모아
    실패시킨다. 통과하면 배포판은 개발 PC에서 테스트한 버전 조합 그대로다.

    Returns:
        배포 런타임의 {패키지: 버전} — BUILD_INFO.txt에 재사용
    """
    subprocess.run(
        [str(python_dir / "python.exe"), "-m", "pip", "install",
         "-r", str(req_path), "--no-warn-script-location", "-q"],
        check=True,
    )

    pins = parse_pins(req_path.read_text(encoding="utf-8"))
    installed = runtime_packages(python_dir)
    dev = dev_env_versions(set(pins))

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
            f"        새 버전으로 테스트를 통과시킨 뒤 {req_path.name} 핀을 갱신하세요."
        )
    print(f"  OK  핀 {len(pins)}개 = 배포 런타임 = 개발 env")
    return installed


def trim_runtime(python_dir: Path, trim_globs: tuple[str, ...], trim_dirs: tuple[str, ...]) -> None:
    """런타임에서 안 쓰는 것을 지운다 — 무엇을 지울지는 배포판이 정한다

    지운 뒤에는 반드시 각 배포판의 verify()가 "그래도 돌아간다"를 증명해야 한다.
    """
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

    for pattern in trim_globs:
        for path in python_dir.glob(pattern):
            _rm(path)
    for rel in trim_dirs:
        _rm(python_dir / rel)
    print(f"  {removed / 1024 / 1024:.0f} MB 제거")


def write_build_info(
    app_dir: Path,
    version: str,
    packages: dict[str, str],
    runtime_py_version: str,
    runtime_tag: str,
) -> None:
    """app/BUILD_INFO.txt — 배포판의 정체

    기계가 읽는 건 `version:` 줄 하나뿐이다 (noah_gui.read_build_info → GUI 타이틀).
    나머지는 문의 대응용 — "어느 빌드 쓰세요?"에 사용자가 이 파일을 열어 답한다.
    """
    lines = [
        f"version: {version}",
        f"built: {dt.datetime.now():%Y-%m-%d %H:%M}",
        f"runtime: python {runtime_py_version} (python-build-standalone {runtime_tag})",
        "packages:",
    ]
    lines += [
        f"  {name} {ver}"
        for name, ver in sorted(packages.items())
        if name not in IGNORE_PACKAGES  # 트리밍으로 제거되는 것은 싣지 않는다
    ]
    write_text(app_dir / "BUILD_INFO.txt", "\n".join(lines) + "\n")
    print(f"  BUILD_INFO.txt  v{version}")


def make_zip(zip_path: Path, build_root: Path, app_dir: Path, python_dir: Path) -> None:
    if zip_path.exists():
        zip_path.unlink()
    zip_path.parent.mkdir(parents=True, exist_ok=True)

    # 앱 쪽 __pycache__는 넣지 않는다 (검증 단계가 만들 수 있고, 첫 실행 때 다시 생긴다).
    # python/ 안의 것은 표준 라이브러리 로딩을 빠르게 하므로 그대로 둔다.
    for cache in app_dir.rglob("__pycache__"):
        if python_dir not in cache.parents:
            shutil.rmtree(cache, ignore_errors=True)

    files = [p for p in build_root.rglob("*") if p.is_file()]
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for i, path in enumerate(files, 1):
            zf.write(path, path.relative_to(build_root))
            if i % 2000 == 0:
                print(f"  {i}/{len(files)}...")

    raw = sum(p.stat().st_size for p in files)
    print(f"\n  {zip_path}")
    print(f"  {zip_path.stat().st_size / 1024 / 1024:.0f} MB "
          f"(압축 전 {raw / 1024 / 1024:.0f} MB, 파일 {len(files):,}개)")
