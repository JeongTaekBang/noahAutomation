"""대시보드 배포판 빌드 — 조용히 틀리는 것들

빌드를 돌리지 않고 검사한다 (실제 빌드는 런타임 내려받기 포함 10분짜리다).
여기서 잡고 싶은 것은 **배포판에서만 드러나는 어긋남**이다:

  - dashboard.py가 쓰는 것이 배포에 안 담기는 것 (import·sql 파일)
  - 빌드한 사람의 경로가 zip에 섞여 나가는 것
  - 재현성 약속(핀 고정)이 조용히 깨지는 것

가장 중요한 건 `test_dashboard_imports_are_packaged`다. dashboard.py에 새 import를
넣고 배포 목록에 안 넣으면, 개발 PC에서는 멀쩡한데 받는 사람 PC에서만 죽는다.
예전 빌드가 정확히 그 형태로 4개월 동안 깨져 있었다.
"""

from __future__ import annotations

import ast
import importlib.util
from pathlib import Path

import pytest

import noah_gui  # 프로젝트 루트를 sys.path에 올리기 위한 앵커 (test_noah_gui와 동일)

PROJECT_ROOT = Path(noah_gui.__file__).resolve().parent
DIST_DIR = PROJECT_ROOT / "dashboard_dist"


def _load_build_module():
    """dashboard_dist/build_portable_dashboard.py 로드 (패키지가 아니라 경로로 연다)

    모듈 레벨은 상수 정의뿐이라 임포트 부작용이 없다.
    """
    path = DIST_DIR / "build_portable_dashboard.py"
    spec = importlib.util.spec_from_file_location("_build_dashboard_probe", path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _toplevel_imports(py: Path) -> set[str]:
    """모듈 최상위에서 import하는 최상위 패키지명"""
    tree = ast.parse(py.read_text(encoding="utf-8"))
    names: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            names.update(a.name.split(".")[0] for a in node.names)
        elif isinstance(node, ast.ImportFrom) and node.level == 0 and node.module:
            names.add(node.module.split(".")[0])
    return names


# === 배포 목록 ↔ 실제 사용 ==================================================

def test_dashboard_imports_are_packaged():
    """dashboard.py가 최상위에서 import하는 서드파티가 전부 핀에 있는가

    표준 라이브러리와 동봉 소스(po_generator)는 제외한다.
    """
    build = _load_build_module()
    pins = build.parse_pins((DIST_DIR / "requirements.txt").read_text(encoding="utf-8"))

    import sys
    stdlib = sys.stdlib_module_names
    bundled = {"po_generator"}
    # import 이름 → 배포 패키지 이름이 다른 것들
    alias = {"dateutil": "python-dateutil", "PIL": "pillow", "win32com": "pywin32"}

    third_party = _toplevel_imports(PROJECT_ROOT / "dashboard.py") - stdlib - bundled
    missing = [m for m in third_party if build.canon(alias.get(m, m)) not in pins]
    assert not missing, (
        f"dashboard.py가 쓰는데 배포 핀에 없음: {missing} — "
        "받는 사람 PC에서 대시보드가 아예 안 뜹니다."
    )


def test_packaged_sql_covers_what_dashboard_reads():
    """dashboard.py가 읽는 .sql 파일이 sql/ 에 실제로 있는가

    배포는 sql/ 폴더를 통째로 담으므로, 원본에 있으면 배포에도 있다.
    """
    source = (PROJECT_ROOT / "dashboard.py").read_text(encoding="utf-8")
    referenced = {name for name in ("order_book.sql",) if name in source}
    assert referenced, "dashboard.py가 읽는 sql 파일을 찾지 못했습니다 (테스트가 낡았을 수 있음)"
    for name in referenced:
        assert (PROJECT_ROOT / "sql" / name).exists(), f"sql/{name} 없음"


def test_po_generator_entrypoints_stay_light():
    """대시보드가 쓰는 po_generator 경로가 무거운 의존성을 끌지 않는가

    패키지 __init__이 excel_generator를 eager import 하던 시절에는
    `from po_generator.config import DB_FILE` 한 줄에 pandas·openpyxl·xlwings가
    딸려왔다. 그 상태로 되돌아가면 대시보드 배포판이 Excel 라이브러리 없이는
    뜨지 못한다 — 그런데 그건 배포해 보기 전에는 드러나지 않는다.
    """
    heavy = {"pandas", "openpyxl", "xlwings", "win32com", "numpy"}
    for rel in ("po_generator/__init__.py", "po_generator/config.py",
                "po_generator/db_schema.py"):
        found = _toplevel_imports(PROJECT_ROOT / rel) & heavy
        assert not found, (
            f"{rel}가 {found}를 최상위에서 import합니다 — "
            "대시보드 배포판이 이 패키지를 싣게 됩니다."
        )


# === 재현성 / 유출 방지 =====================================================

def test_requirements_pins_direct_deps():
    """전부 핀이고 직접 의존성을 담는가 (transitive 전수는 빌드가 3자 대조로 강제)"""
    build = _load_build_module()
    req = (DIST_DIR / "requirements.txt").read_text(encoding="utf-8")
    pins = build.parse_pins(req)  # 범위 지정이 섞이면 여기서 ValueError
    assert {"streamlit", "plotly", "pandas", "pywin32"} <= set(pins)


def test_excel_libs_are_not_shipped():
    """대시보드는 Excel을 건드리지 않는다 — 다시 들어오면 알아채야 한다"""
    build = _load_build_module()
    pins = build.parse_pins((DIST_DIR / "requirements.txt").read_text(encoding="utf-8"))
    assert "xlwings" not in pins
    assert "openpyxl" not in pins


def test_user_settings_is_excluded():
    """빌드한 사람의 OneDrive 경로가 배포판에 실려 나가면 안 된다"""
    build = _load_build_module()
    assert "user_settings.py" in build.EXCLUDE_FILES


def test_build_root_is_outside_project():
    """빌드는 %TEMP%에서 — 프로젝트는 OneDrive 안이라 런타임이 통째로 동기화된다"""
    build = _load_build_module()
    assert PROJECT_ROOT not in build.BUILD_ROOT.parents


def test_build_root_differs_from_cli_dist():
    """두 배포판을 연달아 빌드해도 서로의 작업 폴더를 지우지 않아야 한다"""
    dash = _load_build_module()
    spec = importlib.util.spec_from_file_location(
        "_build_gui_probe", PROJECT_ROOT / "cli_dist" / "build_portable_gui.py")
    gui = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(gui)
    assert dash.BUILD_ROOT != gui.BUILD_ROOT
    assert dash.ZIP_PATH != gui.ZIP_PATH


# === 트리밍 안전장치 ========================================================

def test_streamlit_frontend_is_not_trimmed():
    """streamlit/static은 프런트엔드 번들 그 자체다 — 지우면 빈 화면이 된다"""
    build = _load_build_module()
    patterns = [*build.TRIM_GLOBS, *build.TRIM_DIRS]
    assert not [p for p in patterns if "streamlit" in p]


def test_pyarrow_compute_is_kept():
    """streamlit이 pyarrow.compute를 로드한다 (실측) — arrow_compute.dll은 남겨야 한다"""
    build = _load_build_module()
    patterns = [*build.TRIM_GLOBS, *build.TRIM_DIRS]
    for kept in ("arrow_compute", "arrow.dll", "arrow_python"):
        assert not [p for p in patterns if kept in p], f"{kept}를 지우면 표가 안 그려집니다"


@pytest.mark.parametrize("trimmed", ["arrow_flight", "parquet", "arrow_substrait"])
def test_pyarrow_unused_components_are_trimmed(trimmed):
    """안 쓰는 pyarrow 컴포넌트는 실제로 지워지는가 (합쳐서 ~40MB)"""
    build = _load_build_module()
    patterns = [*build.TRIM_GLOBS, *build.TRIM_DIRS]
    assert [p for p in patterns if trimmed in p]
