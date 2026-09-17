"""
noah_menu.bat / setup_data_path.py 테스트
==========================================

동료 PC용 콘솔 메뉴는 배포판에서만 쓰이므로 개발 PC에서는 안 깨져도 모른다.
여기서 잡고 싶은 것은 **배포판에서만 드러나는 어긋남**이다:

  - 메뉴가 부르는 스크립트가 배포 목록(APP_FILES)에 없는 것 (GUI의 DOC_TYPES 대조와 같은 패턴)
  - 메뉴에 [국내]/[해외] 문서 외의 기능(DB 동기화·마감·대사·DN 쓰기)이 새는 것
  - create_po.bat이 문서 블록을 다시 갖게 되어 두 메뉴가 갈리는 것 (2026-08-07 교훈)
  - bat 인코딩/줄끝 — LF만 있는 bat은 goto가 라벨을 못 찾는 경우가 있다

마법사(setup_data_path)는 콘솔 입력을 함수로 받으므로 input을 바꿔 끼워 검사한다.
"""

from __future__ import annotations

import importlib.util
import re
import zipfile
from pathlib import Path

import pytest

import noah_gui
import setup_data_path

PROJECT_ROOT = Path(noah_gui.__file__).resolve().parent
MENU_BAT = PROJECT_ROOT / "noah_menu.bat"
FULL_BAT = PROJECT_ROOT / "create_po.bat"

DOC_SCRIPTS = {"create_po.py", "create_ts.py", "create_pi.py", "create_fi.py",
               "create_oc.py", "create_ci.py", "create_pl.py"}
# 동료 PC 메뉴에 있으면 안 되는 것 — 마스터 워크북·DB를 건드리거나 운영자 전용
FORBIDDEN_SCRIPTS = {"sync_db.py", "close_period.py", "create_dn.py", "reconcile_po.py",
                     "reconcile_so.py", "reconcile_ind.py", "dashboard.py",
                     "delivery_status.py"}


def _load_build_module():
    path = PROJECT_ROOT / "cli_dist" / "build_portable_gui.py"
    spec = importlib.util.spec_from_file_location("_build_portable_gui_probe", path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _bat_text(path: Path) -> str:
    return path.read_bytes().decode("utf-8")


def _invoked_scripts(text: str) -> set[str]:
    """bat이 `%~dp0xxx.py`로 부르는 스크립트명"""
    return set(re.findall(r"%~dp0(\w+\.py)", text))


# === 배포 목록 ↔ 메뉴 =======================================================

def test_menu_scripts_are_packaged():
    """메뉴가 부르는 스크립트가 전부 배포판에 들어가는가 — 빠지면 그 번호만 동료 PC에서 죽는다"""
    build = _load_build_module()
    missing = sorted(_invoked_scripts(_bat_text(MENU_BAT)) - set(build.APP_FILES))
    assert not missing, f"배포판 APP_FILES에 빠진 스크립트: {missing}"


def test_menu_and_wizard_are_packaged():
    build = _load_build_module()
    assert "noah_menu.bat" in build.APP_FILES
    assert "setup_data_path.py" in build.APP_FILES


def test_menu_has_desktop_shortcut():
    build = _load_build_module()
    assert "noah_menu.bat" in build.SHORTCUTS.values()


def test_help_smoke_skips_bat():
    """--help 스모크는 파이썬 CLI만 — .bat에 --help를 주면 메뉴가 뜨고 멈춘다"""
    build = _load_build_module()
    clis = build.cli_scripts()
    assert all(name.endswith(".py") for name in clis)
    assert "noah_gui.py" not in clis
    assert "setup_data_path.py" in clis


# === 메뉴 범위 ==============================================================

def test_menu_covers_exactly_domestic_and_overseas_docs():
    """[국내] PO·TS + [해외] PI·FI·OC·CI·PL — 그 이상은 동료 PC에 열지 않는다"""
    invoked = _invoked_scripts(_bat_text(MENU_BAT))
    assert invoked == DOC_SCRIPTS | {"setup_data_path.py"}
    assert not invoked & FORBIDDEN_SCRIPTS


def test_full_menu_delegates_documents():
    """create_po.bat은 문서 7종을 직접 부르지 않고 noah_menu.bat에 위임한다

    블록이 양쪽에 있으면 옵션을 한쪽에만 넣는 실수가 반복된다 (tasks/lessons.md 2026-08-07).
    create_po.py는 [H] 발주 이력 조회(--history)로만 남는다.
    """
    text = _bat_text(FULL_BAT)
    for script in sorted(DOC_SCRIPTS - {"create_po.py"}):
        assert script not in text, f"create_po.bat이 {script}를 직접 부른다"
    for line in text.splitlines():
        if "create_po.py" in line:
            assert "--history" in line, f"create_po.bat에 문서 생성 블록이 남아 있다: {line}"
    for key in ("po", "ts", "pi", "fi", "oc", "ci", "pl"):
        assert f'call "%~dp0noah_menu.bat" {key}' in text


def test_menu_accepts_delegation_keys():
    """위임 키(po/ts/...)를 메뉴가 받는가 — 한쪽만 바뀌면 그 항목이 '올바른 번호를 입력'으로 끝난다"""
    text = _bat_text(MENU_BAT)
    for key in ("po", "ts", "pi", "fi", "oc", "ci", "pl"):
        assert f'if /i "%CHOICE%"=="{key}" goto create_{key}' in text
    assert "if defined ONESHOT exit /b 0" in text


@pytest.mark.parametrize("bat", [MENU_BAT, FULL_BAT])
def test_bat_is_utf8_crlf_without_bom(bat):
    raw = bat.read_bytes()
    assert not raw.startswith(b"\xef\xbb\xbf"), "BOM이 있으면 @echo off 줄이 깨진다"
    raw.decode("utf-8")
    assert b"\r\n" in raw
    assert b"\n" not in raw.replace(b"\r\n", b""), "LF만 있는 줄이 섞여 있다"
    assert raw.startswith(b"@echo off\r\nchcp 65001")


# === 마법사 (setup_data_path) ================================================

def _fake_data_file(folder: Path, sheets=noah_gui.REQUIRED_SHEETS) -> Path:
    """필수 시트명만 든 최소 xlsx (noah_gui.sheet_names가 workbook.xml만 읽는다)"""
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / setup_data_path.DATA_FILE_NAME
    xml = "<workbook><sheets>" + "".join(
        f'<sheet name="{s}" sheetId="{i}"/>' for i, s in enumerate(sheets, 1)
    ) + "</sheets></workbook>"
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("xl/workbook.xml", xml)
    return path


def test_normalize_input_strips_quotes_and_accepts_folder(tmp_path):
    data = _fake_data_file(tmp_path)
    assert setup_data_path.normalize_input(f'  "{tmp_path}"  ') == data
    assert setup_data_path.normalize_input(str(data)) == data
    assert setup_data_path.normalize_input("   ") is None


def test_check_rejects_missing_and_foreign_files(tmp_path):
    assert "없습니다" in setup_data_path.check(tmp_path / "x.xlsx")
    other = _fake_data_file(tmp_path, sheets=("Sheet1",))
    assert "아닙니다" in setup_data_path.check(other)
    real = _fake_data_file(tmp_path / "ok")
    assert setup_data_path.check(real) is None


def test_ensure_noop_when_configured(monkeypatch, tmp_path):
    """이미 정해져 있으면 묻지도 쓰지도 않는다 — 매 실행이 이 경로다"""
    data = _fake_data_file(tmp_path)
    monkeypatch.setattr(setup_data_path, "configured_file", lambda: data)
    monkeypatch.setattr(noah_gui, "write_ini", lambda folder: pytest.fail("ini를 쓰면 안 된다"))
    monkeypatch.setattr("builtins.input", lambda *_: pytest.fail("물으면 안 된다"))
    assert setup_data_path.ensure(interactive=True) == data


def test_ensure_first_run_accepts_found_file_with_enter(monkeypatch, tmp_path):
    """동료 PC 첫 실행: OneDrive에서 찾은 파일을 보여주고 Enter → ini 기록"""
    data = _fake_data_file(tmp_path)
    written: list[Path] = []
    monkeypatch.setattr(setup_data_path, "configured_file", lambda: None)
    monkeypatch.setattr(noah_gui, "find_data_files", lambda *a, **k: [data])
    monkeypatch.setattr(noah_gui, "write_ini", written.append)
    monkeypatch.setattr("builtins.input", lambda *_: "")
    assert setup_data_path.ensure(interactive=True) == data
    assert written == [tmp_path]


def test_ensure_first_run_typed_path_overrides_found(monkeypatch, tmp_path):
    found = _fake_data_file(tmp_path / "stale")
    typed = _fake_data_file(tmp_path / "shared")
    written: list[Path] = []
    monkeypatch.setattr(setup_data_path, "configured_file", lambda: None)
    monkeypatch.setattr(noah_gui, "find_data_files", lambda *a, **k: [found])
    monkeypatch.setattr(noah_gui, "write_ini", written.append)
    monkeypatch.setattr("builtins.input", lambda *_: str(typed.parent))  # 폴더만 붙여넣음
    assert setup_data_path.ensure(interactive=True) == typed
    assert written == [typed.parent]


def test_ensure_lists_all_copies_newest_first_and_picks_by_number(monkeypatch, tmp_path, capsys):
    """사본이 여럿이면 전부 보여주고 고르게 한다 — 첫 BFS 결과를 그냥 쓰면 낡은 사본이 잡힌다

    실측(2026-09-14): `문서\`에 8/7자 사본, `NOAH ACTUATION\`에 진짜(9/14). 탐색 순서는
    사본이 먼저였다. 최근 수정 순이면 진짜가 1번이라 Enter만 쳐도 맞는다.
    """
    import os
    stale = _fake_data_file(tmp_path / "문서")
    real = _fake_data_file(tmp_path / "NOAH ACTUATION")
    os.utime(stale, (1_700_000_000, 1_700_000_000))
    os.utime(real, (1_800_000_000, 1_800_000_000))
    written: list[Path] = []
    monkeypatch.setattr(setup_data_path, "configured_file", lambda: None)
    monkeypatch.setattr(noah_gui, "find_data_files", lambda *a, **k: [stale, real])  # 탐색 순서: 사본 먼저
    monkeypatch.setattr(noah_gui, "write_ini", written.append)

    monkeypatch.setattr("builtins.input", lambda *_: "")  # Enter = 1번 = 최근 수정
    assert setup_data_path.ensure(interactive=True) == real
    out = capsys.readouterr().out
    assert out.index(str(real)) < out.index(str(stale))

    monkeypatch.setattr("builtins.input", lambda *_: "2")  # 번호로 고르기
    assert setup_data_path.ensure(interactive=True) == stale
    assert written == [real.parent, stale.parent]


def test_ensure_change_shows_current_first_then_found(monkeypatch, tmp_path):
    current = _fake_data_file(tmp_path / "cur")
    other = _fake_data_file(tmp_path / "other")
    written: list[Path] = []
    monkeypatch.setattr(setup_data_path, "configured_file", lambda: current)
    monkeypatch.setattr(noah_gui, "find_data_files", lambda *a, **k: [current, other])
    monkeypatch.setattr(noah_gui, "write_ini", written.append)
    monkeypatch.setattr("builtins.input", lambda *_: "2")
    assert setup_data_path.ensure(change=True, interactive=True) == other
    assert written == [other.parent]


def test_ensure_nothing_found_and_empty_input_cancels(monkeypatch):
    monkeypatch.setattr(setup_data_path, "configured_file", lambda: None)
    monkeypatch.setattr(noah_gui, "find_data_files", lambda *a, **k: [])
    monkeypatch.setattr(noah_gui, "write_ini", lambda folder: pytest.fail("ini를 쓰면 안 된다"))
    monkeypatch.setattr("builtins.input", lambda *_: "")
    assert setup_data_path.ensure(interactive=True) is None


def test_ensure_gives_up_after_three_bad_inputs(monkeypatch, tmp_path):
    """무한 루프에 갇힌 창은 닫는 것 말고 방법이 없다 — 세 번 틀리면 취소"""
    monkeypatch.setattr(setup_data_path, "configured_file", lambda: None)
    monkeypatch.setattr(noah_gui, "find_data_files", lambda *a, **k: [])
    monkeypatch.setattr(noah_gui, "write_ini", lambda folder: pytest.fail("ini를 쓰면 안 된다"))
    answers = iter([str(tmp_path / "없음1.xlsx"), str(tmp_path / "없음2.xlsx"), str(tmp_path / "없음3.xlsx")])
    monkeypatch.setattr("builtins.input", lambda *_: next(answers))
    assert setup_data_path.ensure(interactive=True) is None


def test_ensure_non_interactive_uses_found_without_asking(monkeypatch, tmp_path):
    """파이프 실행(비대화형)은 물을 수 없다 — 찾은 것을 그대로 쓴다"""
    data = _fake_data_file(tmp_path)
    written: list[Path] = []
    monkeypatch.setattr(setup_data_path, "configured_file", lambda: None)
    monkeypatch.setattr(noah_gui, "find_data_files", lambda *a, **k: [data])
    monkeypatch.setattr(noah_gui, "write_ini", written.append)
    monkeypatch.setattr("builtins.input", lambda *_: pytest.fail("물으면 안 된다"))
    assert setup_data_path.ensure(interactive=False) == data
    assert written == [tmp_path]


def test_ensure_change_keeps_current_on_enter(monkeypatch, tmp_path):
    """[P] 경로 변경: 현재 경로가 기본값이라 Enter는 '유지'"""
    data = _fake_data_file(tmp_path)
    written: list[Path] = []
    monkeypatch.setattr(setup_data_path, "configured_file", lambda: data)
    monkeypatch.setattr(noah_gui, "find_data_files", lambda *a, **k: [])
    monkeypatch.setattr(noah_gui, "write_ini", written.append)
    monkeypatch.setattr("builtins.input", lambda *_: "")
    assert setup_data_path.ensure(change=True, interactive=True) == data
    assert written == [tmp_path]


def test_main_print_never_asks(monkeypatch, capsys):
    monkeypatch.setattr(setup_data_path, "configured_file", lambda: None)
    monkeypatch.setattr("builtins.input", lambda *_: pytest.fail("물으면 안 된다"))
    assert setup_data_path.main(["--print"]) == 0
    assert setup_data_path.UNSET in capsys.readouterr().out


def test_main_exit_code_reflects_result(monkeypatch, tmp_path):
    data = _fake_data_file(tmp_path)
    monkeypatch.setattr(setup_data_path, "configured_file", lambda: data)
    assert setup_data_path.main([]) == 0

    monkeypatch.setattr(setup_data_path, "configured_file", lambda: None)
    monkeypatch.setattr(setup_data_path, "ensure", lambda **k: None)
    assert setup_data_path.main([]) == 1


def test_configured_file_matches_config_on_dev_pc():
    """개발 PC: user_settings.py가 정한 파일 = config.py 결론 (있을 때만 경로를 준다)"""
    from po_generator import config

    expected = Path(config.NOAH_SO_PO_DN_FILE)
    result = setup_data_path.configured_file()
    assert result == (expected if expected.is_file() else None)
