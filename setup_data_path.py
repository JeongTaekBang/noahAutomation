"""
데이터 파일 경로 준비 — noah_menu.bat 전용 콘솔 마법사
=====================================================

배포판에는 user_settings.py가 없어 `noah_config.ini`가 데이터 파일 위치를 정한다.
GUI는 첫 실행 마법사(파일 대화상자)로 ini를 만드는데, 메뉴(bat)에는 창이 없으니
같은 일을 콘솔에서 한다. 탐색·검증·ini 쓰기는 **GUI 것을 그대로 쓴다** — 두 진입점이
같은 ini를 다른 규칙으로 쓰면 GUI에서 고른 경로와 메뉴가 쓰는 경로가 갈린다.

동료 PC 시나리오: 공유받은 OneDrive 폴더에 NOAH_SO_PO_DN.xlsx가 있다. 첫 실행에
그 파일을 찾아 보여주고 Enter 한 번으로 확정한다. 이후 실행은 묻지 않는다.

**찾은 것이 여럿이면 전부 보여주고 고르게 한다** — 첫 번째 것을 그냥 쓰면 안 된다.
실측(2026-09-14, 빌드 폴더 첫 실행): 진짜 파일은 `바탕 화면\\업무\\NOAH ACTUATION\\`인데
`문서\\`에 낡은 사본이 있어 BFS가 그쪽을 먼저 잡았다. 최근 수정 순으로 보여주면
쓰고 있는 파일이 1번에 온다 (낡은 사본은 수정일이 멈춰 있다).

Usage:
    python setup_data_path.py            # 미설정이면 OneDrive 탐색 → 확인 → ini 기록
    python setup_data_path.py --print    # 현재 경로 한 줄 출력 (메뉴 머리말용, 묻지 않음)
    python setup_data_path.py --change   # 현재 경로가 있어도 다시 지정

종료 코드: 0 = 데이터 파일 확정, 1 = 지정 안 됨 (사용자 취소 · 입력 없음)

`noah_gui`는 tkinter를 끌고 와 임포트에 0.6초가 든다. `--print`는 메뉴를 그릴 때마다
불리므로 그 경로에서는 임포트하지 않는다 — config.py만 단독 로드해 결론을 읽는다.
"""

from __future__ import annotations

import argparse
import datetime as dt
import importlib.util
import sys
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
DATA_FILE_NAME = "NOAH_SO_PO_DN.xlsx"

UNSET = "(미설정)"


# === 현재 설정 ==============================================================

def configured_file() -> Path | None:
    """config.py가 결론 내린 데이터 파일 — 실제로 존재할 때만

    user_settings.py(개발 PC) → ini(배포판) 우선순위는 config.py가 정한다.
    여기서 다시 계산하지 않고 그 결론을 읽는다 (noah_gui.effective_data_file과 같은 방식).
    ini가 낡은 경로를 가리키면(공유 폴더를 옮긴 경우) None — 다시 찾게 된다.
    """
    config_path = APP_DIR / "po_generator" / "config.py"
    if not config_path.exists():
        return None
    try:
        spec = importlib.util.spec_from_file_location("_noah_config_probe", config_path)
        if spec is None or spec.loader is None:
            return None
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        path = Path(module.NOAH_SO_PO_DN_FILE)
    except Exception:
        return None
    return path if path.is_file() else None


# === 입력 해석 / 검증 =======================================================

def normalize_input(text: str) -> Path | None:
    """사람이 붙여넣은 경로 → 데이터 파일 경로

    탐색기의 [경로 복사]는 따옴표를 붙이고, 폴더만 붙여넣는 사람도 있다.
    빈 입력은 None (호출부가 '기본값 사용' 또는 '취소'로 해석).
    """
    text = text.strip().strip('"').strip("'").strip()
    if not text:
        return None
    path = Path(text).expanduser()
    if path.is_dir():
        path = path / DATA_FILE_NAME
    return path


def check(path: Path) -> str | None:
    """데이터 파일로 쓸 수 있으면 None, 아니면 사람이 읽을 사유

    잠긴 파일(Excel이 열어 둠)은 시트를 못 읽어 검증이 안 된다 — 이름이 맞으면
    통과시킨다. 잠김 자체는 문서 생성 때 CLI가 다시 안내한다.
    """
    import noah_gui  # 지연 임포트 — tkinter (위 모듈 docstring 참조)

    if not path.is_file():
        return f"파일이 없습니다: {path}"
    if noah_gui.data_file_locked(path):
        return None
    if not noah_gui.looks_like_data_file(path):
        return (f"NOAH 데이터 파일이 아닙니다 (시트 {', '.join(noah_gui.REQUIRED_SHEETS)} 없음): "
                f"{path}")
    return None


def _modified(path: Path) -> str:
    try:
        return dt.datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
    except OSError:
        return "?"


def ask(candidates: list[Path]) -> Path | None:
    """콘솔에서 고르게 한다 — Enter는 1번, 번호 또는 경로 입력. 후보가 없으면 빈 Enter가 취소

    검증에 실패하면 사유를 보여주고 다시 묻는다. 세 번 틀리면 취소 —
    무한 루프에 갇힌 창은 닫는 것 말고 방법이 없다.
    """
    if candidates:
        print("  찾은 파일 (최근 수정 순):")
        for i, path in enumerate(candidates, 1):
            print(f"    [{i}] {path}   (수정 {_modified(path)})")
        prompt = ("  번호를 고르세요 (Enter = 1번). 다른 파일이면 경로 입력 (폴더도 가능): "
                  if len(candidates) > 1
                  else "  이 파일을 쓰려면 Enter, 다른 파일이면 경로 입력 (폴더도 가능): ")
    else:
        print(f"  {DATA_FILE_NAME} 를 찾지 못했습니다.")
        prompt = "  파일 또는 폴더 경로 입력 (빈 Enter = 취소): "

    for _ in range(3):
        try:
            raw = input(prompt).strip()
        except EOFError:
            raw = ""
        if raw.isdigit() and 1 <= int(raw) <= len(candidates):
            return candidates[int(raw) - 1]
        path = normalize_input(raw)
        if path is None:
            return candidates[0] if candidates else None
        reason = check(path)
        if reason is None:
            return path
        print(f"  [오류] {reason}")
    return None


# === 흐름 ===================================================================

def find_candidates() -> list[Path]:
    """OneDrive 탐색 결과 중 데이터 파일로 검증된 것 — 최근 수정 순 (쓰는 파일이 앞에 온다)"""
    import noah_gui  # 지연 임포트

    found = [p for p in noah_gui.find_data_files() if check(p) is None]
    return sorted(found, key=lambda p: p.stat().st_mtime, reverse=True)


def ensure(change: bool = False, interactive: bool | None = None) -> Path | None:
    """데이터 파일이 정해진 상태로 만든다 (필요하면 ini 기록)

    Returns:
        확정된 데이터 파일. None이면 사용자가 지정하지 않은 것.

    비대화형(파이프)에서는 물을 수 없다 — 찾은 것 중 첫째를 쓰고, 없으면 None.
    개발 PC는 user_settings.py가 파일을 정하므로 change가 아니면 아무것도 하지 않는다.
    """
    import noah_gui  # 지연 임포트

    if interactive is None:
        interactive = sys.stdin.isatty()

    current = configured_file()
    if current is not None and not change:
        return current

    # change면 지금 경로가 1번(Enter = 유지)이고 그 뒤에 새로 찾은 것을 붙인다
    candidates = [current] if current else []
    candidates += [p for p in find_candidates() if p != current]

    chosen = ask(candidates) if interactive else (candidates[0] if candidates else None)
    if chosen is None:
        return None

    noah_gui.write_ini(chosen.parent)
    print(f"  데이터 파일: {chosen}")
    print(f"  (설정 저장: {noah_gui.INI_FILE})")
    return chosen


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="NOAH 데이터 파일 경로 준비 (noah_menu.bat용)")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument('--print', dest='print_only', action='store_true',
                      help="현재 경로 한 줄 출력 (묻지 않음)")
    mode.add_argument('--change', action='store_true',
                      help="현재 경로가 있어도 다시 지정")
    args = parser.parse_args(argv)

    if args.print_only:
        current = configured_file()
        print(f"  데이터 파일: {current if current else UNSET}")
        return 0

    if not args.change and configured_file() is not None:
        return 0

    print()
    print("  데이터 파일(NOAH_SO_PO_DN.xlsx) 지정" if args.change
          else "  처음 실행입니다 — 데이터 파일(NOAH_SO_PO_DN.xlsx)을 찾습니다...")
    return 0 if ensure(change=args.change) is not None else 1


if __name__ == "__main__":
    sys.exit(main())
