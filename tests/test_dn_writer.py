"""
dn_writer.py 테스트
===================

`DN_국내` 표에 행을 덧붙이는 쓰기 경로. Excel COM 없이 워크북 객체를 흉내 내
**"쓰기가 실패했는데 성공이라고 보고하는" 경로가 없는지**만 본다.

이게 회귀 방지의 전부인 이유: 2026-08-07에 실제로 그런 일이 났다.
`create_dn.py`가 읽기용 `pd.ExcelFile` 핸들을 열어 둔 채로 Excel에게 같은 파일을
열라고 해서, Excel이 **읽기 전용**으로 열었다. 쓰기 자체는 메모리에서 멀쩡히 되고
표 범위도 늘어나 "완료: 4행 추가"라고 보고했는데, `Save()`는 예외도 없이 무시됐고
파일은 그대로였다.

순서가 규약인 이유도 같은 사건에서 나왔다: **Excel이 워크북을 열면 읽기조차 막혀서**
백업 사본을 뜰 수 없다. 그래서 (1) 닫혀 있는지 확인 → (2) 백업 → (3) Excel 열기 순이다.
"""

from __future__ import annotations

import sys
from pathlib import Path
from types import SimpleNamespace

import pandas as pd
import pytest

from po_generator import dn_writer
from po_generator.dn_recorder import DnLine
from po_generator.dn_writer import append_lines

SHEET = 'DN_국내'

LINES = [DnLine(
    dn_id='DND-2026-0753', so_id='SOD-2026-0474', line_item=2, qty=20,
    currency='KRW', ship_date=pd.Timestamp(2026, 8, 7),
    tax_date=pd.Timestamp(2026, 8, 7), remarks=None, seq=1432,
)]


# === COM 흉내 ==============================================================

class FakeBook:
    """`save()`가 파일을 실제로 건드리는지 여부를 흉내 낼 수 있는 워크북"""

    def __init__(self, path: Path, *, read_only: bool = False,
                 persists: bool = True, save_raises: bool = False):
        self.api = SimpleNamespace(ReadOnly=read_only, Saved=True)
        self.sheets = {SHEET: object()}
        self._path = path
        self._persists = persists
        self._save_raises = save_raises
        self.save_called = False
        self.closed = False

    def save(self) -> None:
        self.save_called = True
        if self._save_raises:
            raise Exception("Microsoft Excel cannot access the file ...")
        if self._persists:
            # 실제 저장처럼 파일 내용을 바꾼다 (mtime/size가 움직인다)
            self._path.write_bytes(self._path.read_bytes() + b'written')

    def close(self) -> None:
        self.closed = True


class FakeBooks(list):
    """`app.books` — 열 수 있고(`open`) 순회할 수 있어야 한다(정리 루프)"""

    def __init__(self, opener):
        super().__init__()
        self._opener = opener

    def open(self, path):
        return self._opener(path)


class FakeExcelApp:
    def __init__(self, book: FakeBook | None, open_raises: bool = False):
        self._book = book
        self._open_raises = open_raises
        self.display_alerts = True
        self.screen_updating = True
        self.calculated = False
        self.quit_called = False
        self.books = FakeBooks(self._open)

    def _open(self, path):
        if self._open_raises:
            raise Exception("Microsoft Excel cannot access the file ...")
        self._book.app = self          # book.app.calculate() 경로
        self.books.append(self._book)
        return self._book

    def calculate(self) -> None:
        self.calculated = True

    def quit(self) -> None:
        self.quit_called = True


@pytest.fixture
def workbook(tmp_path: Path) -> Path:
    path = tmp_path / "NOAH_SO_PO_DN.xlsx"
    path.write_bytes(b'original workbook bytes')
    return path


@pytest.fixture
def appended(monkeypatch) -> list:
    """`_append_to_table` 호출 기록 (COM 대신)"""
    calls: list = []

    def fake_append(sheet, lines):
        calls.append((sheet, lines))
        return len(lines), 1433, 1433 + len(lines) - 1, 'A1:T1436'

    monkeypatch.setattr(dn_writer, '_append_to_table', fake_append)
    return calls


def use_excel(monkeypatch, app: FakeExcelApp, *, already_open=None) -> None:
    """xlwings 모듈과 '이미 열린 워크북' 탐지를 갈아끼운다"""
    monkeypatch.setattr(dn_writer, '_find_open_book', lambda path: already_open)
    monkeypatch.setitem(sys.modules, 'xlwings',
                        SimpleNamespace(App=lambda **kw: app, apps=[]))


# === 쓰기 전에 막아야 하는 경로 ==============================================

def test_Excel에_열려_있으면_시작도_안_한다(workbook, appended, monkeypatch, tmp_path):
    """열려 있으면 백업 사본을 못 뜬다 — 되돌릴 곳 없이 마스터를 건드리지 않는다"""
    app = FakeExcelApp(FakeBook(workbook))
    use_excel(monkeypatch, app, already_open=object())
    backup_dir = tmp_path / "backup"

    with pytest.raises(RuntimeError, match='Excel에 열려 있습니다'):
        append_lines(workbook, SHEET, LINES, backup_dir=backup_dir)

    assert not appended
    assert not backup_dir.exists()


def test_다른_프로세스가_잡고_있으면_시작도_안_한다(workbook, appended, monkeypatch):
    """`_find_open_book`이 못 봐도(다른 PC·다른 세션) 'r+b' 열기가 잡아낸다"""
    app = FakeExcelApp(FakeBook(workbook))
    use_excel(monkeypatch, app)

    def boom(*a, **kw):
        raise PermissionError(13, 'Permission denied')

    monkeypatch.setattr('builtins.open', boom)

    with pytest.raises(RuntimeError, match='닫고 다시 실행'):
        append_lines(workbook, SHEET, LINES)

    assert not appended


def test_읽기_전용으로_열리면_쓰지_않는다(workbook, appended, monkeypatch):
    """읽기 전용에서는 Save가 조용히 무시된다 — 표를 건드리지도 말아야 한다"""
    book = FakeBook(workbook, read_only=True)
    use_excel(monkeypatch, FakeExcelApp(book))

    with pytest.raises(RuntimeError, match='읽기 전용'):
        append_lines(workbook, SHEET, LINES)

    assert not appended
    assert not book.save_called


def test_추가할_행이_없으면_거부한다(workbook, monkeypatch):
    use_excel(monkeypatch, FakeExcelApp(FakeBook(workbook)))

    with pytest.raises(ValueError):
        append_lines(workbook, SHEET, [])


# === 쓴 뒤에 잡아야 하는 경로 ===============================================

def test_저장이_파일에_안_닿으면_실패로_보고한다(workbook, appended, monkeypatch):
    """COM은 저장 실패를 조용히 삼킨다 — 파일이 안 바뀌었으면 성공이 아니다"""
    book = FakeBook(workbook, persists=False)
    use_excel(monkeypatch, FakeExcelApp(book))

    with pytest.raises(RuntimeError, match='반영되지 않았습니다'):
        append_lines(workbook, SHEET, LINES)

    assert book.save_called        # 저장은 시도했다


def test_저장이_예외로_죽으면_안내로_바꾼다(workbook, appended, monkeypatch):
    use_excel(monkeypatch, FakeExcelApp(FakeBook(workbook, save_raises=True)))

    with pytest.raises(RuntimeError, match='닫고 다시 실행'):
        append_lines(workbook, SHEET, LINES)


def test_열기가_예외로_죽으면_안내로_바꾸고_Excel을_정리한다(workbook, appended, monkeypatch):
    """Excel은 열기·저장 실패를 전부 같은 COM 문구로 돌려준다 — 그대로 보여 주면
    무엇을 해야 하는지 알 수 없다"""
    app = FakeExcelApp(None, open_raises=True)
    use_excel(monkeypatch, app)

    with pytest.raises(RuntimeError, match='닫고 다시 실행'):
        append_lines(workbook, SHEET, LINES)

    assert not appended
    assert app.quit_called         # 띄운 Excel은 반드시 정리한다


# === 정상 경로 =============================================================

def test_정상_쓰기(workbook, appended, monkeypatch, tmp_path):
    book = FakeBook(workbook)
    app = FakeExcelApp(book)
    use_excel(monkeypatch, app)
    backup_dir = tmp_path / "backup"

    result = append_lines(workbook, SHEET, LINES, backup_dir=backup_dir)

    assert result.rows_added == 1
    assert (result.first_row, result.last_row) == (1433, 1433)
    assert result.table_ref == 'A1:T1436'
    assert book.save_called and app.calculated and app.quit_called
    assert result.backup is not None and result.backup.exists()
    # 백업은 쓰기 전 상태여야 되돌릴 수 있다
    assert result.backup.read_bytes() == b'original workbook bytes'


# === 백업 ==================================================================

def test_백업은_최근_것만_남긴다(tmp_path):
    """5MB짜리 워크북이 매 실행마다 쌓이면 OneDrive가 그걸 다 동기화한다"""
    source = tmp_path / "NOAH_SO_PO_DN.xlsx"
    source.write_bytes(b'x')
    backup_dir = tmp_path / "backup"
    backup_dir.mkdir()
    for day in range(1, 6):                      # 지난 백업 5개
        (backup_dir / f"NOAH_SO_PO_DN_2026080{day}_000000.xlsx").write_bytes(b'old')

    dn_writer.backup_workbook(source, backup_dir, keep=3)

    names = sorted(p.name for p in backup_dir.iterdir())
    assert len(names) == 3
    assert names[0] == "NOAH_SO_PO_DN_20260804_000000.xlsx"   # 오래된 것부터 지운다
    assert names[-1].startswith("NOAH_SO_PO_DN_2026")          # 방금 뜬 사본이 남는다


def test_백업_보관수가_0이면_지우지_않는다(tmp_path):
    source = tmp_path / "NOAH_SO_PO_DN.xlsx"
    source.write_bytes(b'x')
    backup_dir = tmp_path / "backup"

    target = dn_writer.backup_workbook(source, backup_dir, keep=0)

    assert target.exists()
