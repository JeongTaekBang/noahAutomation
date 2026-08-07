"""
dn_writer.py 테스트
===================

`DN_국내` 표에 행을 덧붙이는 쓰기 경로. Excel COM 없이 워크북 객체를 흉내 내
**"쓰기가 실패했는데 성공이라고 보고하는" 경로가 없는지**만 본다.

이게 회귀 방지의 전부인 이유: 2026-08-07에 실제로 그런 일이 났다.
워크북이 이미 열려 있어 두 번째 Excel 인스턴스가 **읽기 전용**으로 열었는데,
쓰기 자체는 메모리에서 멀쩡히 되고 표 범위도 늘어나서 "완료: 4행 추가"라고
보고했다. `Save()`는 예외도 없이 조용히 무시됐고 파일은 그대로였다.
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

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

class FakeApi:
    def __init__(self, read_only: bool, saved: bool):
        self.ReadOnly = read_only
        self.Saved = saved


class FakeApp:
    def __init__(self):
        self.calculated = False

    def calculate(self) -> None:
        self.calculated = True


class FakeBook:
    """`save()`가 파일을 실제로 건드리는지 여부를 흉내 낼 수 있는 워크북"""

    def __init__(self, path: Path, *, read_only: bool = False,
                 saved: bool = True, persists: bool = True):
        self.api = FakeApi(read_only, saved)
        self.app = FakeApp()
        self.sheets = {SHEET: object()}
        self._path = path
        self._persists = persists
        self.save_called = False

    def save(self) -> None:
        self.save_called = True
        if self._persists:
            # 실제 저장처럼 파일 내용을 바꾼다 (mtime/size가 움직인다)
            self._path.write_bytes(self._path.read_bytes() + b'written')


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


def use_book(monkeypatch, book: FakeBook | None) -> None:
    monkeypatch.setattr(dn_writer, '_find_open_book', lambda path: book)


# === 실패해야 하는 경로 =====================================================

def test_읽기_전용이면_쓰기_전에_막는다(workbook, appended, monkeypatch):
    """읽기 전용에서는 Save가 조용히 무시된다 — 쓰기를 시작하지도 말아야 한다"""
    book = FakeBook(workbook, read_only=True)
    use_book(monkeypatch, book)

    with pytest.raises(RuntimeError, match='읽기 전용'):
        append_lines(workbook, SHEET, LINES)

    assert not appended            # 표를 건드리지 않았다
    assert not book.save_called


def test_읽기_전용이면_백업도_만들지_않는다(workbook, appended, monkeypatch, tmp_path):
    """못 쓸 걸 알면서 5MB짜리 사본을 쌓지 않는다"""
    use_book(monkeypatch, FakeBook(workbook, read_only=True))
    backup_dir = tmp_path / "backup"

    with pytest.raises(RuntimeError):
        append_lines(workbook, SHEET, LINES, backup_dir=backup_dir)

    assert not backup_dir.exists()


def test_저장이_파일에_안_닿으면_실패로_보고한다(workbook, appended, monkeypatch):
    """COM은 저장 실패를 조용히 삼킨다 — 파일이 안 바뀌었으면 성공이 아니다"""
    book = FakeBook(workbook, persists=False)
    use_book(monkeypatch, book)

    with pytest.raises(RuntimeError, match='반영되지 않았습니다'):
        append_lines(workbook, SHEET, LINES)

    assert book.save_called        # 저장은 시도했다


def test_저장_안_된_변경이_있으면_막는다(workbook, appended, monkeypatch):
    use_book(monkeypatch, FakeBook(workbook, saved=False))

    with pytest.raises(RuntimeError, match='저장하지 않은 변경'):
        append_lines(workbook, SHEET, LINES)

    assert not appended


def test_추가할_행이_없으면_거부한다(workbook, monkeypatch):
    use_book(monkeypatch, FakeBook(workbook))

    with pytest.raises(ValueError):
        append_lines(workbook, SHEET, [])


# === 정상 경로 =============================================================

def test_정상_쓰기(workbook, appended, monkeypatch, tmp_path):
    book = FakeBook(workbook)
    use_book(monkeypatch, book)
    backup_dir = tmp_path / "backup"

    result = append_lines(workbook, SHEET, LINES, backup_dir=backup_dir)

    assert result.rows_added == 1
    assert (result.first_row, result.last_row) == (1433, 1433)
    assert result.table_ref == 'A1:T1436'
    assert book.save_called and book.app.calculated
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
