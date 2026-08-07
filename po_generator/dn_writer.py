"""
DN_국내 워크북 쓰기 (xlwings/COM)
=================================

`dn_recorder.build_plan()`이 계산한 행을 `NOAH_SO_PO_DN.xlsx`의 `DN_국내` 표에 덧붙인다.

**openpyxl로 저장하면 안 된다.** 이 워크북에는 피벗 3개·파워쿼리 연결 14개·쿼리테이블이
들어 있고 openpyxl은 그것들을 읽지도 쓰지도 못해서, 한 번 저장하면 통째로 사라진다.
(`reconcile_*.py`가 openpyxl을 쓰는 건 전부 **새 결과 파일**을 만들 때뿐이다.)

왜 "마지막 행을 복사"하는가
---------------------------
`DN_국내` 표에서 Excel이 자동으로 채워주는 계산 열은 6개뿐이다. `Item`(XLOOKUP)과
`Unit Price`·`AX Project no`(배열 수식)는 열마다 수식이 있는데도 표의 계산 열로
등록돼 있지 않아(`calculatedColumnFormula` 없음) 행을 늘려도 **비어 있는 채로 남는다**.
마지막 데이터 행을 통째로 복사해 내리면 상대참조(`B1432` → `B1433`)가 알아서 따라오고,
그 위에 값 열만 덮어쓰면 된다.
"""

from __future__ import annotations

import datetime as dt
import logging
import shutil
from dataclasses import dataclass
from pathlib import Path

from po_generator.dn_recorder import DnLine

logger = logging.getLogger(__name__)

# 값으로 채우는 열 (나머지는 수식이라 복사한 그대로 둔다)
VALUE_COLUMNS: tuple[str, ...] = (
    'DN_ID', 'SO_ID', 'Line item', 'Qty', 'Currency',
    '출고일', '세금계산서 발행일', 'Remarks', 'IP 여부', 'Seq',
)

BACKUP_KEEP = 10


@dataclass(frozen=True)
class WriteResult:
    rows_added: int
    first_row: int
    last_row: int
    backup: Path | None
    table_ref: str


def backup_workbook(source: Path, backup_dir: Path, keep: int = BACKUP_KEEP) -> Path:
    """쓰기 전 사본 — 되돌릴 곳이 있어야 마스터 파일을 건드릴 수 있다.

    OneDrive 버전 기록이 1차 안전망이지만, 그건 웹에서 찾아 들어가야 하고 동기화가
    밀리면 최신이 아닐 수 있다. 로컬 사본이 있으면 바로 되돌린다.
    """
    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime('%Y%m%d_%H%M%S')
    target = backup_dir / f"{source.stem}_{stamp}{source.suffix}"
    shutil.copy2(source, target)

    olds = sorted(backup_dir.glob(f"{source.stem}_*{source.suffix}"))
    for old in olds[:-keep] if keep > 0 else []:
        try:
            old.unlink()
        except OSError:
            logger.debug("오래된 백업 삭제 실패 (무시): %s", old.name)
    logger.debug("백업 생성: %s", target.name)
    return target


def _assert_closed_and_writable(workbook: Path) -> None:
    """워크북이 닫혀 있고 쓸 수 있는지 — Excel을 띄우기 **전에** 확인

    `open(path, 'r+b')`가 이 판정을 정확히 한다: Excel이(내 PC든 남의 PC든) 파일을
    열고 있으면 PermissionError가 난다. COM을 띄우기 전에 알아야 하는 이유는 두 가지다 —
    (1) Excel이 열면 읽기조차 막혀 백업 사본을 못 뜬다, (2) 그 뒤 `Save()`는 조용히
    실패해 "완료"라고 보고하게 된다.
    """
    open_book = _find_open_book(workbook)
    if open_book is not None:
        raise RuntimeError(
            f"'{workbook.name}'이 Excel에 열려 있습니다 — 닫고 다시 실행하세요.\n"
            "       (열려 있으면 되돌릴 백업 사본을 뜰 수 없습니다)")
    try:
        with open(workbook, 'r+b'):
            pass
    except OSError as e:
        raise RuntimeError(_busy_message(workbook, '쓰기로 열지', e)) from e


def _busy_message(workbook: Path, verb: str, error: Exception) -> str:
    """COM 원문 대신 사람이 읽고 행동할 수 있는 문장으로

    Excel은 열기·저장 실패를 전부 같은 문구('cannot access the file' + 가능성 3가지)로
    돌려준다 — 그대로 보여 주면 무엇을 해야 하는지 알 수 없다. 실제 원인은 거의 항상
    "누가 잡고 있다"(Excel 창, 다른 PC, OneDrive 동기화)이므로 그걸 앞에 세운다.
    """
    return (
        f"'{workbook.name}'을 {verb} 못했습니다 — 다른 곳에서 잡고 있는 것 같습니다.\n"
        "       1) Excel에서 이 파일을 열어 두었다면 닫고 다시 실행하세요\n"
        "       2) 공유 폴더라 다른 PC에서 열고 있어도 막힙니다\n"
        "       3) OneDrive 동기화 중이면 끝난 뒤 다시 실행하세요\n"
        f"       (원문: {error})")


def _find_open_book(path: Path):
    """이미 Excel에 열려 있으면 그 워크북 반환 (없으면 None)

    같은 파일을 두 번째 인스턴스에서 열면 Excel이 읽기 전용으로 열거나 대화상자를 띄운다.
    사용자가 이 워크북을 하루 종일 열어 두는 것이 정상이라 붙는 경로가 기본이다.
    """
    import xlwings as xw

    target = str(path.resolve()).lower()
    for app in xw.apps:
        for book in app.books:
            try:
                full = str(Path(book.fullname).resolve()).lower()
            except Exception:
                full = ''
            if full == target or book.name.lower() == path.name.lower():
                return book
    return None


def _header_columns(sheet, header_row: int = 1) -> dict[str, int]:
    """헤더명 → 열 번호 (열 순서가 바뀌어도 따라가게)"""
    used = sheet.api.UsedRange
    width = used.Column + used.Columns.Count - 1
    values = sheet.range((header_row, 1), (header_row, width)).value or []
    if not isinstance(values, list):
        values = [values]
    return {str(v).strip(): i + 1 for i, v in enumerate(values) if v is not None}


def _cell_value(line: DnLine, column: str):
    """값 열 하나에 쓸 값 (수식 열은 여기 오지 않는다)"""
    if column == 'DN_ID':
        return line.dn_id
    if column == 'SO_ID':
        return line.so_id
    if column == 'Line item':
        return line.line_item
    if column == 'Qty':
        return line.qty
    if column == 'Currency':
        return line.currency
    if column == '출고일':
        return line.ship_date.to_pydatetime() if hasattr(
            line.ship_date, 'to_pydatetime') else line.ship_date
    if column == '세금계산서 발행일':
        if line.tax_date is None:
            return None
        return line.tax_date.to_pydatetime() if hasattr(
            line.tax_date, 'to_pydatetime') else line.tax_date
    if column == 'Remarks':
        return line.remarks or None
    if column == 'IP 여부':
        return None          # 복사해 온 값을 반드시 지운다
    if column == 'Seq':
        return line.seq
    raise KeyError(column)


def append_lines(
    workbook: Path,
    sheet_name: str,
    lines: list[DnLine],
    *,
    backup_dir: Path | None = None,
) -> WriteResult:
    """`lines`를 시트의 표 끝에 덧붙이고 저장

    Raises:
        RuntimeError: 표를 못 찾거나, 열려 있는 워크북에 저장 안 된 변경이 있을 때
    """
    import xlwings as xw

    if not lines:
        raise ValueError("추가할 행이 없습니다")

    # 파일이 닫혀 있어야 한다. Excel이 워크북을 열면 **읽기조차 막혀서**(실측)
    # 백업 사본을 뜰 수 없다 — 되돌릴 곳 없이 마스터 파일을 건드리게 된다.
    # 그래서 '열려 있으면 붙어서 쓴다'가 아니라 '닫고 오라'가 규약이다.
    _assert_closed_and_writable(workbook)

    backup = (backup_workbook(workbook, backup_dir)
              if backup_dir is not None else None)
    before = workbook.stat()

    book = None
    app = None
    try:
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            book = app.books.open(str(workbook))
        except Exception as e:
            raise RuntimeError(_busy_message(workbook, '열지', e)) from e

        # 읽기 전용이면 **여기서 멈춰야 한다.** 쓰기 자체는 메모리에서 멀쩡히 되고
        # 표 범위도 늘어나지만, `Save()`는 예외도 없이 조용히 무시된다
        # (DisplayAlerts=False면 '다른 이름으로 저장' 대화상자도 안 뜬다).
        # 그대로 두면 "완료 4행 추가"라고 보고하고 파일은 그대로다 — 2026-08-07 실측.
        if book.api.ReadOnly:
            raise RuntimeError(
                f"'{workbook.name}'이 읽기 전용으로 열렸습니다 — 저장이 무시됩니다.\n"
                "       Excel에서 이 파일을 닫고 다시 실행하세요.\n"
                "       (공유 폴더라 다른 PC에서 열고 있어도 읽기 전용이 됩니다)")

        sheet = book.sheets[sheet_name]
        result = _append_to_table(sheet, lines)
        book.app.calculate()
        try:
            book.save()
        except Exception as e:
            raise RuntimeError(_busy_message(workbook, '저장하지', e)) from e

        # 저장이 진짜 파일에 닿았는지 확인 — COM은 실패를 조용히 삼킬 수 있다
        after = workbook.stat()
        if (after.st_mtime_ns, after.st_size) == (before.st_mtime_ns, before.st_size):
            raise RuntimeError(
                f"'{workbook.name}' 저장이 파일에 반영되지 않았습니다 "
                "(파일이 그대로입니다). Excel에서 이 파일을 닫고 다시 실행하세요.")
        logger.debug("저장 완료: %s", workbook.name)
    finally:
        if app is not None:
            try:
                for wb in app.books:
                    wb.close()
            except Exception:
                pass
            try:
                app.quit()
            except Exception:
                pass

    return WriteResult(
        rows_added=result[0], first_row=result[1], last_row=result[2],
        backup=backup, table_ref=result[3],
    )


def _append_to_table(sheet, lines: list[DnLine]) -> tuple[int, int, int, str]:
    """표 끝에 행을 붙이고 값 열을 덮어쓴다 (COM 왕복은 열당 1회)"""
    table = _find_list_object(sheet)
    if table is None:
        raise RuntimeError(f"'{sheet.name}' 시트에서 표(ListObject)를 찾지 못했습니다")

    header_row = table.HeaderRowRange.Row
    first_col = table.Range.Column
    last_col = first_col + table.Range.Columns.Count - 1
    source_row = table.Range.Row + table.Range.Rows.Count - 1
    if source_row <= header_row:
        raise RuntimeError(
            f"'{sheet.name}' 표에 복사할 데이터 행이 없습니다 — "
            "수식을 물려받을 원본 행이 필요합니다")

    n = len(lines)
    start, end = source_row + 1, source_row + n

    # 1) 마지막 행을 새 행들에 타일 복사 — 수식·서식·상대참조가 따라온다
    sheet.range(f'{source_row}:{source_row}').api.Copy(
        sheet.range(f'{start}:{end}').api)

    # 2) 표 범위를 새 행까지 확장 (복사만으로는 표가 늘어나지 않는다)
    sheet.api.ListObjects(table.Name).Resize(
        sheet.range((header_row, first_col), (end, last_col)).api)

    # 3) 값 열만 덮어쓰기
    columns = _header_columns(sheet, header_row)
    missing = [c for c in VALUE_COLUMNS if c not in columns]
    if missing:
        raise RuntimeError(f"'{sheet.name}'에서 열을 찾지 못했습니다: {missing}")
    for column in VALUE_COLUMNS:
        col = columns[column]
        sheet.range((start, col), (end, col)).value = [
            [_cell_value(line, column)] for line in lines]

    # Address는 pywin32 동적 디스패치에서 속성이다 — 인자를 주면 str을 호출하려 든다
    new_ref = str(sheet.api.ListObjects(table.Name).Range.Address).replace('$', '')
    logger.debug("행 추가: %d~%d (%d행), 표 범위 %s", start, end, n, new_ref)
    return n, start, end, new_ref


def _find_list_object(sheet):
    """시트의 표(ListObject) — 여러 개면 가장 큰 것"""
    tables = sheet.api.ListObjects
    if tables.Count == 0:
        return None
    best = None
    for i in range(1, tables.Count + 1):
        table = tables(i)
        if best is None or table.Range.Rows.Count > best.Range.Rows.Count:
            best = table
    return best
