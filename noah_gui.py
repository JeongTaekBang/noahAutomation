"""
NOAH 문서 생성기 GUI
=====================

발주서/거래명세표/PI/FI/OC/CI/PL을 창 하나에서 생성합니다.
사내 배포판(포터블 폴더)의 진입점이며, 개발 PC에서도 그대로 실행됩니다.

설계
----
GUI는 문서를 직접 만들지 않고 기존 CLI(`create_*.py`)를 **자식 프로세스로 실행**하고
표준출력을 화면에 흘립니다. 덕분에

  - 생성 로직에 GUI 코드가 섞이지 않는다 (CLI와 완전히 같은 코드 경로)
  - Excel COM이 자식 프로세스에 격리된다 (COM 오류가 나도 창이 죽지 않음)

Usage:
    python noah_gui.py
"""

from __future__ import annotations

import configparser
import os
import re
import subprocess
import sys
import threading
import time
import zipfile
from collections import deque
from pathlib import Path
from queue import Empty, Queue
from xml.sax.saxutils import unescape

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


APP_DIR = Path(__file__).resolve().parent
INI_FILE = APP_DIR / "noah_config.ini"

DATA_FILE_NAME = "NOAH_SO_PO_DN.xlsx"

# 데이터 파일 검증용 — 이 시트들이 있어야 NOAH 데이터 파일로 인정
REQUIRED_SHEETS = ("SO_국내", "PO_국내", "DN_국내")

# OneDrive 자동 탐색 한도 (회사 PC의 OneDrive 트리는 크고 Files On-Demand까지 겹쳐 느리다)
SEARCH_TIMEOUT_SEC = 3.0
SEARCH_MAX_DEPTH = 5


# === 문서 종류 정의 =========================================================
# key      : 내부 식별자
# label    : 라디오 버튼 표시명
# script   : 실행할 CLI
# id_label : ID 입력란 라벨
# hint     : 입력 예시
# out_attr : config.py의 출력 폴더 상수명 ([출력 폴더 열기]에 사용)
# options  : 이 문서에서만 보이는 옵션 위젯 키
DOC_TYPES: tuple[dict[str, object], ...] = (
    {
        'key': 'po', 'label': '발주서 (PO)', 'script': 'create_po.py',
        'id_label': 'PO_ID', 'hint': '예: ND-0001, NO-0001',
        'out_attr': 'OUTPUT_DIR', 'options': ('force',),
    },
    {
        'key': 'ts', 'label': '거래명세표 (TS)', 'script': 'create_ts.py',
        'id_label': 'DN_ID / 선수금_ID', 'hint': '예: DND-2026-0001, ADV_2026-0001',
        'out_attr': 'TS_OUTPUT_DIR', 'options': ('merge', 'mail'),
    },
    {
        'key': 'pi', 'label': 'Proforma Invoice (PI)', 'script': 'create_pi.py',
        'id_label': 'SO_ID', 'hint': '예: SOO-2026-0001',
        'out_attr': 'PI_OUTPUT_DIR', 'options': (),
    },
    {
        'key': 'fi', 'label': 'Final Invoice (FI)', 'script': 'create_fi.py',
        'id_label': 'DN_ID', 'hint': '예: DNO-2026-0001',
        'out_attr': 'FI_OUTPUT_DIR', 'options': ('fi_mode',),
    },
    {
        'key': 'oc', 'label': 'Order Confirmation (OC)', 'script': 'create_oc.py',
        'id_label': 'SO_ID', 'hint': '예: SOO-2026-0001',
        'out_attr': 'OC_OUTPUT_DIR', 'options': (),
    },
    {
        'key': 'ci', 'label': 'Commercial Invoice (CI)', 'script': 'create_ci.py',
        'id_label': 'DN_ID', 'hint': '예: DNO-2026-0001',
        'out_attr': 'CI_OUTPUT_DIR', 'options': (),
    },
    {
        'key': 'pl', 'label': 'Packing List (PL)', 'script': 'create_pl.py',
        'id_label': 'DN_ID', 'hint': '예: DNO-2026-0001',
        'out_attr': 'PL_OUTPUT_DIR', 'options': (),
    },
)

DOC_BY_KEY: dict[str, dict[str, object]] = {d['key']: d for d in DOC_TYPES}  # type: ignore[index]


# === 설정 파일 ==============================================================

def read_ini() -> dict[str, str]:
    """noah_config.ini의 [paths] 섹션 로드"""
    if not INI_FILE.exists():
        return {}
    # utf-8-sig — 메모장으로 편집하면 BOM이 붙어 섹션을 못 찾는다 (config.py와 동일)
    parser = configparser.ConfigParser(interpolation=None)
    try:
        parser.read(INI_FILE, encoding='utf-8-sig')
    except (configparser.Error, OSError):
        return {}
    if not parser.has_section('paths'):
        return {}
    return {k: v.strip() for k, v in parser.items('paths') if v.strip()}


def write_ini(data_folder: Path) -> None:
    """데이터 폴더를 ini에 기록

    출력·이력도 같은 공유 폴더에 쌓아 여러 사람의 발주 이력이 합쳐지게 한다
    (중복 발주 감지가 전사 기준으로 동작하려면 이력이 한곳에 모여야 한다).
    """
    parser = configparser.ConfigParser(interpolation=None)
    parser['paths'] = {
        'data_folder': str(data_folder),
        'output_base_dir': str(data_folder),
    }
    with INI_FILE.open('w', encoding='utf-8') as f:
        f.write("# NOAH 문서 생성기 설정 — GUI의 [변경] 버튼으로 수정하는 것을 권장합니다.\n")
        parser.write(f)


def effective_data_file() -> Path | None:
    """실제로 사용될 데이터 파일 경로

    config.py를 패키지 임포트 없이 단독 로드해 얻는다. `po_generator`를 정식으로
    임포트하면 __init__.py가 pandas까지 끌고 오므로 GUI 시작이 느려진다.

    개발 PC에는 user_settings.py가 있어 ini보다 우선한다. 화면에 ini 값을 보여주면
    실제 생성에 쓰이는 경로와 달라질 수 있어, 항상 config.py의 결론을 따른다.
    """
    config_path = APP_DIR / "po_generator" / "config.py"
    if not config_path.exists():
        return None
    try:
        import importlib.util
        spec = importlib.util.spec_from_file_location("_noah_config_probe", config_path)
        if spec is None or spec.loader is None:
            return None
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return Path(module.NOAH_SO_PO_DN_FILE)
    except Exception:
        return None


def output_dir_for(doc_key: str) -> Path | None:
    """문서 종류별 출력 폴더 (config.py 기준)"""
    config_path = APP_DIR / "po_generator" / "config.py"
    attr = DOC_BY_KEY[doc_key]['out_attr']
    try:
        import importlib.util
        spec = importlib.util.spec_from_file_location("_noah_config_probe", config_path)
        if spec is None or spec.loader is None:
            return None
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return Path(getattr(module, attr))  # type: ignore[arg-type]
    except Exception:
        return None


# === 데이터 파일 탐색 / 검증 ================================================

def sheet_names(xlsx_path: Path) -> list[str]:
    """xlsx의 시트명 목록 (pandas/openpyxl 없이 zip에서 직접)

    검증에 수 초짜리 pandas 로딩을 물리지 않기 위해 workbook.xml만 읽는다.
    """
    try:
        with zipfile.ZipFile(xlsx_path) as zf:
            xml = zf.read("xl/workbook.xml").decode("utf-8", "replace")
    except (zipfile.BadZipFile, KeyError, OSError):
        return []
    raw = re.findall(r'<sheet[^>]*\bname="([^"]*)"', xml)
    return [unescape(n, {"&quot;": '"', "&apos;": "'"}) for n in raw]


def looks_like_data_file(xlsx_path: Path) -> bool:
    """NOAH 데이터 파일로 보이는지 (필수 시트 존재 여부)"""
    names = set(sheet_names(xlsx_path))
    return all(s in names for s in REQUIRED_SHEETS)


def data_file_locked(xlsx_path: Path) -> bool:
    """다른 프로그램이 파일을 잡고 있는지

    누군가 Excel로 열어두었거나 OneDrive가 동기화 중이면 pandas가
    PermissionError로 죽는다. 5초 기다렸다가 트레이스백을 보여주는 대신
    실행 전에 걸러 사람이 읽을 수 있는 안내를 준다.
    """
    try:
        with xlsx_path.open('rb'):
            return False
    except PermissionError:
        return True
    except OSError:
        return False


LOCK_HINT = (
    "\n[안내] NOAH_SO_PO_DN.xlsx를 다른 프로그램이 사용 중입니다.\n"
    "       Excel에서 파일을 닫거나 OneDrive 동기화가 끝난 뒤 다시 시도하세요.\n"
)


def find_data_file(timeout: float = SEARCH_TIMEOUT_SEC) -> Path | None:
    """OneDrive 폴더에서 NOAH_SO_PO_DN.xlsx 탐색

    시간·깊이 제한이 핵심이다. 회사 PC의 OneDrive 트리를 제한 없이 훑으면
    분 단위로 멈춘다. 못 찾으면 곧바로 파일 선택 대화상자로 넘긴다.
    """
    deadline = time.monotonic() + timeout
    target = DATA_FILE_NAME.lower()

    try:
        home = Path.home()
        roots = [d for d in home.iterdir()
                 if d.is_dir() and d.name.lower().startswith("onedrive")]
    except (PermissionError, OSError):
        return None

    queue: deque[tuple[Path, int]] = deque((r, 0) for r in roots)
    while queue:
        if time.monotonic() > deadline:
            return None
        folder, depth = queue.popleft()
        try:
            for item in folder.iterdir():
                if item.name.lower() == target and item.is_file():
                    return item
                if depth < SEARCH_MAX_DEPTH and item.is_dir():
                    queue.append((item, depth + 1))
        except (PermissionError, OSError):
            continue
    return None


# === 자식 프로세스 ==========================================================

def child_python() -> str:
    """자식 CLI를 실행할 파이썬

    GUI는 pythonw.exe(콘솔 없음)로 뜨지만, 자식은 python.exe로 돌린다.
    """
    exe = Path(sys.executable)
    if exe.name.lower() == "pythonw.exe":
        console = exe.with_name("python.exe")
        if console.exists():
            return str(console)
    return str(exe)


def child_env() -> dict[str, str]:
    """자식 프로세스 환경변수

    UNBUFFERED가 없으면 진행 로그가 끝날 때 한꺼번에 쏟아지고,
    인코딩을 지정하지 않으면 파이프로 넘어온 한글이 깨진다.
    """
    env = os.environ.copy()
    env['PYTHONUNBUFFERED'] = '1'
    env['PYTHONIOENCODING'] = 'utf-8'
    env['PYTHONUTF8'] = '1'
    return env


def build_command(doc_key: str, ids: list[str], options: dict[str, object]) -> list[str]:
    """문서 종류 + 옵션 → CLI 명령"""
    script = str(DOC_BY_KEY[doc_key]['script'])
    cmd = [child_python(), script]

    if doc_key == 'fi' and options.get('fi_mode') == 'po':
        # 발주번호 기준 (복수 DN 통합). ID가 비면 사용 가능한 발주번호 목록 표시
        cmd.append('--po')
        cmd.extend(ids)
        return cmd

    cmd.extend(ids)

    if doc_key == 'po' and options.get('force'):
        cmd.append('--force')

    if doc_key == 'ts':
        if options.get('merge'):
            cmd.append('--merge')
        # 자식은 비대화형이라 메일 프롬프트가 자동으로 꺼지지만, 의도를 명령에 남긴다
        cmd.append('--mail' if options.get('mail') else '--no-mail')

    return cmd


# === 데이터 파일 지정 창 ====================================================

class SetupDialog(tk.Toplevel):
    """데이터 파일(NOAH_SO_PO_DN.xlsx) 지정"""

    def __init__(self, master: tk.Misc, current: Path | None, first_run: bool):
        super().__init__(master)
        self.title("데이터 파일 지정")
        self.resizable(False, False)
        self.result: Path | None = None
        self._search_thread: threading.Thread | None = None

        pad = {'padx': 12, 'pady': 6}

        intro = (
            "NOAH_SO_PO_DN.xlsx 파일의 위치를 지정하세요.\n"
            "OneDrive로 동기화된 로컬 파일이어야 합니다 (웹 링크는 사용할 수 없습니다)."
        )
        ttk.Label(self, text=intro, justify='left').grid(
            row=0, column=0, columnspan=3, sticky='w', **pad)

        self.path_var = tk.StringVar(value=str(current) if current else "")
        entry = ttk.Entry(self, textvariable=self.path_var, width=72)
        entry.grid(row=1, column=0, columnspan=2, sticky='we', padx=(12, 6), pady=6)
        ttk.Button(self, text="파일 찾기...", command=self._browse).grid(
            row=1, column=2, sticky='we', padx=(0, 12), pady=6)

        self.status_var = tk.StringVar(value="")
        ttk.Label(self, textvariable=self.status_var, foreground='#555').grid(
            row=2, column=0, columnspan=3, sticky='w', padx=12)

        btns = ttk.Frame(self)
        btns.grid(row=3, column=0, columnspan=3, sticky='e', **pad)
        self.search_btn = ttk.Button(btns, text="자동으로 찾기", command=self._auto_search)
        self.search_btn.pack(side='left', padx=4)
        ttk.Button(btns, text="확인", command=self._confirm).pack(side='left', padx=4)
        ttk.Button(btns, text="취소", command=self.destroy).pack(side='left', padx=4)

        self.columnconfigure(0, weight=1)
        self.transient(master)
        self.grab_set()
        entry.focus_set()

        if first_run and not self.path_var.get():
            self.after(100, self._auto_search)

    def _browse(self) -> None:
        initial = self.path_var.get().strip()
        initial_dir = str(Path(initial).parent) if initial else str(Path.home())
        chosen = filedialog.askopenfilename(
            parent=self,
            title="NOAH_SO_PO_DN.xlsx 선택",
            initialdir=initial_dir,
            filetypes=[("Excel 파일", "*.xlsx"), ("모든 파일", "*.*")],
        )
        if chosen:
            self.path_var.set(chosen)
            self.status_var.set("")

    def _auto_search(self) -> None:
        if self._search_thread and self._search_thread.is_alive():
            return
        self.search_btn.state(['disabled'])
        self.status_var.set("OneDrive에서 찾는 중...")

        result: list[Path | None] = []

        def work() -> None:
            result.append(find_data_file())

        self._search_thread = threading.Thread(target=work, daemon=True)
        self._search_thread.start()
        self._poll_search(result)

    def _poll_search(self, result: list[Path | None]) -> None:
        if self._search_thread and self._search_thread.is_alive():
            self.after(100, lambda: self._poll_search(result))
            return
        self.search_btn.state(['!disabled'])
        found = result[0] if result else None
        if found:
            self.path_var.set(str(found))
            self.status_var.set("찾았습니다. 경로를 확인하고 [확인]을 누르세요.")
        else:
            self.status_var.set(
                "자동으로 찾지 못했습니다. [파일 찾기...]로 직접 지정하세요.")

    def _confirm(self) -> None:
        raw = self.path_var.get().strip().strip('"')
        if not raw:
            messagebox.showwarning("경로 없음", "파일 경로를 지정하세요.", parent=self)
            return

        # 공유 링크를 붙여넣는 경우가 가장 흔한 실수다. pandas는 URL을 읽지 못한다.
        if raw.lower().startswith(("http://", "https://")):
            messagebox.showerror(
                "웹 링크는 사용할 수 없습니다",
                "웹 링크(https://...)는 사용할 수 없습니다.\n\n"
                "OneDrive 웹에서 해당 파일을 열고 [내 파일에 바로가기 추가]를 누른 뒤,\n"
                "동기화가 끝나면 [파일 찾기...]로 내 PC의 파일을 선택하세요.",
                parent=self,
            )
            return

        path = Path(raw)
        if path.is_dir():
            path = path / DATA_FILE_NAME
        if not path.exists():
            messagebox.showerror(
                "파일 없음", f"파일을 찾을 수 없습니다:\n{path}", parent=self)
            return

        if not looks_like_data_file(path):
            found = ", ".join(sheet_names(path)[:8]) or "(시트를 읽지 못함)"
            proceed = messagebox.askyesno(
                "시트 확인 필요",
                f"필요한 시트({', '.join(REQUIRED_SHEETS)})를 찾지 못했습니다.\n\n"
                f"이 파일의 시트: {found}\n\n"
                "그래도 이 파일을 사용할까요?",
                parent=self,
            )
            if not proceed:
                return

        try:
            write_ini(path.parent)
        except OSError as e:
            messagebox.showerror("설정 저장 실패", f"{e}", parent=self)
            return

        self.result = path
        self.destroy()


# === 메인 창 ================================================================

class MainWindow(ttk.Frame):

    def __init__(self, master: tk.Tk):
        super().__init__(master, padding=10)
        self.master: tk.Tk = master
        self.grid(sticky='nsew')
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)

        self.proc: subprocess.Popen[str] | None = None
        self.queue: Queue[tuple[str, object]] = Queue()
        self.option_vars: dict[str, tk.Variable] = {}
        self.saw_lock_error = False

        self._build_header()
        self._build_doc_types()
        self._build_input()
        self._build_log()

        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)

        self._on_doc_change()
        self._refresh_data_path()

    # --- 화면 구성 ---------------------------------------------------------

    def _build_header(self) -> None:
        frame = ttk.LabelFrame(self, text="데이터 파일", padding=8)
        frame.grid(row=0, column=0, sticky='we', pady=(0, 8))
        frame.columnconfigure(0, weight=1)

        self.data_path_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=self.data_path_var, foreground='#333').grid(
            row=0, column=0, sticky='w')
        ttk.Button(frame, text="변경...", command=self._change_data_file).grid(
            row=0, column=1, sticky='e', padx=(8, 0))

    def _build_doc_types(self) -> None:
        frame = ttk.LabelFrame(self, text="문서 종류", padding=8)
        frame.grid(row=1, column=0, sticky='we', pady=(0, 8))

        self.doc_var = tk.StringVar(value='po')
        for i, doc in enumerate(DOC_TYPES):
            ttk.Radiobutton(
                frame, text=str(doc['label']), value=str(doc['key']),
                variable=self.doc_var, command=self._on_doc_change,
            ).grid(row=i % 4, column=i // 4, sticky='w', padx=(0, 24), pady=2)

    def _build_input(self) -> None:
        frame = ttk.LabelFrame(self, text="입력", padding=8)
        frame.grid(row=2, column=0, sticky='we', pady=(0, 8))
        frame.columnconfigure(0, weight=1)

        self.id_label_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=self.id_label_var).grid(row=0, column=0, sticky='w')

        self.id_text = tk.Text(frame, height=5, width=60, font=('맑은 고딕', 10))
        self.id_text.grid(row=1, column=0, sticky='we', pady=(2, 6))

        self.options_frame = ttk.Frame(frame)
        self.options_frame.grid(row=2, column=0, sticky='w')

        btns = ttk.Frame(frame)
        btns.grid(row=3, column=0, sticky='e', pady=(6, 0))
        self.run_btn = ttk.Button(btns, text="생성", command=self._run)
        self.run_btn.pack(side='left', padx=4)
        self.stop_btn = ttk.Button(btns, text="중지", command=self._stop, state='disabled')
        self.stop_btn.pack(side='left', padx=4)
        ttk.Button(btns, text="출력 폴더 열기", command=self._open_output).pack(
            side='left', padx=4)

    def _build_log(self) -> None:
        frame = ttk.LabelFrame(self, text="진행 상황", padding=8)
        frame.grid(row=3, column=0, sticky='nsew')
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self.log = tk.Text(frame, height=14, width=90, font=('맑은 고딕', 9),
                           state='disabled', wrap='word', background='#fbfbfb')
        self.log.grid(row=0, column=0, sticky='nsew')
        scroll = ttk.Scrollbar(frame, orient='vertical', command=self.log.yview)
        scroll.grid(row=0, column=1, sticky='ns')
        self.log.configure(yscrollcommand=scroll.set)

        self.log.tag_configure('error', foreground='#c00000')
        self.log.tag_configure('warn', foreground='#b35c00')
        self.log.tag_configure('meta', foreground='#666')

    # --- 상태 갱신 ---------------------------------------------------------

    def _refresh_data_path(self) -> None:
        path = effective_data_file()
        if path and path.exists():
            self.data_path_var.set(str(path))
        elif path:
            self.data_path_var.set(f"{path}  ← 파일 없음")
        else:
            self.data_path_var.set("(지정되지 않음)")

    def _on_doc_change(self) -> None:
        doc = DOC_BY_KEY[self.doc_var.get()]
        self.id_label_var.set(f"{doc['id_label']}   ({doc['hint']}, 여러 건은 줄바꿈)")

        for child in self.options_frame.winfo_children():
            child.destroy()
        self.option_vars.clear()

        for opt in doc['options']:  # type: ignore[union-attr]
            if opt == 'force':
                var = tk.BooleanVar(value=False)
                ttk.Checkbutton(self.options_frame, variable=var,
                                text="검증 오류 무시하고 생성 (--force)").pack(anchor='w')
                self.option_vars['force'] = var
            elif opt == 'merge':
                var = tk.BooleanVar(value=False)
                ttk.Checkbutton(self.options_frame, variable=var,
                                text="월합 — 여러 DN을 한 장으로 (--merge)").pack(anchor='w')
                self.option_vars['merge'] = var
            elif opt == 'mail':
                var = tk.BooleanVar(value=False)
                ttk.Checkbutton(self.options_frame, variable=var,
                                text="메일 초안 만들기 (--mail)").pack(anchor='w')
                self.option_vars['mail'] = var
            elif opt == 'fi_mode':
                var = tk.StringVar(value='dn')
                ttk.Radiobutton(self.options_frame, variable=var, value='dn',
                                text="DN_ID 기준", command=self._on_fi_mode).pack(anchor='w')
                ttk.Radiobutton(
                    self.options_frame, variable=var, value='po',
                    text="발주번호 기준 (복수 DN 통합) — 비워두면 발주번호 목록 표시",
                    command=self._on_fi_mode).pack(anchor='w')
                self.option_vars['fi_mode'] = var

    def _on_fi_mode(self) -> None:
        mode = self.option_vars.get('fi_mode')
        if mode is None:
            return
        doc = DOC_BY_KEY['fi']
        if mode.get() == 'po':
            self.id_label_var.set("발주번호 (Customer PO)   (예: 26KPO00144, 여러 건은 줄바꿈)")
        else:
            self.id_label_var.set(f"{doc['id_label']}   ({doc['hint']}, 여러 건은 줄바꿈)")

    # --- 동작 --------------------------------------------------------------

    def _change_data_file(self) -> None:
        current = effective_data_file()
        dialog = SetupDialog(self.master, current, first_run=False)
        self.master.wait_window(dialog)
        if dialog.result:
            self._refresh_data_path()
            self._append(f"데이터 파일 변경: {dialog.result}\n", 'meta')

    def ensure_data_file(self) -> None:
        """첫 실행 — 데이터 파일이 없으면 지정 창을 띄운다"""
        path = effective_data_file()
        if path and path.exists():
            return
        dialog = SetupDialog(self.master, path, first_run=True)
        self.master.wait_window(dialog)
        self._refresh_data_path()

    def _collect_ids(self) -> list[str]:
        raw = self.id_text.get('1.0', 'end')
        return [line.strip() for line in raw.splitlines() if line.strip()]

    def _run(self) -> None:
        if self.proc is not None:
            return

        path = effective_data_file()
        if not path or not path.exists():
            messagebox.showerror(
                "데이터 파일 없음",
                "NOAH_SO_PO_DN.xlsx를 먼저 지정하세요.\n[변경...] 버튼을 누르세요.")
            return

        if data_file_locked(path):
            messagebox.showwarning(
                "파일 사용 중",
                "NOAH_SO_PO_DN.xlsx를 다른 프로그램이 사용 중입니다.\n\n"
                "Excel에서 파일을 닫거나 OneDrive 동기화가 끝난 뒤\n"
                "다시 시도하세요.")
            return

        doc_key = self.doc_var.get()
        ids = self._collect_ids()
        options = {k: v.get() for k, v in self.option_vars.items()}

        list_mode = doc_key == 'fi' and options.get('fi_mode') == 'po' and not ids
        if not ids and not list_mode:
            messagebox.showwarning("입력 없음", "ID를 한 줄에 하나씩 입력하세요.")
            return

        cmd = build_command(doc_key, ids, options)

        self._clear_log()
        self._append(f"$ {' '.join(cmd[1:])}\n\n", 'meta')

        try:
            self.proc = subprocess.Popen(
                cmd,
                cwd=str(APP_DIR),
                env=child_env(),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                stdin=subprocess.DEVNULL,
                text=True,
                encoding='utf-8',
                errors='replace',
                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0),
            )
        except OSError as e:
            self._append(f"[오류] 실행 실패: {e}\n", 'error')
            self.proc = None
            return

        self.run_btn.state(['disabled'])
        self.stop_btn.state(['!disabled'])

        threading.Thread(target=self._pump, args=(self.proc,), daemon=True).start()
        self.after(50, self._drain)

    def _pump(self, proc: subprocess.Popen[str]) -> None:
        """자식 stdout을 큐로 옮긴다 (워커 스레드)"""
        if proc.stdout is not None:
            for line in proc.stdout:
                self.queue.put(('log', line.rstrip('\n')))
        proc.wait()
        self.queue.put(('done', proc.returncode))

    def _drain(self) -> None:
        """큐 → 로그 위젯 (GUI 스레드)"""
        finished = False
        code = 0
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == 'log':
                    line = str(payload)
                    if 'PermissionError' in line:
                        # 실행 직전 검사를 통과했더라도 그 사이에 잠길 수 있다
                        self.saw_lock_error = True
                    tag = ('error' if ('[오류]' in line or 'Error' in line)
                           else 'warn' if ('[경고]' in line or '[주의]' in line)
                           else None)
                    self._append(line + '\n', tag)
                else:
                    finished = True
                    code = int(payload)  # type: ignore[arg-type]
        except Empty:
            pass

        if finished:
            self._on_finished(code)
        elif self.proc is not None:
            self.after(50, self._drain)

    def _on_finished(self, code: int) -> None:
        self.proc = None
        self.run_btn.state(['!disabled'])
        self.stop_btn.state(['disabled'])
        if code == 0:
            self._append("\n완료되었습니다.\n", 'meta')
        else:
            self._append(f"\n종료 코드 {code} — 위 메시지를 확인하세요.\n", 'error')
        if self.saw_lock_error:
            self._append(LOCK_HINT, 'warn')
            self.saw_lock_error = False

    def _stop(self) -> None:
        if self.proc is None:
            return
        confirm = messagebox.askyesno(
            "생성 중지",
            "생성을 중지할까요?\n\n"
            "작업 중이던 Excel이 백그라운드에 남을 수 있습니다.")
        if not confirm:
            return
        try:
            self.proc.terminate()
        except OSError:
            pass
        self._append("\n[주의] 사용자가 중지했습니다.\n", 'warn')

    def _open_output(self) -> None:
        out = output_dir_for(self.doc_var.get())
        if out is None:
            messagebox.showerror("오류", "출력 폴더를 알 수 없습니다.")
            return
        if not out.exists():
            messagebox.showinfo(
                "폴더 없음",
                f"아직 생성된 문서가 없습니다.\n\n{out}")
            return
        os.startfile(str(out))  # noqa: S606 — Windows 전용 도구

    # --- 로그 --------------------------------------------------------------

    def _append(self, text: str, tag: str | None = None) -> None:
        self.log.configure(state='normal')
        self.log.insert('end', text, tag or '')
        self.log.see('end')
        self.log.configure(state='disabled')

    def _clear_log(self) -> None:
        self.log.configure(state='normal')
        self.log.delete('1.0', 'end')
        self.log.configure(state='disabled')


def main() -> int:
    root = tk.Tk()
    root.title("NOAH 문서 생성기")
    root.geometry("760x720")
    try:
        root.call('tk', 'scaling', 1.2)
    except tk.TclError:
        pass

    window = MainWindow(root)
    root.after(200, window.ensure_data_file)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
