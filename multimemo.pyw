"""멀티 메모 - 한 화면에 여러 개의 스티커 메모를 동시에 띄우는 프로그램.

- 엑셀 시트처럼 아래쪽 탭마다 별도의 메모판 (탭 추가/이름 변경/삭제/순서 이동)
- 메모 개수에 맞춰 화면을 자동으로 나눠 배치 (1개=전체, 2개=반반, 4개=2x2 ...)
- 각 메모는 <저장 폴더>/<탭 이름>/ 에 개별 .txt 파일로 자동 저장
  (윈도우: 프로그램 옆 notes 폴더, 맥: 문서/멀티메모)
- 탭 저장: 탭 하나를 .zip 파일로 저장하고, 다른 곳에서 탭으로 불러오기
- 단축키: Ctrl(맥 ⌘)+N 새 메모, Ctrl(⌘)+T 새 탭, Ctrl+Tab 다음 탭,
          Ctrl(⌘)+마우스휠 글자 크기 (5~20pt)
- 윈도우/맥 공용 코드
"""
import json
import math
import os
import re
import shutil
import subprocess
import sys
import time
import zipfile
import ctypes
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, simpledialog

IS_MAC = sys.platform == "darwin"
if IS_MAC:
    # 앱(.app)으로 묶어도 메모가 남도록 '문서' 폴더에 저장
    NOTES_DIR = os.path.join(os.path.expanduser("~"), "Documents", "멀티메모")
else:
    NOTES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "notes")
TRASH_DIR = os.path.join(NOTES_DIR, "_삭제됨")
META_FILE = os.path.join(NOTES_DIR, "_meta.json")
SHEET_META = "_meta.json"

# (배경색, 포인트색) - 아주 연한 파스텔 톤
COLORS = [
    ("#FFFCE3", "#F2E6A0"),  # 노랑
    ("#EEF9E8", "#C9E6B8"),  # 초록
    ("#FFF0F5", "#F2C9D8"),  # 분홍
    ("#F4F0FF", "#D6CCF2"),  # 보라
    ("#EDF6FF", "#C3DCF2"),  # 파랑
    ("#FAFAFA", "#DADADA"),  # 회색
]
BOARD_BG = "#EDEEF0"
TAB_BG = "#E2E4E8"
RADIUS = 16   # 모서리 둥근 정도
INSET = 7     # 둥근 모서리 안쪽으로 내용 들여쓰기
UI_SCALE = 0.9           # 윈도우 글자 배율 (작게 하려면 낮추기)
MAC_SCALE = 1.4          # 맥 글자 배율 (맥 Tk는 1pt=1px이라 그대로 두면 너무 작음)
FONT_MIN, FONT_MAX = 5, 20
if IS_MAC:
    FONT_FAMILY = "Apple SD Gothic Neo"
    # 굵기: 이름 -> (후보 글꼴들, weight). 설치된 첫 번째 글꼴을 사용
    WEIGHT_SPECS = {
        "얇게": (["Apple SD Gothic Neo Light", "AppleSDGothicNeo-Light", "Apple SD Gothic Neo"], "normal"),
        "보통": (["Apple SD Gothic Neo"], "normal"),
        "굵게": (["Apple SD Gothic Neo"], "bold"),
    }
    MOD, MOD_LABEL = "Command", "⌘"
else:
    FONT_FAMILY = "맑은 고딕"
    WEIGHT_SPECS = {
        "얇게": (["맑은 고딕 Semilight"], "normal"),
        "보통": (["맑은 고딕"], "normal"),
        "굵게": (["맑은 고딕"], "bold"),
    }
    MOD, MOD_LABEL = "Control", "Ctrl+"
WEIGHTS = {}  # Tk 시작 후 resolve_fonts()가 채움
SAVE_DELAY_MS = 500


def read_json(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return default


def write_json(path, data):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def clean_name(name):
    """탭 이름을 폴더 이름으로 쓸 수 있게 정리."""
    name = re.sub(r'[\\/:*?"<>|]', "_", name or "").strip().strip(".")
    return name.lstrip("_")[:40]


def resolve_fonts():
    installed = set(tkfont.families())
    for name, (cands, wt) in WEIGHT_SPECS.items():
        WEIGHTS[name] = (next((c for c in cands if c in installed), FONT_FAMILY), wt)


def open_folder(path):
    if IS_MAC:
        subprocess.Popen(["open", path])
    else:
        os.startfile(path)


def rounded_rect(canvas, x1, y1, x2, y2, r, **kw):
    pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
           x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return canvas.create_polygon(pts, smooth=True, **kw)


class Note:
    def __init__(self, sheet, note_id, title="", color=0):
        self.sheet = sheet
        self.app = app = sheet.app
        self.id = note_id
        self.color = color
        self._save_job = None

        bg, accent = COLORS[color]
        # 둥근 배경은 Canvas에 그리고, 실제 내용은 그 안쪽에 배치
        self.frame = tk.Canvas(sheet.board, bg=BOARD_BG, highlightthickness=0, bd=0,
                               width=50, height=50)
        self.frame.bind("<Configure>", self._redraw)

        self.inner = tk.Frame(self.frame, bg=bg)
        self.inner.place(x=INSET, y=INSET, relwidth=1, relheight=1,
                         width=-2 * INSET, height=-2 * INSET - 2)

        self.header = tk.Frame(self.inner, bg=bg)
        self.header.pack(fill="x")

        self.title_var = tk.StringVar(value=title)
        self.title_entry = tk.Entry(self.header, textvariable=self.title_var, bd=0,
                                    bg=bg, fg="#444", font=(FONT_FAMILY, 10, "bold"),
                                    relief="flat", highlightthickness=0)
        self.title_entry.pack(side="left", fill="x", expand=True, padx=(6, 4), pady=4)
        self.title_var.trace_add("write", lambda *_: self.schedule_save())

        self.btn_del = self._header_button("✕", self.delete, "메모 삭제")
        self.btn_color = self._header_button("●", self.next_color, "색 바꾸기")
        self.btn_color.config(fg=accent)
        self.btn_new = self._header_button("＋", sheet.add_note, "새 메모")

        self.sep = tk.Frame(self.inner, bg=accent, height=1)
        self.sep.pack(fill="x", padx=4)

        self.text = tk.Text(self.inner, wrap="word", bd=0, bg=bg, padx=8, pady=8,
                            font=app.body_font, undo=True,
                            relief="flat", insertbackground="#333", fg="#333",
                            highlightthickness=0)
        self.text.pack(fill="both", expand=True)
        self.text.bind("<<Modified>>", self._on_modified)
        self.text.bind(f"<{MOD}-MouseWheel>", app.on_zoom)

        self.status = tk.Label(self.inner, text="", bg=bg, fg="#AAA",
                               font=(FONT_FAMILY, 8), anchor="e")
        self.status.pack(fill="x")

    def _redraw(self, _event=None):
        c = self.frame
        c.delete("bg")
        w, h = c.winfo_width(), c.winfo_height()
        bg, accent = COLORS[self.color]
        # 살짝 그림자 + 둥근 메모 본체
        rounded_rect(c, 2, 3, w - 1, h - 1, RADIUS, fill="#DCDDE0", outline="", tags="bg")
        rounded_rect(c, 0, 0, w - 3, h - 4, RADIUS, fill=bg, outline=accent, tags="bg")
        c.tag_lower("bg")

    def _header_button(self, label, cmd, tip):
        bg, _ = COLORS[self.color]
        b = tk.Label(self.header, text=label, bg=bg, fg="#999", cursor="hand2",
                     font=(FONT_FAMILY, 10), padx=5)
        b.pack(side="right")
        normal = lambda: COLORS[self.color][1] if b is getattr(self, "btn_color", None) else "#999"
        b.bind("<Button-1>", lambda e: cmd())
        b.bind("<Enter>", lambda e: (b.config(fg="#333"), self.app.show_tip(tip)))
        b.bind("<Leave>", lambda e: (b.config(fg=normal()), self.app.show_tip("")))
        return b

    @property
    def path(self):
        return os.path.join(self.sheet.folder, f"{self.id}.txt")

    def load_text(self):
        if os.path.exists(self.path):
            with open(self.path, encoding="utf-8") as f:
                self.text.insert("1.0", f.read())
        self.text.edit_modified(False)
        self.text.edit_reset()

    def _on_modified(self, _event):
        if self.text.edit_modified():
            self.text.edit_modified(False)
            self.schedule_save()

    def schedule_save(self):
        if self._save_job:
            self.app.root.after_cancel(self._save_job)
        self.status.config(text="입력 중…")
        self._save_job = self.app.root.after(SAVE_DELAY_MS, self.save)

    def save(self):
        self._save_job = None
        content = self.text.get("1.0", "end-1c")
        with open(self.path, "w", encoding="utf-8") as f:
            f.write(content)
        self.sheet.save_meta()
        self.status.config(text=f"저장됨 {time.strftime('%H:%M:%S')}")

    def flush(self):
        if self._save_job:
            self.app.root.after_cancel(self._save_job)
            self.save()

    def next_color(self):
        self.color = (self.color + 1) % len(COLORS)
        bg, accent = COLORS[self.color]
        for w in (self.inner, self.header, self.title_entry, self.btn_del,
                  self.btn_color, self.btn_new, self.text, self.status):
            w.config(bg=bg)
        self.btn_color.config(fg=accent)
        self.sep.config(bg=accent)
        self._redraw()
        self.sheet.save_meta()

    def delete(self):
        content = self.text.get("1.0", "end-1c").strip()
        name = self.title_var.get() or "이 메모"
        if content and not messagebox.askyesno("메모 삭제", f"'{name}'를 삭제할까요?\n(_삭제됨 폴더로 옮겨집니다)"):
            return
        self.flush()
        if os.path.exists(self.path):
            os.makedirs(TRASH_DIR, exist_ok=True)
            stamp = time.strftime("%Y%m%d_%H%M%S")
            shutil.move(self.path, os.path.join(TRASH_DIR, f"{self.sheet.name}_{self.id}_{stamp}.txt"))
        self.sheet.remove_note(self)


class Sheet:
    """탭 하나 = 메모판 하나 = notes/<탭 이름>/ 폴더 하나."""

    def __init__(self, app, name):
        self.app = app
        self.name = name
        self.notes = []
        self.board = tk.Frame(app.container, bg=BOARD_BG)
        self._last_size = None
        self.board.bind("<Configure>", self._on_resize)

    @property
    def folder(self):
        return os.path.join(NOTES_DIR, self.name)

    def load(self):
        os.makedirs(self.folder, exist_ok=True)
        meta = read_json(os.path.join(self.folder, SHEET_META), {})
        notes_meta = meta.get("notes", {})
        order = [i for i in meta.get("order", [])
                 if os.path.exists(os.path.join(self.folder, f"{i}.txt"))]
        # meta에 없지만 폴더에 있는 txt 파일도 불러오기
        for fn in sorted(os.listdir(self.folder)):
            nid, ext = os.path.splitext(fn)
            if ext == ".txt" and nid not in order:
                order.append(nid)
        for nid in order:
            m = notes_meta.get(nid, {})
            note = Note(self, nid, m.get("title", nid), m.get("color", 0) % len(COLORS))
            note.load_text()
            self.notes.append(note)
        if not self.notes:
            self.add_note(focus=False)
        else:
            self.relayout()

    def save_meta(self):
        if not os.path.isdir(self.folder):
            return
        write_json(os.path.join(self.folder, SHEET_META), {
            "order": [n.id for n in self.notes],
            "notes": {n.id: {"title": n.title_var.get(), "color": n.color} for n in self.notes},
        })

    def flush(self):
        for n in self.notes:
            n.flush()

    def add_note(self, focus=True):
        nid = time.strftime("memo_%Y%m%d_%H%M%S")
        existing = {n.id for n in self.notes}
        base, k = nid, 1
        while nid in existing or os.path.exists(os.path.join(self.folder, f"{nid}.txt")):
            nid, k = f"{base}_{k}", k + 1
        color = len(self.notes) % len(COLORS)
        note = Note(self, nid, f"메모 {len(self.notes) + 1}", color)
        self.notes.append(note)
        note.save()
        self.relayout()
        if focus:
            note.text.focus_set()
        return "break"

    def remove_note(self, note):
        self.notes.remove(note)
        note.frame.destroy()
        self.save_meta()
        if not self.notes:
            self.add_note()
        else:
            self.relayout()

    def show(self):
        self.board.pack(fill="both", expand=True, padx=6, pady=(6, 0))
        self.relayout()

    def hide(self):
        self.board.pack_forget()

    def relayout(self):
        """메모 개수에 맞춰 격자를 다시 계산해서 화면을 균등 분할."""
        n = len(self.notes)
        if not n:
            return
        w, h = self.board.winfo_width(), self.board.winfo_height()
        # 화면이 가로로 넓으면 열을 더 많이 (메모 한 칸이 살짝 가로로 긴 비율이 되도록)
        ratio = w / h if w > 1 and h > 1 else 1.5
        cols = min(n, max(1, round(math.sqrt(n * ratio / 1.2))))
        rows = math.ceil(n / cols)

        for c in range(12):
            self.board.columnconfigure(c, weight=0, uniform="")
        for r in range(12):
            self.board.rowconfigure(r, weight=0, uniform="")
        for c in range(cols):
            self.board.columnconfigure(c, weight=1, uniform="col")
        for r in range(rows):
            self.board.rowconfigure(r, weight=1, uniform="row")

        for i, note in enumerate(self.notes):
            r, c = divmod(i, cols)
            # 마지막 줄이 비면 남은 메모가 가로로 늘어나도록
            span = cols - c if (r == rows - 1 and i == n - 1) else 1
            note.frame.grid(row=r, column=c, columnspan=span, sticky="nsew", padx=5, pady=5)
        if self.app.current is self:
            self.app.update_count()

    def _on_resize(self, event):
        size = (event.width, event.height)
        if size != self._last_size:
            self._last_size = size
            self.relayout()


class App:
    def __init__(self):
        os.makedirs(NOTES_DIR, exist_ok=True)
        meta = self._migrate(read_json(META_FILE, {}))
        self.font_size = max(FONT_MIN, min(FONT_MAX, meta.get("font_size", 11)))
        self.font_weight = meta.get("font_weight", "보통")
        if self.font_weight not in WEIGHT_SPECS:
            self.font_weight = "보통"

        # 고해상도(DPI) 인식: 윈도우가 흐리게 확대하지 않고 선명하게 그리도록
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except (AttributeError, OSError):
            pass
        self.root = tk.Tk()
        if IS_MAC:
            self.root.tk.call("tk", "scaling", MAC_SCALE)
        else:
            dpi = self.root.winfo_fpixels("1i")
            self.root.tk.call("tk", "scaling", dpi / 72 * UI_SCALE)
        resolve_fonts()
        family, weight = WEIGHTS[self.font_weight]
        self.body_font = tkfont.Font(family=family, size=self.font_size, weight=weight)
        self.root.title("멀티 메모")
        self.root.geometry(meta.get("geometry", "1200x800"))
        self.root.configure(bg=BOARD_BG)
        self.root.minsize(400, 300)

        self._build_toolbar()
        self.container = tk.Frame(self.root, bg=BOARD_BG)
        self.tabbar = tk.Frame(self.root, bg=TAB_BG)
        self.tabbar.pack(side="bottom", fill="x")
        self.container.pack(fill="both", expand=True)

        # 탭(시트) 불러오기: meta 순서 + 폴더에만 있는 탭
        names = [n for n in meta.get("tabs", []) if os.path.isdir(os.path.join(NOTES_DIR, n))]
        for fn in sorted(os.listdir(NOTES_DIR)):
            if (os.path.isdir(os.path.join(NOTES_DIR, fn)) and not fn.startswith("_")
                    and fn not in names):
                names.append(fn)
        if not names:
            names = ["시트1"]
        self.current = None
        self.sheets = []
        for name in names:
            sheet = Sheet(self, name)
            sheet.load()
            self.sheets.append(sheet)
        active = meta.get("active", 0)
        self.switch_to(self.sheets[active if 0 <= active < len(self.sheets) else 0])

        self.root.bind_all(f"<{MOD}-n>", lambda e: self.current.add_note())
        self.root.bind_all(f"<{MOD}-N>", lambda e: self.current.add_note())
        self.root.bind_all(f"<{MOD}-t>", lambda e: self.add_sheet())
        self.root.bind_all(f"<{MOD}-T>", lambda e: self.add_sheet())
        self.root.bind_all("<Control-Tab>", lambda e: self._step_tab(1))
        self.root.bind_all("<Control-Shift-Tab>", lambda e: self._step_tab(-1))
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------- 예전 버전(탭 없음) 데이터 옮기기 ----------
    def _migrate(self, meta):
        loose = [f for f in os.listdir(NOTES_DIR)
                 if f.endswith(".txt") and os.path.isfile(os.path.join(NOTES_DIR, f))]
        if "tabs" in meta and not loose:
            return meta
        if loose or "order" in meta:
            dest = os.path.join(NOTES_DIR, "시트1")
            os.makedirs(dest, exist_ok=True)
            for f in loose:
                if not os.path.exists(os.path.join(dest, f)):
                    shutil.move(os.path.join(NOTES_DIR, f), os.path.join(dest, f))
            old = read_json(os.path.join(dest, SHEET_META), {"order": [], "notes": {}})
            old["order"] += [i for i in meta.get("order", []) if i not in old["order"]]
            old["notes"].update(meta.get("notes", {}))
            write_json(os.path.join(dest, SHEET_META), old)
            meta.pop("order", None)
            meta.pop("notes", None)
            meta.setdefault("tabs", [])
            if "시트1" not in meta["tabs"]:
                meta["tabs"].insert(0, "시트1")
            write_json(META_FILE, meta)
        return meta

    # ---------- 툴바 ----------
    def _build_toolbar(self):
        bar = tk.Frame(self.root, bg=BOARD_BG)
        bar.pack(fill="x", padx=8, pady=(6, 0))
        self._small_button(bar, f"＋ 새 메모 ({MOD_LABEL}N)", lambda: self.current.add_note(),
                           side="left", padx=10, pady=3)
        self._small_button(bar, "탭 저장", self.export_sheet, side="left", padx=10, pady=3,
                           tip="현재 탭을 .zip 파일로 저장").pack_configure(padx=(6, 1))
        self._small_button(bar, "탭 불러오기", self.import_sheet, side="left", padx=10, pady=3,
                           tip="저장해 둔 탭(.zip)을 새 탭으로 열기")
        self._small_button(bar, "저장 폴더 열기", lambda: open_folder(self.current.folder),
                           side="left", padx=10, pady=3).pack_configure(padx=(6, 1))

        # 글자 크기 / 굵기 조절
        self.weight_btns = {}
        for name in reversed(list(WEIGHTS)):
            fam, wt = WEIGHTS[name]
            b = tk.Label(bar, text=name, font=(fam, 9, wt), padx=8, pady=2, cursor="hand2")
            b.pack(side="right", padx=(0, 1))
            b.bind("<Button-1>", lambda e, n=name: self.set_font_weight(n))
            self.weight_btns[name] = b
        tk.Label(bar, text="굵기", bg=BOARD_BG, fg="#666",
                 font=(FONT_FAMILY, 9)).pack(side="right", padx=(14, 4))
        self._small_button(bar, "＋", lambda: self.set_font_size(self.font_size + 1))
        self.size_label = tk.Label(bar, text="", bg="#FFFFFF", fg="#333", width=5,
                                   font=(FONT_FAMILY, 9), pady=2)
        self.size_label.pack(side="right")
        self._small_button(bar, "－", lambda: self.set_font_size(self.font_size - 1))
        tk.Label(bar, text="글자 크기", bg=BOARD_BG, fg="#666",
                 font=(FONT_FAMILY, 9)).pack(side="right", padx=(14, 4))
        self._refresh_font_controls()

        self.tip = tk.Label(bar, text="", bg=BOARD_BG, fg="#777", font=(FONT_FAMILY, 9))
        self.tip.pack(side="left", padx=8)
        self.count_label = tk.Label(bar, text="", bg=BOARD_BG, fg="#555",
                                    font=(FONT_FAMILY, 9))
        self.count_label.pack(side="right", padx=(0, 6))

    def _small_button(self, parent, text, cmd, side="right", padx=7, pady=2, tip=""):
        b = tk.Label(parent, text=text, bg="#FFFFFF", fg="#555", font=(FONT_FAMILY, 9),
                     padx=padx, pady=pady, cursor="hand2")
        b.pack(side=side, padx=1)
        b.bind("<Button-1>", lambda e: cmd())
        b.bind("<Enter>", lambda e: (b.config(bg="#E6E9EE"), self.show_tip(tip)))
        b.bind("<Leave>", lambda e: (b.config(bg="#FFFFFF"), self.show_tip("")))
        return b

    def show_tip(self, text):
        self.tip.config(text=text)

    def update_count(self):
        self.count_label.config(text=f"메모 {len(self.current.notes)}개")

    # ---------- 탭(시트) ----------
    def render_tabs(self):
        for w in self.tabbar.winfo_children():
            w.destroy()
        tk.Frame(self.tabbar, bg="#CCCED3", height=1).pack(side="top", fill="x")
        row = tk.Frame(self.tabbar, bg=TAB_BG)
        row.pack(side="left", padx=(8, 0))
        for sheet in self.sheets:
            on = sheet is self.current
            t = tk.Label(row, text=sheet.name, cursor="hand2", padx=14, pady=4,
                         bg="#FFFFFF" if on else TAB_BG, fg="#1A7F37" if on else "#555",
                         font=(FONT_FAMILY, 9, "bold" if on else "normal"))
            t.pack(side="left", padx=(0, 1), pady=(0, 3))
            t.bind("<Button-1>", lambda e, s=sheet: self.switch_to(s))
            t.bind("<Double-Button-1>", lambda e, s=sheet: self.rename_sheet(s))
            # 우클릭 메뉴 (맥은 Button-2 / control+클릭)
            for ev in ("<Button-2>", "<Control-Button-1>") if IS_MAC else ("<Button-3>",):
                t.bind(ev, lambda e, s=sheet: self._tab_menu(e, s))
            if on:
                # 엑셀처럼 선택된 탭 아래에 초록 줄
                tk.Frame(t, bg="#1A7F37", height=2).place(relx=0, rely=1, relwidth=1, y=-2)
        plus = tk.Label(row, text="＋", bg=TAB_BG, fg="#666", cursor="hand2",
                        font=(FONT_FAMILY, 11), padx=10)
        plus.pack(side="left", pady=(0, 3))
        plus.bind("<Button-1>", lambda e: self.add_sheet())
        tk.Label(self.tabbar, text="더블클릭: 이름 변경 · 우클릭: 메뉴", bg=TAB_BG, fg="#999",
                 font=(FONT_FAMILY, 8)).pack(side="right", padx=10)

    def _tab_menu(self, event, sheet):
        self.switch_to(sheet)
        m = tk.Menu(self.root, tearoff=0, font=(FONT_FAMILY, 9))
        m.add_command(label="이름 바꾸기", command=lambda: self.rename_sheet(sheet))
        m.add_command(label="탭 저장 (.zip)", command=self.export_sheet)
        m.add_separator()
        m.add_command(label="◀ 왼쪽으로 이동", command=lambda: self.move_sheet(sheet, -1))
        m.add_command(label="오른쪽으로 이동 ▶", command=lambda: self.move_sheet(sheet, 1))
        m.add_separator()
        m.add_command(label="탭 삭제", command=lambda: self.delete_sheet(sheet))
        m.tk_popup(event.x_root, event.y_root)

    def switch_to(self, sheet):
        if self.current is sheet:
            return
        if self.current:
            self.current.flush()
            self.current.hide()
        self.current = sheet
        sheet.show()
        self.update_count()
        self.render_tabs()
        self.root.title(f"멀티 메모 - {sheet.name}")
        self.save_meta()

    def _step_tab(self, d):
        i = self.sheets.index(self.current)
        self.switch_to(self.sheets[(i + d) % len(self.sheets)])
        return "break"

    def _unique_name(self, base):
        base = clean_name(base) or "시트"
        taken = {s.name.lower() for s in self.sheets}
        if base.lower() not in taken and not os.path.exists(os.path.join(NOTES_DIR, base)):
            return base
        k = 2
        while (f"{base}{k}".lower() in taken or
               os.path.exists(os.path.join(NOTES_DIR, f"{base}{k}"))):
            k += 1
        return f"{base}{k}"

    def add_sheet(self, name=None):
        sheet = Sheet(self, self._unique_name(name or f"시트{len(self.sheets) + 1}"))
        sheet.load()
        self.sheets.append(sheet)
        self.switch_to(sheet)
        return sheet

    def rename_sheet(self, sheet):
        new = simpledialog.askstring("탭 이름 바꾸기", "새 이름:", initialvalue=sheet.name,
                                     parent=self.root)
        if new is None:
            return
        new = clean_name(new)
        if not new or new == sheet.name:
            return
        if new.lower() != sheet.name.lower() and (
                any(s.name.lower() == new.lower() for s in self.sheets)
                or os.path.exists(os.path.join(NOTES_DIR, new))):
            messagebox.showwarning("탭 이름", f"'{new}' 탭이 이미 있어요.")
            return
        sheet.flush()
        os.rename(sheet.folder, os.path.join(NOTES_DIR, new))
        sheet.name = new
        self.render_tabs()
        self.root.title(f"멀티 메모 - {new}")
        self.save_meta()

    def move_sheet(self, sheet, d):
        i = self.sheets.index(sheet)
        j = i + d
        if 0 <= j < len(self.sheets):
            self.sheets[i], self.sheets[j] = self.sheets[j], self.sheets[i]
            self.render_tabs()
            self.save_meta()

    def delete_sheet(self, sheet):
        if len(self.sheets) == 1:
            messagebox.showinfo("탭 삭제", "탭이 하나뿐이라 삭제할 수 없어요.")
            return
        if not messagebox.askyesno("탭 삭제", f"'{sheet.name}' 탭과 안의 메모 {len(sheet.notes)}개를 "
                                   f"삭제할까요?\n(_삭제됨 폴더로 옮겨집니다)"):
            return
        sheet.flush()
        i = self.sheets.index(sheet)
        self.sheets.remove(sheet)
        if self.current is sheet:
            self.current = None
            sheet.hide()
            self.switch_to(self.sheets[max(0, i - 1)])
        sheet.board.destroy()
        os.makedirs(TRASH_DIR, exist_ok=True)
        shutil.move(sheet.folder, os.path.join(TRASH_DIR, f"{sheet.name}_{time.strftime('%Y%m%d_%H%M%S')}"))
        self.render_tabs()
        self.save_meta()

    # ---------- 탭 저장 / 불러오기 (.zip) ----------
    def export_sheet(self):
        sheet = self.current
        sheet.flush()
        sheet.save_meta()
        path = filedialog.asksaveasfilename(
            parent=self.root, title="탭 저장", initialfile=f"{sheet.name}.zip",
            defaultextension=".zip", filetypes=[("멀티 메모 탭", "*.zip")])
        if not path:
            return
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for fn in os.listdir(sheet.folder):
                p = os.path.join(sheet.folder, fn)
                if os.path.isfile(p) and (fn.endswith(".txt") or fn == SHEET_META):
                    z.write(p, fn)
        self.show_tip(f"'{sheet.name}' 탭 저장됨 → {os.path.basename(path)}")

    def import_sheet(self):
        path = filedialog.askopenfilename(parent=self.root, title="탭 불러오기",
                                          filetypes=[("멀티 메모 탭", "*.zip")])
        if not path:
            return
        try:
            with zipfile.ZipFile(path) as z:
                members = [m for m in z.infolist() if not m.is_dir()]
                name = self._unique_name(os.path.splitext(os.path.basename(path))[0])
                folder = os.path.join(NOTES_DIR, name)
                os.makedirs(folder)
                count = 0
                for m in members:
                    fn = os.path.basename(m.filename)  # 하위 폴더 경로는 무시
                    if fn.endswith(".txt") or fn == SHEET_META:
                        with z.open(m) as src, open(os.path.join(folder, fn), "wb") as dst:
                            shutil.copyfileobj(src, dst)
                        count += fn.endswith(".txt")
        except (zipfile.BadZipFile, OSError) as e:
            messagebox.showerror("탭 불러오기", f"파일을 열 수 없어요.\n{e}")
            return
        sheet = Sheet(self, name)
        sheet.load()
        self.sheets.append(sheet)
        self.switch_to(sheet)
        self.show_tip(f"'{name}' 탭으로 메모 {count}개를 불러왔어요")

    # ---------- 글자 ----------
    def _refresh_font_controls(self):
        self.size_label.config(text=f"{self.font_size}pt")
        for name, b in self.weight_btns.items():
            on = name == self.font_weight
            b.config(bg="#FFFFFF" if on else TAB_BG, fg="#222" if on else "#888")

    def set_font_size(self, size):
        self.font_size = max(FONT_MIN, min(FONT_MAX, size))
        self.body_font.config(size=self.font_size)
        self._refresh_font_controls()
        self.save_meta()

    def set_font_weight(self, name):
        self.font_weight = name
        family, weight = WEIGHTS[name]
        self.body_font.config(family=family, weight=weight)
        self._refresh_font_controls()
        self.save_meta()

    def on_zoom(self, event):
        self.set_font_size(self.font_size + (1 if event.delta > 0 else -1))
        return "break"

    # ---------- 저장 / 종료 ----------
    def save_meta(self):
        if not getattr(self, "current", None):
            return
        write_json(META_FILE, {
            "tabs": [s.name for s in self.sheets],
            "active": self.sheets.index(self.current) if self.current in self.sheets else 0,
            "font_size": self.font_size,
            "font_weight": self.font_weight,
            "geometry": self.root.geometry(),
        })

    def on_close(self):
        for s in self.sheets:
            s.flush()
            s.save_meta()
        self.save_meta()
        self.root.destroy()

    def run(self):
        self.root.after(100, self.current.relayout)
        self.root.mainloop()


if __name__ == "__main__":
    App().run()
