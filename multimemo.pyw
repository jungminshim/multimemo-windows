"""멀티 메모 - 한 화면에 여러 개의 스티커 메모를 동시에 띄우는 프로그램.

- 엑셀 시트처럼 아래쪽 탭마다 별도의 메모판 (탭 추가/이름 변경/삭제/순서 이동)
- 메모 개수에 맞춰 화면을 자동으로 나눠 배치 (1개=전체, 2개=반반, 4개=2x2 ...)
- 각 메모는 <저장 폴더>/<탭 이름>/ 에 개별 .txt 파일로 자동 저장
  (윈도우: 프로그램 옆 notes 폴더, 맥: 문서/멀티메모)
- 탭 저장: 탭 하나를 .zip 파일로 저장하고, 다른 곳에서 탭으로 불러오기
- 글자 서식: 글자를 드래그하면 서식 바(굵게/기울임/취소선/형광펜/체크리스트)
  체크리스트 ☐ 를 누르면 ☑ + 취소선. 서식은 메모별 .fmt.json 에 따로 저장
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
ICON_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "multimemo.ico")
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
BOX_FONTS = ["Segoe UI Symbol", "Apple Symbols"]   # 체크박스 ☐☑ 를 예쁘게 그리는 글꼴
BOX_FAMILY = None
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
    global BOX_FAMILY
    installed = set(tkfont.families())
    BOX_FAMILY = next((c for c in BOX_FONTS if c in installed), FONT_FAMILY)
    for name, (cands, wt) in WEIGHT_SPECS.items():
        WEIGHTS[name] = (next((c for c in cands if c in installed), FONT_FAMILY), wt)


def open_folder(path):
    if IS_MAC:
        subprocess.Popen(["open", path])
    else:
        os.startfile(path)


APP_ID = "MultiMemo.App"


def set_relaunch_info(root):
    """작업표시줄에 고정할 때 파이썬이 아니라 '멀티 메모'(아이콘·실행 명령 포함)로 고정되도록 등록."""
    if IS_MAC:
        return
    try:
        import pythoncom
        from win32com.propsys import propsys, pscon
    except ImportError:  # pywin32가 없으면 건너뜀 (고정 시 파이썬 아이콘으로 고정될 수 있음)
        return
    try:
        root.update_idletasks()
        hwnd = int(root.wm_frame(), 16)
        store = propsys.SHGetPropertyStoreForWindow(hwnd, propsys.IID_IPropertyStore)
        pyw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
        values = [
            (pscon.PKEY_AppUserModel_ID, APP_ID),
            (pscon.PKEY_AppUserModel_RelaunchCommand, f'"{pyw}" "{os.path.abspath(__file__)}"'),
            (pscon.PKEY_AppUserModel_RelaunchDisplayNameResource, "멀티 메모"),
            (pscon.PKEY_AppUserModel_RelaunchIconResource, ICON_FILE + ",0"),
        ]
        for key, val in values:
            store.SetValue(key, propsys.PROPVARIANTType(val, pythoncom.VT_LPWSTR))
        store.Commit()
    except Exception:
        pass


def rounded_rect(canvas, x1, y1, x2, y2, r, **kw):
    pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
           x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return canvas.create_polygon(pts, smooth=True, **kw)


BOX_OPEN, BOX_DONE = "☐", "☑"
FORMAT_TAGS = ("b", "i", "s", "hl")   # 저장되는 서식: 굵게, 기울임, 취소선, 형광펜
HIGHLIGHT = "#FFF08A"


class FormatBar:
    """글자를 드래그하면 위에 뜨는 작은 서식 바 (굵게/기울임/취소선/형광펜/체크리스트/지우기)."""

    ITEMS = [
        ("b", "B", "굵게 (Ctrl+B)"),
        ("i", "I", "기울임 (Ctrl+I)"),
        ("s", "S", "취소선 (Ctrl+Shift+X)"),
        ("hl", "형광펜", "형광펜 (Ctrl+Shift+H)"),
        ("check", "☑ 체크", "체크리스트 (Ctrl+L)"),
        ("clear", "지우기", "서식 지우기"),
    ]

    def __init__(self, app):
        self.app = app
        self.note = None
        self.win = tk.Toplevel(app.root)
        self.win.withdraw()
        self.win.overrideredirect(True)
        self.win.attributes("-topmost", True)
        self.win.configure(bg="#C9CCD2")
        row = tk.Frame(self.win, bg="#FFFFFF")
        row.pack(padx=1, pady=1)
        self.buttons = {}
        for key, label, tip in self.ITEMS:
            font = {
                "b": (FONT_FAMILY, 10, "bold"),
                "i": (FONT_FAMILY, 10, "italic"),
                "s": (FONT_FAMILY, 10, "overstrike"),
            }.get(key, (FONT_FAMILY, 9))
            b = tk.Label(row, text=label, bg="#FFFFFF", fg="#333", font=font, padx=9, pady=3,
                         cursor="hand2")
            if key == "hl":
                b.config(bg=HIGHLIGHT)
            b.pack(side="left", padx=1, pady=1)
            b.bind("<Button-1>", lambda e, k=key: self.apply(k))
            b.bind("<Enter>", lambda e, t=tip: app.show_tip(t))
            b.bind("<Leave>", lambda e: app.show_tip(""))
            self.buttons[key] = b
        self._noactivate()

    def _noactivate(self):
        # 서식 바를 눌러도 메모 창의 포커스/선택이 풀리지 않도록 (윈도우 전용)
        if IS_MAC:
            return
        try:
            self.win.update_idletasks()
            hwnd = int(self.win.wm_frame(), 16)
            u = ctypes.windll.user32
            style = u.GetWindowLongW(hwnd, -20)
            u.SetWindowLongW(hwnd, -20, style | 0x08000000 | 0x80)  # NOACTIVATE | TOOLWINDOW
        except Exception:
            pass

    def show(self, note, x_root, y_root):
        self.note = note
        self.refresh_state()
        self.win.update_idletasks()
        w, h = self.win.winfo_reqwidth(), self.win.winfo_reqheight()
        sw = self.win.winfo_screenwidth()
        x = max(0, min(x_root - w // 2, sw - w))
        y = max(0, y_root - h - 14)
        self.win.geometry(f"+{x}+{y}")
        self.win.deiconify()
        self.win.lift()

    def hide(self):
        if self.win.winfo_ismapped():
            self.win.withdraw()
        self.note = None

    def refresh_state(self):
        n = self.note
        if not n:
            return
        for key in ("b", "i", "s"):
            on = n.sel_has(key)
            self.buttons[key].config(bg="#DDE3EC" if on else "#FFFFFF")

    def apply(self, key):
        n = self.note
        if not n:
            return
        if key == "check":
            n.toggle_checklist()
        elif key == "clear":
            n.clear_format()
        else:
            n.toggle_format(key)
        self.refresh_state()


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
        self.btn_check = self._header_button("☑", self.toggle_checklist,
                                             f"체크리스트 ({MOD_LABEL}L) - 커서가 있는 줄/선택한 줄")

        self.sep = tk.Frame(self.inner, bg=accent, height=1)
        self.sep.pack(fill="x", padx=4)

        self.text = t = tk.Text(self.inner, wrap="word", bd=0, bg=bg, padx=8, pady=8,
                                font=app.body_font, undo=True,
                                relief="flat", insertbackground="#333", fg="#333",
                                highlightthickness=0)
        t.pack(fill="both", expand=True)
        self._setup_tags()
        t.bind("<<Modified>>", self._on_modified)
        t.bind(f"<{MOD}-MouseWheel>", app.on_zoom)
        t.bind("<ButtonRelease-1>", self._maybe_show_bar)
        t.bind("<KeyRelease>", self._on_key_release)
        t.bind("<Return>", self._on_return)
        t.bind(f"<{MOD}-b>", lambda e: self.toggle_format("b"))
        t.bind(f"<{MOD}-i>", lambda e: self.toggle_format("i"))
        t.bind(f"<{MOD}-Shift-X>", lambda e: self.toggle_format("s"))
        t.bind(f"<{MOD}-Shift-H>", lambda e: self.toggle_format("hl"))
        t.bind(f"<{MOD}-l>", lambda e: self.toggle_checklist())

        self.status = tk.Label(self.inner, text="", bg=bg, fg="#AAA",
                               font=(FONT_FAMILY, 8), anchor="e")
        self.status.pack(fill="x")

    # ---------- 서식 태그 ----------
    def _setup_tags(self):
        t, f = self.text, self.app.fonts
        t.tag_configure("hl", background=HIGHLIGHT)
        t.tag_configure("s", overstrike=True)
        t.tag_configure("f_b", font=f["b"])
        t.tag_configure("f_i", font=f["i"])
        t.tag_configure("f_bi", font=f["bi"])
        t.tag_configure("done", overstrike=True, foreground="#A0A0A0")
        t.tag_configure("box", font=f["box"], foreground="#8A8F98")
        t.tag_configure("boxdone", font=f["box"], foreground="#1A7F37")
        for tag in ("box", "boxdone"):
            t.tag_bind(tag, "<Enter>", lambda e: t.config(cursor="hand2"))
            t.tag_bind(tag, "<Leave>", lambda e: t.config(cursor="xterm"))
            t.tag_bind(tag, "<Button-1>", self._on_box_click)
        t.tag_raise("sel")

    def _offset(self, index):
        return len(self.text.get("1.0", index))

    def _ranges(self, tag):
        r = self.text.tag_ranges(tag)
        return [[self._offset(r[k]), self._offset(r[k + 1])] for k in range(0, len(r), 2)]

    def _refresh_fonts(self):
        """굵게(b)·기울임(i) 겹침을 계산해서 실제 글꼴 태그(f_b/f_i/f_bi)를 다시 입힘."""
        t = self.text
        for tag in ("f_b", "f_i", "f_bi"):
            t.tag_remove(tag, "1.0", "end")
        b, i = self._ranges("b"), self._ranges("i")
        if not b and not i:
            return
        n = self._offset("end-1c")
        flags = bytearray(n + 1)
        for s, e in b:
            flags[s:e] = bytes(x | 1 for x in flags[s:e])
        for s, e in i:
            flags[s:e] = bytes(x | 2 for x in flags[s:e])
        names = {1: "f_b", 2: "f_i", 3: "f_bi"}
        start = 0
        for k in range(1, n + 1):
            if k == n or flags[k] != flags[start]:
                if flags[start]:
                    t.tag_add(names[flags[start]], f"1.0+{start}c", f"1.0+{k}c")
                start = k

    def _sel(self):
        try:
            return self.text.index("sel.first"), self.text.index("sel.last")
        except tk.TclError:
            return None

    def sel_has(self, tag):
        """선택 영역 전체에 tag가 적용돼 있는지."""
        sel = self._sel()
        if not sel:
            return False
        a, b = sel
        r = self.text.tag_prevrange(tag, f"{a}+1c")
        return bool(r) and self.text.compare(r[0], "<=", a) and self.text.compare(r[1], ">=", b)

    def toggle_format(self, tag):
        sel = self._sel()
        if not sel:
            return "break"
        if self.sel_has(tag):
            self.text.tag_remove(tag, *sel)
        else:
            self.text.tag_add(tag, *sel)
        if tag in ("b", "i"):
            self._refresh_fonts()
        self.schedule_save()
        return "break"

    def clear_format(self):
        sel = self._sel()
        if not sel:
            return
        for tag in FORMAT_TAGS:
            self.text.tag_remove(tag, *sel)
        self._refresh_fonts()
        self.schedule_save()

    # ---------- 체크리스트 ----------
    def _line_range(self):
        sel = self._sel()
        if sel:
            first = int(sel[0].split(".")[0])
            last_idx = sel[1]
            last = int(last_idx.split(".")[0])
            if last > first and last_idx.endswith(".0"):
                last -= 1   # 다음 줄 맨 앞에서 끝난 선택은 그 줄 제외
        else:
            first = last = int(self.text.index("insert").split(".")[0])
        return first, last

    def toggle_checklist(self):
        """커서가 있는 줄(또는 선택한 줄들)을 체크리스트로 바꾸거나 해제."""
        t = self.text
        first, last = self._line_range()
        lines = range(first, last + 1)
        is_box = lambda ln: t.get(f"{ln}.0") in (BOX_OPEN, BOX_DONE)
        t.edit_separator()
        if all(is_box(ln) for ln in lines):
            for ln in lines:
                end = f"{ln}.2" if t.get(f"{ln}.1") == " " else f"{ln}.1"
                t.delete(f"{ln}.0", end)
        else:
            for ln in lines:
                if not is_box(ln):
                    t.insert(f"{ln}.0", BOX_OPEN + " ")
        t.edit_separator()
        self.refresh_checks()
        return "break"

    def refresh_checks(self):
        """☐/☑ 로 시작하는 줄에 체크박스 모양과 완료(취소선) 표시를 입힘."""
        t = self.text
        for tag in ("box", "boxdone", "done"):
            t.tag_remove(tag, "1.0", "end")
        last = int(t.index("end-1c").split(".")[0])
        for ln in range(1, last + 1):
            ch = t.get(f"{ln}.0")
            if ch == BOX_OPEN:
                t.tag_add("box", f"{ln}.0")
            elif ch == BOX_DONE:
                t.tag_add("boxdone", f"{ln}.0")
                t.tag_add("done", f"{ln}.1", f"{ln}.end")

    def _on_box_click(self, event):
        t = self.text
        ln = t.index(f"@{event.x},{event.y}").split(".")[0]
        ch = t.get(f"{ln}.0")
        if ch in (BOX_OPEN, BOX_DONE):
            t.edit_separator()
            t.delete(f"{ln}.0")
            t.insert(f"{ln}.0", BOX_DONE if ch == BOX_OPEN else BOX_OPEN)
            t.edit_separator()
            self.refresh_checks()
            t.mark_set("insert", f"{ln}.end")
        return "break"

    def _on_return(self, _event):
        """체크리스트 줄에서 엔터: 다음 줄도 ☐ 로 시작. 빈 체크 항목에서 엔터: 목록 끝내기."""
        t = self.text
        if self._sel():
            return None
        ln = t.index("insert").split(".")[0]
        line = t.get(f"{ln}.0", f"{ln}.end")
        if not line or line[0] not in (BOX_OPEN, BOX_DONE):
            return None
        if not line[1:].strip():
            t.delete(f"{ln}.0", f"{ln}.end")
            return "break"
        t.insert("insert", "\n" + BOX_OPEN + " ")
        t.see("insert")
        return "break"

    # ---------- 서식 바 ----------
    def _maybe_show_bar(self, event):
        if self._sel():
            self.app.format_bar.show(self, event.x_root, event.y_root)
        else:
            self.app.format_bar.hide()

    def _on_key_release(self, event):
        if event.keysym in ("Shift_L", "Shift_R") or (event.state & 0x1):
            sel = self._sel()
            if sel:
                bbox = self.text.bbox(sel[0])
                if bbox:
                    x = self.text.winfo_rootx() + bbox[0]
                    y = self.text.winfo_rooty() + bbox[1]
                    self.app.format_bar.show(self, x + 40, y)
                return
        if self.app.format_bar.note is self and not self._sel():
            self.app.format_bar.hide()

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

    @property
    def fmt_path(self):
        return os.path.join(self.sheet.folder, f"{self.id}.fmt.json")

    def load_text(self):
        if os.path.exists(self.path):
            with open(self.path, encoding="utf-8") as f:
                self.text.insert("1.0", f.read())
        fmt = read_json(self.fmt_path, {})
        for tag in FORMAT_TAGS:
            for s, e in fmt.get(tag, []):
                self.text.tag_add(tag, f"1.0+{s}c", f"1.0+{e}c")
        self._refresh_fonts()
        self.refresh_checks()
        self.text.edit_modified(False)
        self.text.edit_reset()

    def _on_modified(self, _event):
        if self.text.edit_modified():
            self.text.edit_modified(False)
            self.refresh_checks()
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
        # 서식은 메모별 .fmt.json 에 글자 위치(시작, 끝)로 따로 저장
        fmt = {tag: self._ranges(tag) for tag in FORMAT_TAGS}
        fmt = {k: v for k, v in fmt.items() if v}
        if fmt:
            write_json(self.fmt_path, fmt)
        elif os.path.exists(self.fmt_path):
            os.remove(self.fmt_path)
        self.sheet.save_meta()
        self.status.config(text=f"저장됨 {time.strftime('%H:%M:%S')}")

    def flush(self):
        if self._save_job:
            self.app.root.after_cancel(self._save_job)
            self.save()

    def next_color(self):
        self.color = (self.color + 1) % len(COLORS)
        bg, accent = COLORS[self.color]
        for w in (self.inner, self.header, self.title_entry, self.btn_del, self.btn_check,
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
        os.makedirs(TRASH_DIR, exist_ok=True)
        stamp = time.strftime("%Y%m%d_%H%M%S")
        for src, ext in ((self.path, ".txt"), (self.fmt_path, ".fmt.json")):
            if os.path.exists(src):
                shutil.move(src, os.path.join(TRASH_DIR, f"{self.sheet.name}_{self.id}_{stamp}{ext}"))
        if self.app.format_bar.note is self:
            self.app.format_bar.hide()
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
        # 작업표시줄에 파이썬 대신 멀티 메모 아이콘이 따로 뜨도록 앱 ID 지정
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
        except (AttributeError, OSError):
            pass
        self.root = tk.Tk()
        if not IS_MAC and os.path.exists(ICON_FILE):
            try:
                self.root.iconbitmap(default=ICON_FILE)
            except tk.TclError:
                pass
        if IS_MAC:
            self.root.tk.call("tk", "scaling", MAC_SCALE)
        else:
            dpi = self.root.winfo_fpixels("1i")
            self.root.tk.call("tk", "scaling", dpi / 72 * UI_SCALE)
        resolve_fonts()
        family, weight = WEIGHTS[self.font_weight]
        self.body_font = tkfont.Font(family=family, size=self.font_size, weight=weight)
        # 서식용 글꼴 (굵게/기울임/굵게+기울임/체크박스) - 글자 크기·굵기를 바꾸면 같이 바뀜
        self.fonts = {k: tkfont.Font() for k in ("b", "i", "bi", "box")}
        self._update_fonts()
        self.root.title("멀티 메모")
        self.root.geometry(meta.get("geometry", "1200x800"))
        self.root.configure(bg=BOARD_BG)
        self.root.minsize(400, 300)

        self._build_toolbar()
        self.format_bar = FormatBar(self)
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
        # 다른 곳을 누르거나 창을 옮기거나 다른 프로그램으로 가면 서식 바 숨기기
        self.root.bind_all("<ButtonPress-1>", self._maybe_hide_bar, add="+")
        self.root.bind("<Configure>", lambda e: self.format_bar.hide() if e.widget is self.root else None)
        self.root.bind("<FocusOut>", lambda e: self.root.after(150, self._hide_bar_if_inactive))

    def _maybe_hide_bar(self, event):
        try:
            if event.widget.winfo_toplevel() is self.format_bar.win:
                return
        except (AttributeError, tk.TclError):
            pass
        self.format_bar.hide()

    def _hide_bar_if_inactive(self):
        try:
            focused = self.root.focus_displayof()
        except (KeyError, tk.TclError):
            focused = None
        if focused is None:
            self.format_bar.hide()

    def _update_fonts(self):
        family, weight = WEIGHTS[self.font_weight]
        bold_family = WEIGHTS["굵게"][0]
        size = self.font_size
        self.fonts["b"].config(family=bold_family, size=size, weight="bold", slant="roman")
        self.fonts["i"].config(family=family, size=size, weight=weight, slant="italic")
        self.fonts["bi"].config(family=bold_family, size=size, weight="bold", slant="italic")
        self.fonts["box"].config(family=BOX_FAMILY, size=size + 1)

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
        if hasattr(self, "format_bar"):
            self.format_bar.hide()
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
                if os.path.isfile(p) and (fn.endswith((".txt", ".fmt.json")) or fn == SHEET_META):
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
                    if fn.endswith((".txt", ".fmt.json")) or fn == SHEET_META:
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
        self._update_fonts()
        self._refresh_font_controls()
        self.save_meta()

    def set_font_weight(self, name):
        self.font_weight = name
        family, weight = WEIGHTS[name]
        self.body_font.config(family=family, weight=weight)
        self._update_fonts()
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
        self.root.after(200, lambda: set_relaunch_info(self.root))
        self.root.mainloop()


if __name__ == "__main__":
    App().run()
