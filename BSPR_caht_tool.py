import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog
import pyautogui
import pytesseract
from PIL import Image, ImageTk
import threading
import time
import cv2
import numpy as np
import os
import ctypes
from ctypes import wintypes
import re
import queue
import gc
from collections import deque
import json
import difflib
import unicodedata

# =========================
# Config persistence
# =========================

def get_config_path():
    """Save config next to this script (portable)."""
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        base_dir = os.getcwd()
    return os.path.join(base_dir, "starezo_chat_tool_config.json")

CONFIG_PATH = get_config_path()

# Default configuration (used when config file is missing or to fill missing keys)
DEFAULT_CONFIG = {
    "version": 3,
    "remember_area": True,
    "area": [
        59,
        829,
        290,
        275
    ],
    "window": {
        "geometry": "360x760+-4+1",
        "topmost": True,
        "alpha": 0.95
    },
    "ocr": {
        "threshold": 220,
        "interval": 1.0,
        "scale": 1.0
    },
    "dedupe": {
        "use_similarity": True,
        "threshold": 0.96,
        "history": 220
    },
    "iconmask": {
        "enabled": True,
        "mask_w": 24,
        "mask_h": 26,
        "kernel_w": 60,
        "kernel_h": 18,
        "min_area": 450,
        "merge_gap": 10,
        "use_bubble": True,
        "bubble_v_min": 190,
        "bubble_s_max": 90,
        "bubble_kernel_w": 35,
        "bubble_kernel_h": 21,
        "bubble_min_area": 1600,
        "bubble_pad": 2
    },
    "log": {
        "world_retention_enabled": True,
        "world_retention_minutes": 5,
        "font_size": 8,
        "line_spacing": 0,
        "card_gap": 0,
        "display_format": "simple"
    },
    "sound": {
        "enabled": True,
        "keyword_ignore_spaces": True,
        "kana_variants": True,
        "scope_enabled": {
            "[ワールド]": False,
            "[ギルド]": True,
            "[パーティ]": False,
            "[チャネル]": False
        },
        "scope_file": {
            "[ワールド]": "",
            "[ギルド]": "",
            "[パーティ]": "",
            "[チャネル]": ""
        },
        "keyword_rules": []
    }
}


# =========================
# DPI awareness (Windows)
# =========================
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass



# =========================
# Optional sound (Windows)
# =========================
try:
    import winsound  # Windows only
except Exception:
    winsound = None
try:
    # MCI (winmm) for MP3/WAV playback (Windows). Works without extra packages.
    _winmm = ctypes.WinDLL("winmm")
    _mciSendStringW = _winmm.mciSendStringW
    _mciSendStringW.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.UINT, wintypes.HWND]
    _mciSendStringW.restype = wintypes.UINT
except Exception:
    _mciSendStringW = None


def _mci_cmd(cmd: str) -> bool:
    if not _mciSendStringW:
        return False
    try:
        buf = ctypes.create_unicode_buffer(256)
        r = _mciSendStringW(cmd, buf, 255, 0)
        return r == 0
    except Exception:
        return False


def _mci_status_int(alias: str, what: str) -> int | None:
    """Return integer status (e.g., length in ms) or None."""
    if not _mciSendStringW:
        return None
    try:
        buf = ctypes.create_unicode_buffer(64)
        r = _mciSendStringW(f"status {alias} {what}", buf, 63, 0)
        if r != 0:
            return None
        s = (buf.value or "").strip()
        return int(s) if s.isdigit() else None
    except Exception:
        return None


def _mci_play_file(path: str) -> bool:
    """Play file async via MCI. Supports mp3/wav on Windows."""
    if not _mciSendStringW:
        return False
    try:
        apath = os.path.abspath(path)
        alias = f"ct_{int(time.time()*1000)}"
        # close if collides (rare)
        _mci_cmd(f"close {alias}")
        if not _mci_cmd(f'open "{apath}" alias {alias}'):
            return False
        _mci_cmd(f"play {alias} from 0")
        length = _mci_status_int(alias, "length") or 1200  # ms
        # close later to avoid handle leak
        def _closer():
            try:
                time.sleep(length / 1000.0 + 0.8)
                _mci_cmd(f"close {alias}")
            except Exception:
                pass
        threading.Thread(target=_closer, daemon=True).start()
        return True
    except Exception:
        return False



def play_sound_file(path: str | None, root: tk.Tk | None = None):
    """
    Play notification sound asynchronously.
    - WAV: winsound if available (fast)
    - MP3/WAV: MCI (winmm) if available
    If path is None/empty or playback fails, play a default short beep.
    """
    # Prefer explicit file if provided
    if path and os.path.exists(path):
        ext = os.path.splitext(path)[1].lower()
        # WAV via winsound
        if ext == ".wav" and winsound:
            try:
                winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_ASYNC)
                return
            except Exception:
                pass
        # MP3/WAV via MCI
        try:
            if _mci_play_file(path):
                return
        except Exception:
            pass

    # fallback beep
    if winsound:
        try:
            winsound.Beep(1200, 70)  # "ピッ"
            return
        except Exception:
            try:
                winsound.MessageBeep(winsound.MB_OK)
                return
            except Exception:
                pass

    # non-windows fallback
    try:
        if root:
            root.bell()
    except Exception:
        pass


def get_tesseract_path():

    paths = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        r"C:\Users\{}\AppData\Local\Tesseract-OCR\tesseract.exe".format(os.getlogin()),
    ]
    for p in paths:
        if os.path.exists(p):
            return p
    return "NOT_FOUND"


TESSERACT_PATH = get_tesseract_path()
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH


# =========================
# Region selection overlay
# =========================
class SelectionWindow:
    def __init__(self, master=None):
        self.root = tk.Toplevel(master)
        self.root.attributes("-alpha", 0.25, "-fullscreen", True, "-topmost", True)
        self.root.config(cursor="cross")

        self.canvas = tk.Canvas(self.root, cursor="cross", bg="#000000", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        self.start_x = self.start_y = None
        self.rect_id = None
        self.result = None

        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_move)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.root.bind("<Escape>", lambda e: self.close())

        # ensure this overlay receives mouse/keyboard immediately
        try:
            self.root.focus_force()
        except Exception:
            pass
        try:
            self.root.grab_set()
        except Exception:
            pass

    def on_press(self, e):
        self.start_x, self.start_y = e.x, e.y
        self.rect_id = self.canvas.create_rectangle(e.x, e.y, e.x, e.y, outline="#00E5FF", width=2)

    def on_move(self, e):
        if self.rect_id is not None:
            self.canvas.coords(self.rect_id, self.start_x, self.start_y, e.x, e.y)

    def on_release(self, e):
        self.result = (
            min(self.start_x, e.x),
            min(self.start_y, e.y),
            abs(self.start_x - e.x),
            abs(self.start_y - e.y),
        )
        self.close()

    def close(self):
        # Do NOT call .quit() here. That would stop the application's mainloop.
        try:
            self.root.grab_release()
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass

    def get_selection(self):
        # Wait until this toplevel is destroyed (no nested mainloop).
        try:
            self.root.wait_window()
        except Exception:
            pass
        return self.result

# =========================
# Scrollable frame for log cards
# =========================
class ScrolledFrame(ttk.Frame):
    """Canvas-based vertical scroller for log cards.

    長時間運用でカードの追加/削除を繰り返すと、Tkが<Configure>イベント連鎖を取りこぼして
    scrollregion が古いままになり、「スクロールバーだけ伸びて中身が空白に見える」ことがあります。
    その対策として refresh_scrollregion() を明示的に呼べるようにし、内部の更新も after_idle でデバウンスします。
    """

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        self.canvas = tk.Canvas(self, bg="#FFFFFF", highlightthickness=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.inner = ttk.Frame(self.canvas, style="Card.TFrame")
        self.window_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        # Debounced refresh job id
        self._refresh_job = None

        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # mouse wheel support
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)          # Windows
        self.canvas.bind_all("<Button-4>", self._on_mousewheel_linux)      # Linux
        self.canvas.bind_all("<Button-5>", self._on_mousewheel_linux)

    def _on_inner_configure(self, _evt=None):
        # Debounce refresh to avoid excessive configure churn.
        try:
            if self._refresh_job is not None:
                self.after_cancel(self._refresh_job)
        except Exception:
            pass
        try:
            self._refresh_job = self.after_idle(self.refresh_scrollregion)
        except Exception:
            self._refresh_job = None

    def _on_canvas_configure(self, evt):
        try:
            self.canvas.itemconfigure(self.window_id, width=evt.width)
        except Exception:
            pass

    def refresh_scrollregion(self):
        """Force recompute scrollregion (stale region recovery)."""
        try:
            self.update_idletasks()
        except Exception:
            pass
        try:
            bbox = self.canvas.bbox("all")
            if bbox:
                self.canvas.configure(scrollregion=bbox)
            else:
                self.canvas.configure(scrollregion=(0, 0, 0, 0))
        except Exception:
            pass
        finally:
            self._refresh_job = None

    def _on_mousewheel(self, event):
        try:
            x, y = self.canvas.winfo_pointerx(), self.canvas.winfo_pointery()
            widget = self.canvas.winfo_containing(x, y)
            if widget and (widget == self.canvas or str(widget).startswith(str(self.canvas))):
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass

    def _on_mousewheel_linux(self, event):
        try:
            if event.num == 4:
                self.canvas.yview_scroll(-3, "units")
            elif event.num == 5:
                self.canvas.yview_scroll(3, "units")
        except Exception:
            pass

    def scroll_to_bottom(self):
        try:
            self.refresh_scrollregion()
            self.canvas.yview_moveto(1.0)
        except Exception:
            pass


# =========================
# Settings window
# =========================
class SettingsWindow(tk.Toplevel):
    def __init__(self, app: "ChatMonitorApp"):
        super().__init__(app.root)
        self.app = app
        self.title("設定")
        self.geometry("940x860")
        self.minsize(860, 780)
        self.configure(bg="#F6F8FC")
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        # モーダル（メイン画面の多重起動防止）
        try:
            self.transient(app.root)
        except Exception:
            pass
        try:
            self.grab_set()
        except Exception:
            pass
        try:
            self.focus_set()
        except Exception:
            pass
        try:
            self.bind("<Destroy>", self._on_destroy)
        except Exception:
            pass

        header = ttk.Frame(self, padding=(16, 12), style="Header.TFrame")
        header.pack(fill="x")
        ttk.Label(header, text="設定", style="HeaderTitle.TLabel").pack(anchor="w")
        ttk.Label(header, text="キャプチャ範囲 / OCR / 通知を調整します", style="HeaderSub.TLabel").pack(anchor="w", pady=(2, 0))

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=16, pady=12)

        self.tab_capture = ttk.Frame(nb)
        self.tab_ocr = ttk.Frame(nb)
        self.tab_window = ttk.Frame(nb)
        self.tab_notify = ttk.Frame(nb)
        self.tab_keywords = ttk.Frame(nb)
        self.tab_iconmask = ttk.Frame(nb)
        self.tab_dedupe = ttk.Frame(nb)
        # ログタブは「表示」タブに統合
        self.tab_logs = self.tab_window

        nb.add(self.tab_capture, text="キャプチャ")
        nb.add(self.tab_iconmask, text="アイコンマスク")
        nb.add(self.tab_ocr, text="OCR")
        nb.add(self.tab_window, text="表示")
        nb.add(self.tab_notify, text="通知")
        nb.add(self.tab_keywords, text="キーワード")
        nb.add(self.tab_dedupe, text="重複判定")

        self._build_tab_capture()
        self._build_tab_iconmask()
        self._build_tab_ocr()
        self._build_tab_window()
        self._build_tab_notify()
        self._build_tab_keywords()
        self._build_tab_dedupe()
        self._build_tab_logs()

        footer = ttk.Frame(self, padding=(16, 10), style="Header.TFrame")
        footer.pack(fill="x")
        ttk.Button(footer, text="閉じる", style="Secondary.TButton", command=self.on_close).pack(side="right")

    def _build_tab_capture(self):
        card = ttk.Frame(self.tab_capture, style="Card.TFrame", padding=(16, 14))
        card.pack(fill="both", expand=True, padx=14, pady=14)

        ttk.Label(card, text="キャプチャ範囲", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        self.pill = ttk.Label(
            card,
            text="未設定" if not self.app.area else "設定済み",
            style="PillNg.TLabel" if not self.app.area else "PillOk.TLabel"
        )
        self.pill.grid(row=0, column=1, sticky="e")

        ttk.Label(card, text="ゲームのチャット枠が入るようにドラッグで指定してください。", style="Hint.TLabel")\
            .grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 10))

        ttk.Button(card, text="範囲を指定する", style="Secondary.TButton", command=self._select_area).grid(row=2, column=0, sticky="w")
        self.lbl = ttk.Label(card, text=f"(x, y, w, h): {self.app.area if self.app.area else '—'}", style="Body.TLabel")
        self.lbl.grid(row=2, column=1, sticky="e")

        # remember area toggle
        self.var_remember_area = tk.BooleanVar(value=self.app.remember_area)
        ttk.Checkbutton(
            card,
            text="読み取り範囲を次回起動時に復元する",
            variable=self.var_remember_area,
            command=self._on_toggle_remember_area
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))

        ttk.Label(card, text="指定直後のキャプチャ（確認用）", style="Hint.TLabel")\
            .grid(row=4, column=0, columnspan=2, sticky="w", pady=(12, 6))

        self.snap_wrap = tk.Frame(card, bg="#0B1220", highlightthickness=1, highlightbackground="#E2E8F0")
        self.snap_wrap.grid(row=5, column=0, columnspan=2, sticky="ew")
        self.snap_wrap.configure(height=240)
        self.snap_wrap.grid_propagate(False)

        self.snap_label = tk.Label(self.snap_wrap, bg="#0B1220")
        self.snap_label.pack(expand=True)

        card.columnconfigure(0, weight=1)
        card.columnconfigure(1, weight=1)

        self._refresh_snapshot()

    def _refresh_snapshot(self):
        if not self.app.area:
            self.snap_label.config(image="")
            self.snap_label.image = None
            return
        try:
            shot = pyautogui.screenshot(region=self.app.area)
            self.app.last_area_snapshot = shot.copy()
            img = shot.copy()
            max_w, max_h = 780, 240
            w, h = img.size
            scale = min(max_w / w, max_h / h)
            scale = max(scale, 0.01)
            img = img.resize((int(w * scale), int(h * scale)), Image.Resampling.LANCZOS)
            imgtk = ImageTk.PhotoImage(img)
            self.snap_label.config(image=imgtk)
            self.snap_label.image = imgtk
        except Exception:
            self.snap_label.config(image="")
            self.snap_label.image = None

    def _on_toggle_remember_area(self):
        self.app.remember_area = bool(self.var_remember_area.get())
        # Save immediately (also clears saved area if turned off)
        self.app.save_settings()

    def _select_area(self):
        # Release modal grab so the selection overlay can receive mouse events
        try:
            self.grab_release()
        except Exception:
            pass

        self.withdraw()
        self.app.root.withdraw()
        try:
            sel = SelectionWindow(self.app.root).get_selection()
        finally:
            try:
                self.app.root.deiconify()
            except Exception:
                pass
            try:
                self.deiconify()
                self.lift()
                self.focus_force()
            except Exception:
                pass
            try:
                self.grab_set()
            except Exception:
                pass

        if sel and sel[2] > 10 and sel[3] > 10:
            self.app.area = sel
            self.lbl.config(text=f"(x, y, w, h): {sel}")
            self.pill.config(text="設定済み", style="PillOk.TLabel")
            self._refresh_snapshot()
            self.app._set_status("範囲を設定しました。")
            self.app.add_system_log(f"範囲設定: {sel}")
            self.app.save_settings()
        else:
            self.app._set_status("範囲指定をキャンセルしました。")

    def _build_tab_ocr(self):
        card = ttk.Frame(self.tab_ocr, style="Card.TFrame", padding=(16, 14))
        card.pack(fill="both", expand=True, padx=14, pady=14)

        ttk.Label(card, text="OCR設定", style="Title.TLabel").grid(row=0, column=0, sticky="w")

        ttk.Label(card, text="二値化しきい値", style="Body.TLabel").grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.thr_val = ttk.Label(card, text=str(self.app.current_threshold), style="Hint.TLabel")
        self.thr_val.grid(row=1, column=1, sticky="e", pady=(10, 0))

        self.thr_scale = ttk.Scale(card, from_=0, to=255, value=self.app.current_threshold, command=self._on_thr)
        self.thr_scale.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(4, 0))

        ttk.Label(card, text="白背景のチャット枠は 200〜235 付近が目安です。", style="Hint.TLabel")\
            .grid(row=3, column=0, columnspan=2, sticky="w", pady=(6, 0))

        ttk.Label(card, text="読み取り間隔（秒）", style="Body.TLabel").grid(row=4, column=0, sticky="w", pady=(14, 0))
        self.int_val = ttk.Label(card, text=f"{self.app.current_interval:.1f}", style="Hint.TLabel")
        self.int_val.grid(row=4, column=1, sticky="e", pady=(14, 0))

        self.int_scale = ttk.Scale(card, from_=0.3, to=3.0, value=self.app.current_interval, command=self._on_int)
        self.int_scale.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(4, 0))

        ttk.Label(card, text="短くしすぎるとCPU負荷が上がります（1.0〜2.0推奨）。", style="Hint.TLabel")\
            .grid(row=6, column=0, columnspan=2, sticky="w", pady=(6, 0))

        state = "OK" if TESSERACT_PATH != "NOT_FOUND" else "NG"
        ttk.Separator(card).grid(row=7, column=0, columnspan=2, sticky="ew", pady=14)
        ttk.Label(card, text="Tesseract", style="Title.TLabel").grid(row=8, column=0, sticky="w")

        pill_style = "PillOk.TLabel" if state == "OK" else "PillNg.TLabel"
        ttk.Label(card, text=("検出済み" if state == "OK" else "未検出"), style=pill_style).grid(row=8, column=1, sticky="e")
        ttk.Label(card, text=f"パス: {TESSERACT_PATH}", style="Hint.TLabel")\
            .grid(row=9, column=0, columnspan=2, sticky="w", pady=(6, 0))

        card.columnconfigure(0, weight=1)
        card.columnconfigure(1, weight=0)



    def _build_tab_iconmask(self):
        card = ttk.Frame(self.tab_iconmask, style="Card.TFrame", padding=(16, 14))
        card.pack(fill="both", expand=True, padx=14, pady=14)

        ttk.Label(card, text="アイコンマスク（発言矩形の左上を白塗り）", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(card, text="名前の前にある人型のアイコンが隠れるようにマスク範囲を指定してください。アイコンの誤読を防ぎます。", style="Hint.TLabel")\
            .grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        self.var_iconmask_enabled = tk.BooleanVar(value=self.app.icon_mask_enabled)
        ttk.Checkbutton(card, text="有効にする", variable=self.var_iconmask_enabled,
                        command=self._on_iconmask_enabled).grid(row=2, column=0, columnspan=2, sticky="w", pady=(12, 0))

        ttk.Label(card, text="マスク幅（px）", style="Body.TLabel").grid(row=3, column=0, sticky="w", pady=(12, 0))
        self.maskw_val = ttk.Label(card, text=str(self.app.icon_mask_w), style="Hint.TLabel")
        self.maskw_val.grid(row=3, column=1, sticky="e", pady=(12, 0))
        self.maskw_scale = ttk.Scale(card, from_=10, to=200, value=self.app.icon_mask_w, command=self._on_maskw)
        self.maskw_scale.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(4, 0))

        ttk.Label(card, text="マスク高さ（px）", style="Body.TLabel").grid(row=5, column=0, sticky="w", pady=(10, 0))
        self.maskh_val = ttk.Label(card, text=str(self.app.icon_mask_h), style="Hint.TLabel")
        self.maskh_val.grid(row=5, column=1, sticky="e", pady=(10, 0))
        self.maskh_scale = ttk.Scale(card, from_=10, to=120, value=self.app.icon_mask_h, command=self._on_maskh)
        self.maskh_scale.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        ttk.Separator(card).grid(row=7, column=0, columnspan=2, sticky="ew", pady=14)

        btn_row = ttk.Frame(card, style="Card.TFrame")
        btn_row.grid(row=8, column=0, columnspan=2, sticky="ew", pady=(14, 6))
        ttk.Button(btn_row, text="今の範囲をキャプチャしてテスト", style="Secondary.TButton", command=self._iconmask_test_capture).pack(side="left")
        ttk.Button(btn_row, text="再解析", style="Secondary.TButton", command=self._refresh_iconmask_preview).pack(side="left", padx=(8, 0))

        self.iconmask_info = ttk.Label(card, text="", style="Hint.TLabel")
        self.iconmask_info.grid(row=9, column=0, columnspan=2, sticky="w", pady=(0, 6))

        self.iconmask_preview_lbl = tk.Label(card, bg="#0B1220")
        self.iconmask_preview_lbl.grid(row=10, column=0, columnspan=2, sticky="ew")

        card.columnconfigure(0, weight=1)
        card.columnconfigure(1, weight=0)

        self._refresh_iconmask_preview()

    def _on_iconmask_enabled(self):
        self.app.icon_mask_enabled = bool(self.var_iconmask_enabled.get())
        self.app.save_settings()
        self._refresh_iconmask_preview()

    def _on_maskw(self, val):
        self.app.icon_mask_w = int(float(val))
        self.maskw_val.config(text=str(self.app.icon_mask_w))
        self.app.save_settings()
        self._refresh_iconmask_preview()

    def _on_maskh(self, val):
        self.app.icon_mask_h = int(float(val))
        self.maskh_val.config(text=str(self.app.icon_mask_h))
        self.app.save_settings()
        self._refresh_iconmask_preview()





    def _iconmask_test_capture(self):
        if not self.app.area:
            self.app._set_status("先に読み取り範囲を設定してください。")
            return
        try:
            shot = pyautogui.screenshot(region=self.app.area)
            self.app.last_area_snapshot = shot.copy()
            self._refresh_iconmask_preview()
        except Exception as e:
            self.app._set_status(f"キャプチャ失敗: {e}")

    def _refresh_iconmask_preview(self):
        try:
            if self.app.last_area_snapshot is None and self.app.area:
                shot = pyautogui.screenshot(region=self.app.area)
                self.app.last_area_snapshot = shot.copy()

            if self.app.last_area_snapshot is None:
                self.iconmask_info.config(text="（範囲を指定してからテストしてください）")
                self.iconmask_preview_lbl.config(image="")
                self.iconmask_preview_lbl.image = None
                return

            frame = cv2.cvtColor(np.array(self.app.last_area_snapshot), cv2.COLOR_RGB2BGR)
            bin_ocr, bin_prev, scale = self.app.preprocess_for_ocr(frame)

            # Always try bubble rectangles for icon-mask preview (per-bubble), regardless of other settings.
            rects = self.app.detect_chat_bubble_rects_auto(frame, bin_prev)

            # If bubble detection failed (or gave a huge ROI), retry with stricter thresholds.
            try:
                H0, W0 = bin_prev.shape[:2]
                if (not rects) or (len(rects) == 1 and rects[0][2] * rects[0][3] > 0.65 * (W0 * H0)):
                    rects = self.app.detect_chat_bubble_rects_auto(
                        frame, bin_prev,
                        vmin_override=int(self.app.bubble_v_min) + 15,
                        smax_override=max(0, int(self.app.bubble_s_max) - 15),
                    )
            except Exception:
                pass

            # Fallback: text-blob rectangles
            if not rects:
                rects = self.app.detect_message_rects(bin_prev)

            # visualize mask region on bin_prev copy
            bin_vis = bin_prev.copy()
            for (x, y, rw, rh) in rects:
                x0 = int(x); y0 = int(y)
                x1 = int(min(bin_vis.shape[1], x0 + self.app.icon_mask_w))
                y1 = int(min(bin_vis.shape[0], y0 + self.app.icon_mask_h))
                cv2.rectangle(bin_vis, (x0, y0), (x1, y1), (200,), thickness=-1)

            prev = self.app.make_preview(frame, bin_vis, rects)
            if prev is None:
                return

            max_w, max_h = 880, 260
            w, h = prev.size
            sc = min(max_w / w, max_h / h)
            sc = max(0.01, sc)
            img = prev.resize((int(w * sc), int(h * sc)))

            tkimg = ImageTk.PhotoImage(img)
            self.iconmask_preview_lbl.config(image=tkimg)
            self.iconmask_preview_lbl.image = tkimg

            self.iconmask_info.config(text=f"検出矩形: {len(rects)} 個 / マスク: {self.app.icon_mask_w}px × {self.app.icon_mask_h}px")
        except Exception as e:
            self.iconmask_info.config(text=f"テスト失敗: {e}")
    def _build_tab_dedupe(self):
        card = ttk.Frame(self.tab_dedupe, style="Card.TFrame", padding=(16, 14))
        card.pack(fill="both", expand=True, padx=14, pady=14)

        ttk.Label(card, text="重複判定（同じ発言の弾き方）", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(card, text="OCRのブレがある場合は「類似度」を使うと重複が減ります（誤って別発言を弾く可能性もあります）。",
                  style="Hint.TLabel").grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        self.var_use_sim = tk.BooleanVar(value=self.app.use_similarity_dedupe)
        ttk.Checkbutton(card, text="類似度で重複を判定する", variable=self.var_use_sim,
                        command=self._on_toggle_similarity).grid(row=2, column=0, columnspan=2, sticky="w", pady=(14, 0))

        ttk.Label(card, text="類似度しきい値", style="Body.TLabel").grid(row=3, column=0, sticky="w", pady=(14, 0))
        self.sim_val = ttk.Label(card, text=f"{self.app.similarity_threshold:.2f}", style="Hint.TLabel")
        self.sim_val.grid(row=3, column=1, sticky="e", pady=(14, 0))

        self.sim_scale = ttk.Scale(card, from_=0.80, to=1.00, value=self.app.similarity_threshold, command=self._on_similarity_change)
        self.sim_scale.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(4, 0))

        ttk.Label(card, text="目安：0.96〜0.99（高いほど厳しく、低いほど重複と見なしやすい）", style="Hint.TLabel")\
            .grid(row=5, column=0, columnspan=2, sticky="w", pady=(6, 0))

        ttk.Separator(card).grid(row=6, column=0, columnspan=2, sticky="ew", pady=14)

        ttk.Label(card, text="比較する履歴件数", style="Body.TLabel").grid(row=7, column=0, sticky="w")
        self.hist_val = ttk.Label(card, text=str(self.app.dedupe_history), style="Hint.TLabel")
        self.hist_val.grid(row=7, column=1, sticky="e")

        self.hist_scale = ttk.Scale(card, from_=50, to=800, value=self.app.dedupe_history, command=self._on_history_change)
        self.hist_scale.grid(row=8, column=0, columnspan=2, sticky="ew", pady=(4, 0))

        ttk.Label(card, text="履歴を増やすと重複判定は安定しますが、少しだけ負荷が増えます。", style="Hint.TLabel")\
            .grid(row=9, column=0, columnspan=2, sticky="w", pady=(6, 0))

        card.columnconfigure(0, weight=1)
        card.columnconfigure(1, weight=0)

        self._refresh_dedupe_controls()

    def _refresh_dedupe_controls(self):
        state = "normal" if self.var_use_sim.get() else "disabled"
        try:
            self.sim_scale.configure(state=state)
        except Exception:
            pass
    def _build_tab_notify(self):
        card = ttk.Frame(self.tab_notify, style="Card.TFrame", padding=(16, 14))
        card.pack(fill="both", expand=True, padx=14, pady=14)

        # Title
        top = ttk.Frame(card, style="Card.TFrame")
        top.pack(fill="x")
        ttk.Label(top, text="通知", style="Title.TLabel").pack(side="left", anchor="w")
        ttk.Label(top, text="未指定なら「ピッ」になります（WAV/MP3対応・WAV推奨）", style="Hint.TLabel").pack(side="right", anchor="e")

        ttk.Label(card, text="チャットが追加されたときに音を鳴らします。スコープ別にON/OFFと音声を設定できます。",
                  style="Hint.TLabel").pack(anchor="w", pady=(8, 12))

        # Global enable (pill-like row)
        enable_row = ttk.Frame(card, style="Card.TFrame")
        enable_row.pack(fill="x", pady=(0, 12))

        self.var_enable_sound = tk.BooleanVar(value=self.app.sound_enabled)
        ttk.Checkbutton(enable_row, text="通知音を有効にする", variable=self.var_enable_sound,
                        command=self._sync_sound_enable).pack(side="left")

        ttk.Button(enable_row, text="テスト（ピッ）", style="Secondary.TButton",
                   command=lambda: play_sound_file(None, root=self.app.root)).pack(side="right")

        ttk.Separator(card).pack(fill="x", pady=(0, 12))

        # Scope settings (cool rows)
        ttk.Label(card, text="スコープ別", style="Title.TLabel").pack(anchor="w", pady=(0, 8))

        self.scope_vars = {}
        self.scope_path_labels = {}

        scopes = ["[ワールド]", "[ギルド]", "[パーティ]", "[チャネル]"]
        for scope in scopes:
            row = ttk.Frame(card, style="Card.TFrame")
            row.pack(fill="x", pady=6)

            var = tk.BooleanVar(value=self.app.sound_scope_enabled.get(scope, True))
            self.scope_vars[scope] = var

            left = ttk.Frame(row, style="Card.TFrame")
            left.pack(side="left", fill="x", expand=True)

            ttk.Checkbutton(left, text=f"{scope}", variable=var,
                            command=lambda s=scope: self._sync_scope_toggle(s)).pack(side="left")

            path_text = self.app.sound_scope_file.get(scope, "") or "（未指定：ピッ）"
            path_lbl = ttk.Label(left, text=path_text, style="Hint.TLabel")
            path_lbl.pack(side="left", padx=(10, 0))
            self.scope_path_labels[scope] = path_lbl

            right = ttk.Frame(row, style="Card.TFrame")
            right.pack(side="right")

            ttk.Button(right, text="音声…", style="Secondary.TButton",
                       command=lambda s=scope: self._pick_scope_file(s)).pack(side="left", padx=(0, 8))
            ttk.Button(right, text="クリア", style="Secondary.TButton",
                       command=lambda s=scope: self._clear_scope_file(s)).pack(side="left", padx=(0, 8))
            ttk.Button(right, text="テスト", style="Secondary.TButton",
                       command=lambda s=scope: self._test_scope(s)).pack(side="left")

        ttk.Separator(card).pack(fill="x", pady=(14, 10))

        note = "※ 音声ファイルは WAV/MP3 に対応します（WAV推奨）。"
        ttk.Label(card, text=note, style="Hint.TLabel").pack(anchor="w")

    def _build_tab_keywords(self):
        card = ttk.Frame(self.tab_keywords, style="Card.TFrame", padding=(16, 14))
        card.pack(fill="both", expand=True, padx=14, pady=14)

        ttk.Label(card, text="キーワード通知", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(card, text="内容に特定の文字列が含まれたら通知音を鳴らします（優先：キーワード ＞ スコープ）。", style="Hint.TLabel")\
            .grid(row=1, column=0, columnspan=4, sticky="w", pady=(6, 12))
        # keyword matching options
        self.var_kana_variants = tk.BooleanVar(value=bool(getattr(self.app, "keyword_allow_kana_variants", False)))
        ttk.Checkbutton(
            card,
            text="ひらがな/カタカナの表記ゆれを許可",
            variable=self.var_kana_variants,
            command=self._sync_kana_variants
        ).grid(row=2, column=0, columnspan=4, sticky="w", pady=(0, 10))


        entry_row = ttk.Frame(card, style="Card.TFrame")
        entry_row.grid(row=3, column=0, columnspan=4, sticky="ew")

        self.keyword_var = tk.StringVar(value="")
        ttk.Entry(entry_row, textvariable=self.keyword_var).pack(side="left", fill="x", expand=True)
        ttk.Button(entry_row, text="追加", style="Secondary.TButton", command=self._add_keyword).pack(side="left", padx=(8, 0))

        self.tree = ttk.Treeview(card, columns=("enabled", "keyword", "sound"), show="headings", height=10, style="KW.Treeview")
        self.tree.grid(row=4, column=0, columnspan=4, sticky="nsew", pady=(12, 8))
        self.tree.heading("enabled", text="有効")
        self.tree.heading("keyword", text="キーワード")
        self.tree.heading("sound", text="音声")
        self.tree.column("enabled", width=70, anchor="center")
        self.tree.column("keyword", width=240, anchor="w")
        self.tree.column("sound", width=360, anchor="w")

        sb = ttk.Scrollbar(card, orient="vertical", command=self.tree.yview)
        sb.grid(row=3, column=4, sticky="ns")
        self.tree.configure(yscrollcommand=sb.set)

        btn_row = ttk.Frame(card, style="Card.TFrame")
        btn_row.grid(row=4, column=0, columnspan=4, sticky="ew")

        ttk.Button(btn_row, text="有効/無効", style="Secondary.TButton", command=self._toggle_keyword).pack(side="left")
        ttk.Button(btn_row, text="音声ファイル…", style="Secondary.TButton", command=self._pick_keyword_sound).pack(side="left", padx=(8, 0))
        ttk.Button(btn_row, text="音をクリア", style="Secondary.TButton", command=self._clear_keyword_sound).pack(side="left", padx=(8, 0))
        ttk.Button(btn_row, text="削除", style="Secondary.TButton", command=self._remove_keyword).pack(side="right")

        ttk.Button(btn_row, text="テスト", style="Secondary.TButton", command=self._test_keyword).pack(side="right", padx=(8, 0))

        card.rowconfigure(4, weight=1)
        card.columnconfigure(0, weight=1)

        self._refresh_keyword_tree()

    # ----- Notify tab callbacks -----
    def _sync_sound_enable(self):
        self.app.sound_enabled = bool(self.var_enable_sound.get())
        self.app.save_settings()

    def _sync_kana_variants(self):
        try:
            self.app.keyword_allow_kana_variants = bool(self.var_kana_variants.get())
            self.app.save_settings()
        except Exception:
            pass

    def _sync_scope_toggle(self, scope: str):
        self.app.sound_scope_enabled[scope] = bool(self.scope_vars[scope].get())
        self.app.save_settings()

    def _pick_scope_file(self, scope: str):
        path = filedialog.askopenfilename(
            title=f"{scope} の通知音を選択",
            filetypes=[("音声ファイル", "*.wav *.mp3"), ("WAV 音声", "*.wav"), ("MP3 音声", "*.mp3"), ("すべてのファイル", "*.*")]
        )
        if not path:
            return
        self.app.sound_scope_file[scope] = path
        self.scope_path_labels[scope].config(text=path)
        self.app.save_settings()

    def _clear_scope_file(self, scope: str):
        self.app.sound_scope_file[scope] = ""
        self.scope_path_labels[scope].config(text="（未指定：ピッ）")
        self.app.save_settings()

    def _test_scope(self, scope: str):
        play_sound_file(self.app.sound_scope_file.get(scope) or None, root=self.app.root)

    # ----- Keyword tab callbacks -----
    def _refresh_keyword_tree(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        for idx, rule in enumerate(self.app.keyword_rules):
            enabled = "✓" if rule.get("enabled", True) else ""
            kw = rule.get("keyword", "")
            sound = rule.get("sound", "") or "（未指定：ピッ）"
            self.tree.insert("", "end", iid=str(idx), values=(enabled, kw, sound))

    def _add_keyword(self):
        kw = (self.keyword_var.get() or "").strip()
        if not kw:
            return
        self.app.keyword_rules.append({"keyword": kw, "sound": "", "enabled": True})
        self.app.save_settings()
        self.keyword_var.set("")
        self._refresh_keyword_tree()

    def _selected_keyword_index(self):
        sel = self.tree.selection()
        if not sel:
            return None
        try:
            return int(sel[0])
        except Exception:
            return None

    def _toggle_keyword(self):
        idx = self._selected_keyword_index()
        if idx is None:
            return
        try:
            self.app.keyword_rules[idx]["enabled"] = not bool(self.app.keyword_rules[idx].get("enabled", True))
            self.app.save_settings()
            self._refresh_keyword_tree()
            self.tree.selection_set(str(idx))
        except Exception:
            pass

    def _pick_keyword_sound(self):
        idx = self._selected_keyword_index()
        if idx is None:
            return
        path = filedialog.askopenfilename(
            title="キーワード通知音を選択",
            filetypes=[("音声ファイル", "*.wav *.mp3"), ("WAV 音声", "*.wav"), ("MP3 音声", "*.mp3"), ("すべてのファイル", "*.*")]
        )
        if not path:
            return
        self.app.keyword_rules[idx]["sound"] = path
        self.app.save_settings()
        self._refresh_keyword_tree()
        self.tree.selection_set(str(idx))

    def _clear_keyword_sound(self):
        idx = self._selected_keyword_index()
        if idx is None:
            return
        self.app.keyword_rules[idx]["sound"] = ""
        self.app.save_settings()
        self._refresh_keyword_tree()
        self.tree.selection_set(str(idx))

    def _remove_keyword(self):
        idx = self._selected_keyword_index()
        if idx is None:
            return
        try:
            self.app.keyword_rules.pop(idx)
            self.app.save_settings()
            self._refresh_keyword_tree()
        except Exception:
            pass

    def _test_keyword(self):
        idx = self._selected_keyword_index()
        if idx is None:
            return
        sound = self.app.keyword_rules[idx].get("sound") or None
        play_sound_file(sound, root=self.app.root)

    # ----- OCR tab callbacks -----
    def _on_thr(self, val):
        try:
            self.app.current_threshold = int(float(val))
            self.thr_val.config(text=str(self.app.current_threshold))
            self.app.save_settings()
        except Exception:
            pass

    def _on_int(self, val):
        try:
            self.app.current_interval = float(val)
            self.int_val.config(text=f"{self.app.current_interval:.1f}")
            self.app.save_settings()
        except Exception:
            pass





    def _on_toggle_similarity(self):
        try:
            self.app.use_similarity_dedupe = bool(self.var_use_sim.get())
            self.app.save_settings()
            self._refresh_dedupe_controls()
        except Exception:
            pass

    def _on_similarity_change(self, val):
        try:
            self.app.similarity_threshold = float(val)
            if hasattr(self, "sim_val"):
                self.sim_val.config(text=f"{self.app.similarity_threshold:.2f}")
            self.app.save_settings()
        except Exception:
            pass

    def _on_history_change(self, val):
        try:
            self.app.dedupe_history = int(float(val))
            if hasattr(self, "hist_val"):
                self.hist_val.config(text=str(self.app.dedupe_history))
            self.app.save_settings()
        except Exception:
            pass

    def _on_destroy(self, _evt=None):
        try:
            if getattr(self.app, 'settings_win', None) is self:
                self.app.settings_win = None
        except Exception:
            pass


    def _build_tab_window(self):
        card = ttk.Frame(self.tab_window, style="Card.TFrame", padding=(16, 14))
        card.pack(fill="both", expand=True, padx=14, pady=14)

        ttk.Label(card, text="ウィンドウ表示", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(card, text="常に前面表示や透明度を調整できます。", style="Hint.TLabel").grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(6, 0)
        )

        self.var_topmost = tk.BooleanVar(value=bool(getattr(self.app, "always_on_top", True)))
        ttk.Checkbutton(
            card,
            text="常に前面に表示する（最前面）",
            variable=self.var_topmost,
            command=self._on_toggle_topmost
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(14, 0))

        ttk.Label(card, text="透明度", style="Body.TLabel").grid(row=3, column=0, sticky="w", pady=(16, 0))
        try:
            init_alpha = float(getattr(self.app, "window_alpha", 1.0))
        except Exception:
            init_alpha = 1.0
        init_pct = int(max(30, min(100, round(init_alpha * 100))))
        self.var_alpha_pct = tk.DoubleVar(value=float(init_pct))

        self.lbl_alpha = ttk.Label(card, text=f"{init_pct}%", style="Hint.TLabel")
        self.lbl_alpha.grid(row=3, column=2, sticky="w", padx=(10, 0), pady=(16, 0))

        scale = ttk.Scale(
            card,
            from_=30,
            to=100,
            orient="horizontal",
            variable=self.var_alpha_pct,
            command=self._on_alpha_change
        )
        scale.grid(row=3, column=1, sticky="ew", padx=(12, 0), pady=(16, 0))
        try:
            scale.set(init_pct)
        except Exception:
            pass

        card.columnconfigure(1, weight=1)

    def _on_toggle_topmost(self):
        try:
            self.app.always_on_top = bool(self.var_topmost.get())
            self.app.apply_window_attributes()
            self.app.save_settings()
        except Exception:
            pass

    def _on_alpha_change(self, val=None):
        try:
            pct = int(round(float(val))) if val is not None else int(round(float(self.var_alpha_pct.get())))
        except Exception:
            pct = 100
        pct = max(30, min(100, pct))

        try:
            # avoid recursive jitter
            if int(round(float(self.var_alpha_pct.get()))) != pct:
                self.var_alpha_pct.set(float(pct))
        except Exception:
            pass

        try:
            if getattr(self, "lbl_alpha", None) is not None:
                self.lbl_alpha.config(text=f"{pct}%")
        except Exception:
            pass

        try:
            self.app.window_alpha = pct / 100.0
            self.app.apply_window_attributes()
            self.app.save_settings()
        except Exception:
            pass

    def _build_tab_logs(self):
        card = ttk.Frame(self.tab_logs, style="Card.TFrame", padding=(16, 14))
        card.pack(fill="both", expand=True, padx=14, pady=14)

        ttk.Label(card, text="ログ表示", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(card, text="『全ログ』タブに溜まるワールドチャットを自動で間引きます。通知/キーワードのログは残します。",
                  style="Hint.TLabel").grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 0))

        self.var_world_ret = tk.BooleanVar(value=getattr(self.app, "world_retention_enabled", True))
        ttk.Checkbutton(card, text="ワールドチャットを一定時間で自動削除する（全ログのみ）",
                        variable=self.var_world_ret, command=self._on_toggle_world_ret).grid(row=2, column=0, columnspan=3, sticky="w", pady=(14, 0))

        ttk.Label(card, text="保持時間（分）", style="Body.TLabel").grid(row=3, column=0, sticky="w", pady=(14, 0))
        self.var_world_ret_min = tk.IntVar(value=int(getattr(self.app, "world_retention_minutes", 5) or 5))

        spin = ttk.Spinbox(card, from_=1, to=120, width=6, textvariable=self.var_world_ret_min, command=self._on_world_ret_min_change)
        spin.grid(row=3, column=1, sticky="w", padx=(10, 0), pady=(14, 0))
        ttk.Label(card, text="分（例：5）", style="Hint.TLabel").grid(row=3, column=2, sticky="w", padx=(10, 0), pady=(14, 0))

        def _commit(_=None):
            self._on_world_ret_min_change()

        spin.bind("<Return>", _commit)
        spin.bind("<FocusOut>", _commit)

        sep = ttk.Separator(card, orient="horizontal")
        sep.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(18, 12))

        ttk.Label(card, text="ログ表示", style="Body.TLabel").grid(row=5, column=0, sticky="w")

        ttk.Label(card, text="文字サイズ", style="Body.TLabel").grid(row=6, column=0, sticky="w")
        self.var_log_font_size = tk.IntVar(value=int(getattr(self.app, "log_font_size", 10) or 10))
        spin_fs = ttk.Spinbox(card, from_=8, to=16, width=6, textvariable=self.var_log_font_size, command=self._on_log_style_change)
        spin_fs.grid(row=6, column=1, sticky="w", padx=(10, 0))
        ttk.Label(card, text="（本文）", style="Hint.TLabel").grid(row=6, column=2, sticky="w", padx=(10, 0))

        ttk.Label(card, text="行間（px）", style="Body.TLabel").grid(row=7, column=0, sticky="w", pady=(10, 0))
        self.var_log_line_spacing = tk.IntVar(value=int(getattr(self.app, "log_line_spacing", 0) or 0))
        spin_ls = ttk.Spinbox(card, from_=0, to=10, width=6, textvariable=self.var_log_line_spacing, command=self._on_log_style_change)
        spin_ls.grid(row=7, column=1, sticky="w", padx=(10, 0), pady=(10, 0))
        ttk.Label(card, text="（0で詰める）", style="Hint.TLabel").grid(row=7, column=2, sticky="w", padx=(10, 0), pady=(10, 0))

        ttk.Label(card, text="カード間隔（px）", style="Body.TLabel").grid(row=8, column=0, sticky="w", pady=(10, 0))
        _gap0 = getattr(self.app, "log_card_gap", None)
        try:
            _gap0 = int(_gap0) if _gap0 is not None else 2
        except Exception:
            _gap0 = 2
        self.var_log_card_gap = tk.IntVar(value=_gap0)
        spin_gap = ttk.Spinbox(card, from_=0, to=20, width=6, textvariable=self.var_log_card_gap, command=self._on_log_style_change)
        spin_gap.grid(row=8, column=1, sticky="w", padx=(10, 0), pady=(10, 0))
        ttk.Label(card, text="（小さいほど詰める）", style="Hint.TLabel").grid(row=8, column=2, sticky="w", padx=(10, 0), pady=(10, 0))

        ttk.Label(card, text="表示形式", style="Body.TLabel").grid(row=9, column=0, sticky="w", pady=(10, 0))
        self.var_log_display_format = tk.StringVar(value=str(getattr(self.app, "log_display_format", "full") or "full"))
        fmt_fr = ttk.Frame(card)
        fmt_fr.grid(row=9, column=1, columnspan=2, sticky="w", padx=(10, 0), pady=(10, 0))
        ttk.Radiobutton(fmt_fr, text="標準（[時刻] 【範囲】 発言者: 本文）", value="full", variable=self.var_log_display_format, command=self._on_log_style_change).pack(anchor="w")
        ttk.Radiobutton(fmt_fr, text="簡易（発言者:本文のみ）", value="simple", variable=self.var_log_display_format, command=self._on_log_style_change).pack(anchor="w", pady=(4, 0))

        def _commit_log(_=None):
            self._on_log_style_change()

        spin_fs.bind("<Return>", _commit_log)
        spin_fs.bind("<FocusOut>", _commit_log)
        spin_ls.bind("<Return>", _commit_log)
        spin_ls.bind("<FocusOut>", _commit_log)
        spin_gap.bind("<Return>", _commit_log)
        spin_gap.bind("<FocusOut>", _commit_log)

    def _on_toggle_world_ret(self):
        try:
            self.app.world_retention_enabled = bool(self.var_world_ret.get())
            self.app.save_settings()
        except Exception:
            pass

    def _on_world_ret_min_change(self):
        try:
            v = int(self.var_world_ret_min.get())
            v = max(1, min(120, v))
            self.var_world_ret_min.set(v)
            self.app.world_retention_minutes = v
            self.app.save_settings()
        except Exception:
            pass



    def _on_log_style_change(self):
        try:
            fs = int(self.var_log_font_size.get()) if hasattr(self, "var_log_font_size") else 10
            ls = int(self.var_log_line_spacing.get()) if hasattr(self, "var_log_line_spacing") else 0
            gap = int(self.var_log_card_gap.get()) if hasattr(self, "var_log_card_gap") else int(getattr(self.app, "log_card_gap", 2))
            fmt = str(self.var_log_display_format.get()) if hasattr(self, "var_log_display_format") else str(getattr(self.app, "log_display_format", "full") or "full")

            if fs < 8: fs = 8
            if fs > 16: fs = 16
            if ls < 0: ls = 0
            if ls > 10: ls = 10
            if gap < 0: gap = 0
            if gap > 20: gap = 20

            fmt = fmt.strip().lower()
            if fmt not in ("full", "simple"):
                fmt = "full"

            if hasattr(self, "var_log_font_size"):
                self.var_log_font_size.set(fs)
            if hasattr(self, "var_log_line_spacing"):
                self.var_log_line_spacing.set(ls)
            if hasattr(self, "var_log_card_gap"):
                self.var_log_card_gap.set(gap)
            if hasattr(self, "var_log_display_format"):
                self.var_log_display_format.set(fmt)

            old_fmt = str(getattr(self.app, "log_display_format", "full") or "full").strip().lower()
            self.app.log_font_size = fs
            self.app.log_line_spacing = ls
            self.app.log_card_gap = gap
            self.app.log_display_format = fmt

            try:
                if fmt != old_fmt:
                    self.app.rebuild_log_cards()
            except Exception:
                pass

            try:
                self.app.refresh_log_styles()
            except Exception:
                pass
            self.app.save_settings()
        except Exception:
            pass
    def on_close(self):
        try:
            self.app.save_settings()
        except Exception:
            pass
        # grabを解除してメインに戻す
        try:
            self.grab_release()
        except Exception:
            pass
        try:
            # 参照を解除
            if getattr(self.app, 'settings_win', None) is self:
                self.app.settings_win = None
        except Exception:
            pass
        try:
            self.destroy()
        except Exception:
            pass


# =========================
# Main app
# =========================
class ChatMonitorApp:

    # 発言者行（1行目）左端のアイコンをマスクする幅（px, 1.0倍率時の目安）
    SPEAKER_ICON_MASK_PX = 44

    # pastel background + dark border
    SCOPE_STYLE = {
        "[ワールド]": {"bg": "#E9D5FF", "border": "#7C3AED"},
        "[ギルド]": {"bg": "#ECFCCB", "border": "#65A30D"},
        "[パーティ]": {"bg": "#DBEAFE", "border": "#1D4ED8"},
        "[チャネル]": {"bg": "#F3F4F6", "border": "#64748B"},
    }

    def __init__(self, root: tk.Tk):
        self.root = root
        # Main preview sizing
        self.preview_width_ratio = 0.80  # preview frame width relative to the preview card
        # Log display style
        self.log_font_size = 8
        self.log_line_spacing = 0  # px; 0 = tight
        self.log_card_gap = 0  # px; card vertical gap (smaller = tighter)
        self.log_display_format = "simple"  # "full" or "simple"
        self.main_view_mode = "full"  # "full" or "log_only"

        self.preview_target_w = 320      # updated on resize
        self.preview_collapsed = False  # preview折りたたみ状態

        self.root.title("スタレゾ チャット解析ツール")
        self.root.geometry("360x760+-4+1")
        self.root.minsize(360, 760)

        # window behavior
        self.always_on_top = True
        self.window_alpha = 0.95  # 0.30 - 1.00


        # monitoring settings
        self.monitoring = False
        self.area = (59, 829, 290, 275)
        self.last_keys = []
        self.last_msgs = []  # recent messages for similarity dedupe
        self.use_similarity_dedupe = True
        self.similarity_threshold = 0.96  # 0.80〜1.00
        self.dedupe_history = 220
        self.current_threshold = 220
        self.ocr_scale = 1.0

        # icon mask (per-message rectangle) settings
        self.icon_mask_enabled = True
        self.icon_mask_w = 24  # px in original image
        self.icon_mask_h = 26  # px in original image
        self.rect_kernel_w = 60
        self.rect_kernel_h = 18
        self.rect_min_area = 450
        self.rect_merge_gap = 10
        # bubble detection (recommended): detect chat bubble backgrounds and use as message rectangles
        self.use_bubble_rects = True
        self.bubble_v_min = 190      # HSV V threshold (bright)
        self.bubble_s_max = 90       # HSV S threshold (low saturation)
        self.bubble_kernel_w = 35    # morphology close kernel width
        self.bubble_kernel_h = 21    # morphology close kernel height
        self.bubble_min_area = 1600  # minimum bubble area
        self.bubble_pad = 2          # expand detected rect by px
        self.last_rects = []
        self.last_area_snapshot = None  # PIL.Image
        self.current_interval = 1.0
        self.stop_event = threading.Event()
        self.ui_queue: "queue.Queue[tuple]" = queue.Queue(maxsize=10)
        self.seed_dedupe_on_next = False  # when True, seed dedupe from current screen without logging (used after Clear)

        # message storage
        self.max_buffer_items = 2000  # copy/export buffer cap (prevents memory growth)
        self.message_buffer = deque(maxlen=self.max_buffer_items)  # (t_str, scope, speaker, content, ts_epoch, notify_scope, hit_kw)
        # ワールドログ自動削除（全ログタブのみ）。通知/キーワードタブは保持します。
        self.world_retention_enabled = True
        self.world_retention_minutes = 5  # 分（デフォルト）
        self.log_items = []
        self.log_items_notify = []  # 通知（スコープ）用
        self.log_items_keyword = []  # キーワードヒット用
        self.log_scroller_notify = None
        self.log_scroller_keyword = None
        self.log_notebook = None
        self.max_log_items = 400


        # capture persistence
        self.remember_area = True  # 読み取り範囲を次回起動時に復元する
        # sound settings
        self.sound_enabled = True
        self.sound_scope_enabled = {
            "[ワールド]": False,
            "[ギルド]": True,
            "[パーティ]": False,
            "[チャネル]": False,
        }
        self.sound_scope_file = {
            "[ワールド]": "",
            "[ギルド]": "",
            "[パーティ]": "",
            "[チャネル]": "",
        }
        self.keyword_rules = []  # list of dict {keyword, sound, enabled}


        # keyword match options
        self.keyword_ignore_spaces = True  # キーワード途中に空白が入っても検知
        self.keyword_allow_kana_variants = True  # ひらがな/カタカナ表記ゆれを許可
        self._last_sound_time = 0.0
        self._sound_min_interval = 0.25  # anti-spam

        # load persisted settings (if any)
        self.load_settings()

        # apply window attributes (topmost/alpha)
        self.apply_window_attributes()

        self._setup_style()
        self._load_toolbar_icons()
        self._build_ui()

        self.root.after(30, self.process_ui_queue)
        # periodic GC for long sessions
        self.root.after(600_000, self._periodic_gc)
        # ワールドログの自動削除（全ログタブのみ）
        self.root.after(15_000, self._world_retention_tick)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self._set_status("")


    # ---------- Config persistence ----------
    def load_settings(self):
        """
        Load settings from CONFIG_PATH. Missing/invalid values are ignored safely.
        """
        try:
            data = None
            if os.path.exists(CONFIG_PATH):
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    data = json.load(f)
            else:
                # Seed with defaults and create config file on first run
                data = json.loads(json.dumps(DEFAULT_CONFIG, ensure_ascii=False))
                try:
                    with open(CONFIG_PATH, "w", encoding="utf-8") as wf:
                        json.dump(data, wf, ensure_ascii=False, indent=2)
                except Exception:
                    pass


            # window geometry (size/position)
            try:
                win = data.get("window", {})
                if isinstance(win, dict):
                    geom = (win.get("geometry") or "").strip()
                    # expected like "1020x900+10+10" (allow negative)
                    if geom and re.match(r"^\d+x\d+[+-]\d+[+-]\d+$", geom):
                        self.root.geometry(geom)
                        try:
                            self.root.update_idletasks()
                        except Exception:
                            pass

                    # topmost / alpha (window appearance)
                    try:
                        topmost = win.get("topmost", None)
                        if isinstance(topmost, bool):
                            self.always_on_top = topmost
                        alpha = win.get("alpha", None)
                        if isinstance(alpha, (int, float)):
                            self.window_alpha = max(0.30, min(1.0, float(alpha)))
                    except Exception:
                        pass
                    try:
                        self.apply_window_attributes()
                    except Exception:
                        pass
            except Exception:
                pass

            # remember area
            self.remember_area = bool(data.get("remember_area", True))

            # capture area
            area = data.get("area")
            if self.remember_area and isinstance(area, (list, tuple)) and len(area) == 4:
                try:
                    x, y, w, h = [int(v) for v in area]
                    if w > 10 and h > 10:
                        self.area = (x, y, w, h)
                except Exception:
                    pass

            ocr = data.get("ocr", {})
            if isinstance(ocr, dict):
                thr = ocr.get("threshold")
                itv = ocr.get("interval")
                if isinstance(thr, (int, float)):
                    self.current_threshold = int(max(0, min(255, int(thr))))
                if isinstance(itv, (int, float)):
                    self.current_interval = float(max(0.3, min(10.0, float(itv))))

                sc = ocr.get("scale")
                if isinstance(sc, (int, float)):
                    self.ocr_scale = float(max(1.0, min(3.0, float(sc))))


            # dedupe
            dd = data.get("dedupe", {})
            if isinstance(dd, dict):
                self.use_similarity_dedupe = bool(dd.get("use_similarity", self.use_similarity_dedupe))
                st = dd.get("threshold")
                hist = dd.get("history")
                if isinstance(st, (int, float)):
                    self.similarity_threshold = float(max(0.80, min(1.00, float(st))))
                if isinstance(hist, (int, float)):
                    self.dedupe_history = int(max(50, min(800, int(hist))))

            
            # icon mask
            im = data.get("iconmask", {})
            if isinstance(im, dict):
                self.icon_mask_enabled = bool(im.get("enabled", self.icon_mask_enabled))
                for k, attr, lo, hi in [
                    ("mask_w", "icon_mask_w", 10, 200),
                    ("mask_h", "icon_mask_h", 10, 120),
                    ("kernel_w", "rect_kernel_w", 10, 200),
                    ("kernel_h", "rect_kernel_h", 6, 80),
                    ("min_area", "rect_min_area", 50, 5000),
                    ("merge_gap", "rect_merge_gap", 0, 60),
                    ("use_bubble", "use_bubble_rects", 0, 1),
                    ("bubble_v_min", "bubble_v_min", 120, 250),
                    ("bubble_s_max", "bubble_s_max", 0, 140),
                    ("bubble_kernel_w", "bubble_kernel_w", 5, 120),
                    ("bubble_kernel_h", "bubble_kernel_h", 5, 80),
                    ("bubble_min_area", "bubble_min_area", 200, 20000),
                    ("bubble_pad", "bubble_pad", 0, 12),
                ]:
                    v = im.get(k)
                    if isinstance(v, (int, float)):
                        setattr(self, attr, int(max(lo, min(hi, int(v)))))
            # normalize bool stored as 0/1
            self.use_bubble_rects = bool(self.use_bubble_rects)
                        # log retention
            try:
                lg = data.get("log", {})
                if isinstance(lg, dict):
                    self.world_retention_enabled = bool(lg.get("world_retention_enabled", getattr(self, "world_retention_enabled", True)))
                    self.world_retention_minutes = int(lg.get("world_retention_minutes", getattr(self, "world_retention_minutes", 5)))
                    if self.world_retention_minutes <= 0:
                        self.world_retention_minutes = 5
                    # Log card compactness
                    fs = lg.get("font_size", getattr(self, "log_font_size", 10))
                    ls = lg.get("line_spacing", getattr(self, "log_line_spacing", 0))
                    try:
                        fs = int(fs)
                        if fs < 8: fs = 8
                        if fs > 16: fs = 16
                        self.log_font_size = fs
                    except Exception:
                        pass
                    try:
                        ls = int(ls)
                        if ls < 0: ls = 0
                        if ls > 10: ls = 10
                        self.log_line_spacing = ls
                    except Exception:
                        pass
                    # Card gap (vertical) / display format
                    gap = lg.get("card_gap", getattr(self, "log_card_gap", 2))
                    fmt = lg.get("display_format", getattr(self, "log_display_format", "full"))
                    try:
                        gap = int(gap)
                        if gap < 0: gap = 0
                        if gap > 20: gap = 20
                        self.log_card_gap = gap
                    except Exception:
                        pass
                    try:
                        fmt = str(fmt).strip().lower()
                        if fmt not in ("full", "simple"):
                            fmt = "full"
                        self.log_display_format = fmt
                    except Exception:
                        pass

            except Exception:
                pass

            snd = data.get("sound", {})
            if isinstance(snd, dict):
                self.sound_enabled = bool(snd.get("enabled", self.sound_enabled))


                # keyword match options
                try:
                    self.keyword_ignore_spaces = bool(snd.get("keyword_ignore_spaces", getattr(self, "keyword_ignore_spaces", True)))
                    self.keyword_allow_kana_variants = bool(snd.get("kana_variants", getattr(self, "keyword_allow_kana_variants", False)))
                except Exception:
                    pass
                se = snd.get("scope_enabled", {})
                if isinstance(se, dict):
                    for k in list(self.sound_scope_enabled.keys()):
                        if k in se:
                            self.sound_scope_enabled[k] = bool(se.get(k))

                sf = snd.get("scope_file", {})
                if isinstance(sf, dict):
                    for k in list(self.sound_scope_file.keys()):
                        v = sf.get(k)
                        if isinstance(v, str):
                            self.sound_scope_file[k] = v

                rules = snd.get("keyword_rules")
                if isinstance(rules, list):
                    cleaned = []
                    for r in rules:
                        if not isinstance(r, dict):
                            continue
                        kw = (r.get("keyword") or "").strip()
                        if not kw:
                            continue
                        cleaned.append({
                            "keyword": kw,
                            "sound": (r.get("sound") or ""),
                            "enabled": bool(r.get("enabled", True)),
                        })
                    self.keyword_rules = cleaned
        except Exception:
            # ignore broken config
            return


    def apply_window_attributes(self):
        """Apply window topmost/alpha settings safely (main + settings window)."""
        topmost = bool(getattr(self, "always_on_top", False))
        try:
            alpha = float(getattr(self, "window_alpha", 1.0))
        except Exception:
            alpha = 1.0
        if alpha < 0.30:
            alpha = 0.30
        if alpha > 1.0:
            alpha = 1.0

        try:
            self.root.attributes("-topmost", topmost)
        except Exception:
            pass
        try:
            self.root.attributes("-alpha", alpha)
        except Exception:
            pass

        # settings window should follow the same behavior (if open)
        try:
            win = getattr(self, "settings_win", None)
            if win is not None and win.winfo_exists():
                try:
                    win.attributes("-topmost", topmost)
                except Exception:
                    pass
                try:
                    win.attributes("-alpha", alpha)
                except Exception:
                    pass
        except Exception:
            pass

    def save_settings(self):
        """
        Save settings to CONFIG_PATH. Area is saved only if remember_area is True.
        """
        try:
            data = {
                "version": 3,
                "remember_area": bool(self.remember_area),
                "area": None,
                "window": {"geometry": (self.root.geometry() if self.root else ""), "topmost": bool(getattr(self, "always_on_top", False)), "alpha": float(getattr(self, "window_alpha", 1.0))},
                "ocr": {
                    "threshold": int(self.current_threshold),
                    "interval": float(self.current_interval),
                    "scale": float(self.ocr_scale),
                },
                "dedupe": {
                    "use_similarity": bool(self.use_similarity_dedupe),
                    "threshold": float(self.similarity_threshold),
                    "history": int(self.dedupe_history),
                },
                "iconmask": {
                    "enabled": bool(self.icon_mask_enabled),
                    "mask_w": int(self.icon_mask_w),
                    "mask_h": int(self.icon_mask_h),
                    "kernel_w": int(self.rect_kernel_w),
                    "kernel_h": int(self.rect_kernel_h),
                    "min_area": int(self.rect_min_area),
                    "merge_gap": int(self.rect_merge_gap),
                    "use_bubble": bool(self.use_bubble_rects),
                    "bubble_v_min": int(self.bubble_v_min),
                    "bubble_s_max": int(self.bubble_s_max),
                    "bubble_kernel_w": int(self.bubble_kernel_w),
                    "bubble_kernel_h": int(self.bubble_kernel_h),
                    "bubble_min_area": int(self.bubble_min_area),
                    "bubble_pad": int(self.bubble_pad),
                },
                "log": {
                    "world_retention_enabled": bool(getattr(self, "world_retention_enabled", True)),
                    "world_retention_minutes": int(getattr(self, "world_retention_minutes", 5)),
                    "font_size": int(getattr(self, "log_font_size", 10)),
                    "line_spacing": int(getattr(self, "log_line_spacing", 0)),
                    "card_gap": int(getattr(self, "log_card_gap", 2)),
                    "display_format": str(getattr(self, "log_display_format", "full")),
                },
                "sound": {
                    "enabled": bool(self.sound_enabled),
                    "keyword_ignore_spaces": bool(getattr(self, "keyword_ignore_spaces", True)),
                    "kana_variants": bool(getattr(self, "keyword_allow_kana_variants", False)),
                    "scope_enabled": dict(self.sound_scope_enabled),
                    "scope_file": dict(self.sound_scope_file),
                    "keyword_rules": list(self.keyword_rules),
                }
            }

            if self.remember_area and self.area and isinstance(self.area, (tuple, list)) and len(self.area) == 4:
                data["area"] = [int(self.area[0]), int(self.area[1]), int(self.area[2]), int(self.area[3])]

            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _load_toolbar_icons(self):
        """Load toolbar icon images if available (icon_setting.png / icon_trash.png)."""
        self.icon_setting_img = None
        self.icon_trash_img = None
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        except Exception:
            base_dir = os.getcwd()
        try:
            # Prefer local files next to chat_tool.py
            p_setting = os.path.join(base_dir, "icon_setting.png")
            p_trash = os.path.join(base_dir, "icon_trash.png")
            resample = getattr(Image, "LANCZOS", getattr(Image, "BICUBIC", 3))
            if os.path.exists(p_setting):
                img = Image.open(p_setting).convert("RGBA")
                img = img.resize((18, 18), resample)
                self.icon_setting_img = ImageTk.PhotoImage(img)
            if os.path.exists(p_trash):
                img = Image.open(p_trash).convert("RGBA")
                img = img.resize((18, 18), resample)
                self.icon_trash_img = ImageTk.PhotoImage(img)
        except Exception:
            # Fallback to text buttons
            self.icon_setting_img = None
            self.icon_trash_img = None

    def _setup_style(self):
        self.root.configure(bg="#F6F8FC")
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        default_font = ("Yu Gothic UI", 10)
        self.style.configure(".", font=default_font)
        self.style.configure("Card.TFrame", background="#FFFFFF")
        self.style.configure("Header.TFrame", background="#F6F8FC")

        self.style.configure("HeaderTitle.TLabel", background="#F6F8FC", foreground="#0F172A", font=("Yu Gothic UI", 16, "bold"))
        self.style.configure("HeaderSub.TLabel", background="#F6F8FC", foreground="#475569", font=("Yu Gothic UI", 10))
        self.style.configure("Title.TLabel", background="#FFFFFF", foreground="#0F172A", font=("Yu Gothic UI", 11, "bold"))
        self.style.configure("Body.TLabel", background="#FFFFFF", foreground="#334155")
        self.style.configure("Hint.TLabel", background="#FFFFFF", foreground="#64748B", font=("Yu Gothic UI", 9))

        self.style.configure("PillOk.TLabel", background="#DCFCE7", foreground="#166534", padding=(10, 3), font=("Yu Gothic UI", 9, "bold"))
        self.style.configure("PillNg.TLabel", background="#FEE2E2", foreground="#991B1B", padding=(10, 3), font=("Yu Gothic UI", 9, "bold"))
        self.style.configure("PillRun.TLabel", background="#DBEAFE", foreground="#1D4ED8", padding=(10, 3), font=("Yu Gothic UI", 9, "bold"))

        self.style.configure("Primary.TButton", padding=(12, 8), font=("Yu Gothic UI", 10, "bold"))
        self.style.map("Primary.TButton",
                       foreground=[("disabled", "#94A3B8"), ("!disabled", "#FFFFFF")],
                       background=[("disabled", "#CBD5E1"), ("!disabled", "#2563EB")])

        self.style.configure("Danger.TButton", padding=(12, 8), font=("Yu Gothic UI", 10, "bold"))
        self.style.map("Danger.TButton",
                       foreground=[("disabled", "#94A3B8"), ("!disabled", "#FFFFFF")],
                       background=[("disabled", "#CBD5E1"), ("!disabled", "#EF4444")])

        self.style.configure("Secondary.TButton", padding=(10, 8), font=("Yu Gothic UI", 10, "bold"))
        self.style.configure("Tiny.TButton", padding=(2, 0), font=("Yu Gothic UI", 10, "bold"))
        self.style.configure("Icon.TButton", padding=(4, 2))
        self.style.map("Icon.TButton",
                       background=[("active", "#E2E8F0"), ("!active", "#E2E8F0")])
        self.style.map("Secondary.TButton",
                       foreground=[("disabled", "#94A3B8"), ("!disabled", "#0F172A")],
                       background=[("disabled", "#E2E8F0"), ("!disabled", "#E2E8F0")])

        # Settings keyword table: increase row height to avoid glyph clipping
        self.style.configure("KW.Treeview", rowheight=28, font=("Yu Gothic UI", 10))
        self.style.configure("KW.Treeview.Heading", font=("Yu Gothic UI", 10, "bold"))

    def _build_ui(self):
        header = ttk.Frame(self.root, padding=(12, 10), style="Header.TFrame")
        header.pack(fill="x")

        title_row = ttk.Frame(header, style="Header.TFrame")
        title_row.pack(fill="x")

        ttk.Label(title_row, text="BPSR ChatCap", style="HeaderTitle.TLabel").pack(side="left", anchor="w")
        toolbar = ttk.Frame(title_row, style="Header.TFrame")
        toolbar.pack(side="right")


        # Header toolbar buttons
        use_icon_style = (getattr(self, "icon_setting_img", None) is not None) or (getattr(self, "icon_trash_img", None) is not None)

        if getattr(self, "icon_setting_img", None) is not None:
            self.btn_header_settings = ttk.Button(toolbar, image=self.icon_setting_img, style="Icon.TButton", command=self.open_settings)
        else:
            self.btn_header_settings = ttk.Button(toolbar, text="設定", style="Secondary.TButton", width=6, command=self.open_settings)
        self.btn_header_settings.pack(side="left", padx=(0, 8))

        if getattr(self, "icon_trash_img", None) is not None:
            self.btn_header_clear = ttk.Button(toolbar, image=self.icon_trash_img, style="Icon.TButton", command=self.clear_log)
        else:
            self.btn_header_clear = ttk.Button(toolbar, text="クリア", style="Secondary.TButton", width=6, command=self.clear_log)
        self.btn_header_clear.pack(side="left")

        # ▲/▼: 表示切り替え（▲=ログのみ表示へ / ▼=全体表示へ）
        if use_icon_style:
            self.btn_view_toggle = ttk.Button(toolbar, text="▲", style="Icon.TButton", command=self.toggle_main_view_mode)
        else:
            # 「設定」「クリア」と同じサイズ感で置く（テキストボタン）
            self.btn_view_toggle = ttk.Button(toolbar, text="▲", style="Secondary.TButton", width=6, command=self.toggle_main_view_mode)
        self.btn_view_toggle.pack(side="left", padx=(8, 0))

        container = ttk.Frame(self.root, padding=(12, 10), style="Header.TFrame")
        container.pack(fill="both", expand=True)

        top = ttk.Frame(container, style="Header.TFrame")
        top.pack(fill="x")
        self.top_frame = top

        # Controls + Preview (stacked for compact width)

        card_controls = ttk.Frame(top, style="Card.TFrame", padding=(12, 10))
        card_controls.pack(fill="x")

        ops_row = ttk.Frame(card_controls, style="Card.TFrame")
        ops_row.pack(fill="x")
        ttk.Label(ops_row, text="操作", style="Title.TLabel").pack(side="left", anchor="w")
        self.status_pill = ttk.Label(ops_row, text="停止中", style="PillNg.TLabel")
        self.status_pill.pack(side="right")

        btns = ttk.Frame(card_controls, style="Card.TFrame")
        btns.pack(fill="x", pady=(12, 8))

        self.btn_start = ttk.Button(btns, text="開始", style="Primary.TButton", command=self.start_monitoring)
        self.btn_start.pack(side="left", padx=(0, 8))

        self.btn_stop = ttk.Button(btns, text="停止", style="Danger.TButton", command=self.stop_monitoring, state="disabled")
        self.btn_stop.pack(side="left")


        card_preview = ttk.Frame(top, style="Card.TFrame", padding=(12, 10))
        card_preview.pack(fill="x", pady=(10, 0))

        preview_head = ttk.Frame(card_preview, style="Card.TFrame")
        preview_head.pack(fill="x")
        ttk.Label(preview_head, text="二値化プレビュー", style="Title.TLabel").pack(side="left", anchor="w")

        self.preview_body = ttk.Frame(card_preview, style="Card.TFrame")
        self.preview_body.pack(fill="x", pady=(6, 0))
        ttk.Label(self.preview_body, text="二値化画像（緑枠＝検出バブル）", style="Hint.TLabel").pack(anchor="w", pady=(0, 8))

        self.preview_wrap = tk.Frame(self.preview_body, bg="#0B1220", highlightthickness=1, highlightbackground="#E2E8F0")
        self.preview_wrap.pack(anchor="center")
        self.preview_wrap.configure(height=190)
        self.preview_wrap.configure(width=self.preview_target_w)
        self.preview_wrap.pack_propagate(False)

        # Preview frame width adapts to the preview card width (keeps window compact)
        card_preview.bind("<Configure>", self._on_preview_card_resize)

        self.preview_label = tk.Label(self.preview_wrap, bg="#0B1220")
        self.preview_label.pack(expand=True)

        card_log = ttk.Frame(container, style="Card.TFrame", padding=(14, 12))
        card_log.pack(fill="both", expand=True, pady=(12, 0))
        self.card_log = card_log
        self._log_full_pady = (12, 0)
        self._log_compact_pady = (0, 0)

        row = ttk.Frame(card_log, style="Card.TFrame")
        row.pack(fill="x")
        ttk.Label(row, text="ログ", style="Title.TLabel").pack(side="left")
        self.area_pill = ttk.Label(row, text="範囲: 未設定", style="PillNg.TLabel")
        self.area_pill.pack(side="right")


                # ログ表示（タブ切り替え）
        # ttk.Notebook はタブごとの背景色変更が難しいため、ログ欄だけカスタムタブにしています。
        self.log_notebook = None  # legacy

        # タブ（ボタン）
        self.log_tab_cur = "all"
        self._tab_alert = {'notify': False, 'keyword': False}
        self._tab_blink_on = False
        self._tab_blink_job = None

        self._tab_colors = {
            'normal': '#F1F5F9',
            'selected': '#FFFFFF',
            'text': '#0F172A',
            'border': '#CBD5E1',
            'border_sel': '#1D4ED8',
            'notify_a': '#DBEAFE',
            'notify_b': '#BFDBFE',
            'kw_a': '#FEE2E2',
            'kw_b': '#FECACA',
        }

        self.log_tabbar = tk.Frame(card_log, bg='#FFFFFF')
        self.log_tabbar.pack(fill='x', pady=(10, 0))

        def _mk_tab(lbl_text, kind):
            w = tk.Label(
                self.log_tabbar,
                text=lbl_text,
                bg=self._tab_colors['normal'],
                fg=self._tab_colors['text'],
                padx=14,
                pady=6,
                cursor='hand2',
                font=('Yu Gothic UI', 10, 'bold')
            )
            w.configure(highlightthickness=1, highlightbackground=self._tab_colors['border'])
            w.bind('<Button-1>', lambda e, k=kind: self._select_log_tab(k))
            return w

        self.log_tab_btn_all = _mk_tab('すべて', 'all')
        self.log_tab_btn_notify = _mk_tab('通知', 'notify')
        self.log_tab_btn_keyword = _mk_tab('キーワード', 'keyword')

        self.log_tab_btn_all.pack(side='left', padx=(0, 6))
        self.log_tab_btn_notify.pack(side='left', padx=(0, 6))
        self.log_tab_btn_keyword.pack(side='left')

        self.log_tab_container = ttk.Frame(card_log, style='Card.TFrame')
        self.log_tab_container.pack(fill='both', expand=True, pady=(8, 0))

        tab_all = ttk.Frame(self.log_tab_container, style='Card.TFrame')
        tab_notify = ttk.Frame(self.log_tab_container, style='Card.TFrame')
        tab_kw = ttk.Frame(self.log_tab_container, style='Card.TFrame')

        for _f in (tab_all, tab_notify, tab_kw):
            _f.place(relx=0, rely=0, relwidth=1, relheight=1)

        self._log_tab_frames = {'all': tab_all, 'notify': tab_notify, 'keyword': tab_kw}
        self._log_tab_buttons = {'all': self.log_tab_btn_all, 'notify': self.log_tab_btn_notify, 'keyword': self.log_tab_btn_keyword}

        # 初期表示
        self._select_log_tab('all', initial=True)


        self.log_scroller = ScrolledFrame(tab_all)
        self.log_scroller.pack(fill="both", expand=True)
        self.log_scroller.canvas.bind("<Configure>", lambda e: self._update_wrap_lengths())

        self.log_scroller_notify = ScrolledFrame(tab_notify)
        self.log_scroller_notify.pack(fill="both", expand=True)
        self.log_scroller_notify.canvas.bind("<Configure>", lambda e: self._update_wrap_lengths())

        self.log_scroller_keyword = ScrolledFrame(tab_kw)
        self.log_scroller_keyword.pack(fill="both", expand=True)
        self.log_scroller_keyword.canvas.bind("<Configure>", lambda e: self._update_wrap_lengths())


        # Apply initial main view mode (full / log_only)
        self._apply_main_view_mode()

        footer = ttk.Frame(self.root, style="Header.TFrame", padding=(18, 8))
        footer.pack(fill="x")
        self.footer_label = ttk.Label(footer, text="—", style="HeaderSub.TLabel")
        self.footer_label.pack(anchor="w")

    def _on_preview_card_resize(self, event=None):
        """Resize the preview frame based on the preview card width (compact-friendly)."""
        try:
            # When preview is collapsed, do nothing (avoid expanding window width).
            if bool(getattr(self, "preview_collapsed", False)):
                return
            if event is None:
                return
            w = getattr(event, "width", None)
            if not w:
                return
            avail = max(160, int(w) - 24)
            target = int(avail * float(getattr(self, "preview_width_ratio", 0.80)))
            target = max(180, min(800, target))
            if abs(target - int(getattr(self, "preview_target_w", target))) < 6:
                return
            self.preview_target_w = target
            try:
                if hasattr(self, "preview_wrap") and self.preview_wrap.winfo_exists():
                    self.preview_wrap.configure(width=self.preview_target_w)
            except Exception:
                pass
        except Exception:
            pass






    def toggle_main_view_mode(self):
        """表示モードをトグルする（full <-> log_only）
        - full のときボタン表示は「▲」= 押すとログのみ
        - log_only のときボタン表示は「▼」= 押すと全体表示
        """
        try:
            mode = getattr(self, "main_view_mode", "full")
            self.main_view_mode = ("full" if mode == "log_only" else "log_only")
            self._apply_main_view_mode()
        except Exception:
            pass

    def _apply_main_view_mode(self):
        """メイン画面の表示モードを適用する（full / log_only）"""
        try:
            mode = getattr(self, "main_view_mode", "full")
            top = getattr(self, "top_frame", None)
            card_log = getattr(self, "card_log", None)

            if mode == "log_only":
                try:
                    if top is not None and top.winfo_exists():
                        top.pack_forget()
                except Exception:
                    pass
                try:
                    if card_log is not None and card_log.winfo_exists():
                        card_log.pack_configure(pady=getattr(self, "_log_compact_pady", (0, 0)))
                except Exception:
                    pass
            else:
                # full
                try:
                    if top is not None and top.winfo_exists():
                        # pack order: place above log
                        if card_log is not None and card_log.winfo_exists():
                            if not top.winfo_ismapped():
                                top.pack(before=card_log, fill="x")
                        else:
                            if not top.winfo_ismapped():
                                top.pack(fill="x")
                except Exception:
                    pass
                try:
                    if card_log is not None and card_log.winfo_exists():
                        card_log.pack_configure(pady=getattr(self, "_log_full_pady", (12, 0)))
                except Exception:
                    pass

                        # update toggle button label
            try:
                if hasattr(self, "btn_view_toggle") and self.btn_view_toggle.winfo_exists():
                    # full表示中は「▲」（押すとログのみ）、ログのみ中は「▼」（押すと全体）
                    self.btn_view_toggle.configure(text=("▼" if mode == "log_only" else "▲"))
            except Exception:
                pass


            try:
                self.root.update_idletasks()
            except Exception:
                pass
        except Exception:
            pass





    def open_settings(self):
        # 設定画面は1つだけ（モーダル）
        try:
            win = getattr(self, 'settings_win', None)
            if win is not None and win.winfo_exists():
                try:
                    win.deiconify()
                except Exception:
                    pass
                try:
                    win.lift()
                    win.focus_force()
                except Exception:
                    pass
                return
        except Exception:
            self.settings_win = None

        self.settings_win = SettingsWindow(self)
        try:
            self.apply_window_attributes()
        except Exception:
            pass


    def _set_status(self, text: str):
        # 上部のボタン下には出さず、フッターのみに表示
        self.footer_label.config(text=text)

        if self.monitoring:
            self.status_pill.config(text="監視中", style="PillRun.TLabel")
        else:
            self.status_pill.config(text="停止中", style="PillNg.TLabel")

        if self.area:
            self.area_pill.config(text="範囲: 設定済み", style="PillOk.TLabel")
        else:
            self.area_pill.config(text="範囲: 未設定", style="PillNg.TLabel")

    def clear_log(self):
        # Clear visible log list (UI)
        for item in self.log_items:
            try:
                item["frame"].destroy()
            except Exception:
                pass
        self.log_items.clear()

        # Clear notify log list (UI)
        if hasattr(self, 'log_items_notify'):
            for item in list(self.log_items_notify):
                try:
                    item['frame'].destroy()
                except Exception:
                    pass
            self.log_items_notify.clear()

        # Clear keyword-hit log list (UI)
        if hasattr(self, 'log_items_keyword'):
            for item in list(self.log_items_keyword):
                try:
                    item['frame'].destroy()
                except Exception:
                    pass
            self.log_items_keyword.clear()

        # Clear copy buffer
        self.message_buffer.clear()

        # Reset dedupe history, but do NOT re-print currently visible chat bubbles.
        # We seed the dedupe history on the next capture cycle.
        self.last_keys.clear()
        self.last_msgs.clear()
        self.seed_dedupe_on_next = bool(self.monitoring)

        self._set_status("ログをクリアしました。")


    # ---------- Scope normalization ----------
    def normalize_scope(self, scope_raw: str) -> str:
        s = (scope_raw or "").strip()
        if not s:
            return "[チャネル]"
        core = s.replace("【", "[").replace("】", "]").strip()
        m = re.search(r"\[(.*?)\]", core)
        inside = (m.group(1) if m else core).strip()
        inside2 = re.sub(r"[0-9０-９\s\-_—‐ー\.\,，。:：;；/\\]+", "", inside)

        up = inside2.upper()
        if "ワールド" in inside2 or "WORLD" in up:
            return "[ワールド]"
        if "ギルド" in inside2:
            return "[ギルド]"
        if "パーティ" in inside2 or "パーテイ" in inside2 or "パ一ティ" in inside2:
            return "[パーティ]"
        if "チャネル" in inside2 or "チャンネル" in inside2 or "チヤネル" in inside2:
            return "[チャネル]"

        if inside2.startswith("ワル") or inside2.startswith("ワー"):
            return "[ワールド]"
        if inside2.startswith("ギル"):
            return "[ギルド]"
        if inside2.startswith("パ"):
            return "[パーティ]"

        return "[チャネル]"

    def _fmt_time_label(self, ts_str: str) -> str:
        s = (ts_str or '').strip()
        if not s:
            return ''
        # Ensure [HH:MM:SS] style
        if s.startswith('[') and s.endswith(']'):
            return s
        return f'[{s}]'

    def _scope_display_name(self, scope_norm: str) -> str:
        s = (scope_norm or '').strip()
        if s.startswith('[') and s.endswith(']') and len(s) >= 2:
            return s[1:-1]
        if s.startswith('【') and s.endswith('】') and len(s) >= 2:
            return s[1:-1]
        return s

    def _fmt_scope_label(self, scope_norm: str) -> str:
        return f"【{self._scope_display_name(scope_norm)}】"

    # ---------- Sound rules ----------
    def _should_rate_limit_sound(self) -> bool:
        now = time.time()
        if now - self._last_sound_time < self._sound_min_interval:
            return True
        self._last_sound_time = now
        return False

    def maybe_play_notification(self, scope: str, speaker: str, content: str):
        if not self.sound_enabled:
            return
        if self._should_rate_limit_sound():
            return

        scope_norm = self.normalize_scope(scope)

        # Priority 1: keyword rules
        hay = f"{speaker}\n{content}"
        hay_cmp = self.normalize_for_keyword_match(hay)
        for rule in self.keyword_rules:
            if not rule.get("enabled", True):
                continue
            kw = (rule.get("keyword") or "").strip()
            if not kw:
                continue
            kw_cmp = self.normalize_for_keyword_match(kw)
            if kw_cmp and kw_cmp in hay_cmp:
                play_sound_file(rule.get("sound") or None, root=self.root)
                return

        # Priority 2: scope checkbox
        if self.sound_scope_enabled.get(scope_norm, False):
            play_sound_file(self.sound_scope_file.get(scope_norm) or None, root=self.root)


    def _is_keyword_hit(self, speaker: str, content: str) -> bool:
        """キーワードルール（有効なもの）にヒットしたら True"""
        try:
            hay = f"{speaker}\n{content}"
            hay_cmp = self.normalize_for_keyword_match(hay)
            for rule in self.keyword_rules:
                if not rule.get("enabled", True):
                    continue
                kw = (rule.get("keyword") or "").strip()
                if not kw:
                    continue
                kw_cmp = self.normalize_for_keyword_match(kw)
                if kw_cmp and kw_cmp in hay_cmp:
                    return True
            return False
        except Exception:
            return False

    def _is_scope_notify(self, scope: str) -> bool:
        """スコープ通知チェックがONなら True"""
        try:
            scope_norm = self.normalize_scope(scope)
            return bool(self.sound_scope_enabled.get(scope_norm, False))
        except Exception:
            return False


    # ---------- Log cards ----------
    # ---------- Log cards ----------
    def get_enabled_keywords(self):
        return [ (r.get("keyword") or "").strip() for r in self.keyword_rules if r.get("enabled", True) and (r.get("keyword") or "").strip() ]

    def _content_has_keyword(self, content: str) -> bool:
        if not content:
            return False
        try:
            hay_cmp = self.normalize_for_keyword_match(content)
            for kw in self.get_enabled_keywords():
                if not kw:
                    continue
                kw_cmp = self.normalize_for_keyword_match(kw)
                if kw_cmp and (kw_cmp in hay_cmp):
                    return True
            return False
        except Exception:
            # fallback: original (case-sensitive) behavior
            for kw in self.get_enabled_keywords():
                if kw and (kw in content):
                    return True
            return False

    def _apply_keyword_highlight(self, text_widget: tk.Text, content: str):
        """
        Highlight enabled keywords in content with red-bold.
        NOTE: We do NOT apply a competing "base" tag to avoid tag-priority issues on some Tk builds.
        """
        kws = self.get_enabled_keywords()
        if not kws:
            return

        try:
            bold_font = ("Yu Gothic UI", int(getattr(self, "log_font_size", 10) or 10), "bold")
            text_widget.tag_configure("kw", foreground="#DC2626", font=bold_font)  # red + bold
        except Exception:
            pass

        for kw in kws:
            if not kw:
                continue
            try:
                pat = self.build_keyword_regex(kw)
                if not pat:
                    continue
                for m_ in re.finditer(pat, content, flags=re.IGNORECASE):
                    s_idx = f"1.0+{m_.start()}c"
                    e_idx = f"1.0+{m_.end()}c"
                    text_widget.tag_add("kw", s_idx, e_idx)
            except Exception:
                # fallback: simple substring (case-insensitive)
                try:
                    kw_l = (kw or "")
                    if not kw_l:
                        continue
                    start = "1.0"
                    while True:
                        pos = text_widget.search(kw_l, start, stopindex="end-1c", nocase=1)
                        if not pos:
                            break
                        end = f"{pos}+{len(kw_l)}c"
                        text_widget.tag_add("kw", pos, end)
                        start = end
                except Exception:
                    pass
        try:
            text_widget.tag_raise("kw")
        except Exception:
            pass

    def _recalc_text_height(self, text_widget: tk.Text, max_lines: int = 18):
        """
        Recalculate visible display lines and set the height accordingly.
        """
        try:
            # displaylines counts wrapped lines in current width
            cnt = text_widget.count("1.0", "end-1c", "displaylines")
            lines = int(cnt[0]) if cnt else 1
            lines = max(1, min(max_lines, lines))
            text_widget.configure(height=lines)
        except Exception:
            pass


    def _make_text_readonly(self, text_widget: tk.Text):
        """
        Make a Text widget effectively read-only while keeping tag colors (do NOT use state=disabled).
        Allows copy/select and scrolling, blocks editing.
        """
        try:
            text_widget.configure(insertwidth=0, cursor="arrow", takefocus=0)
        except Exception:
            pass

        def on_key(event):
            try:
                # Allow Ctrl+C / Ctrl+A
                if (event.state & 0x4) and event.keysym.lower() in ("c", "a"):
                    return
                # Allow navigation keys
                if event.keysym in ("Left", "Right", "Up", "Down", "Home", "End", "Prior", "Next"):
                    return
            except Exception:
                pass
            return "break"
        def block(_e):
            return "break"

        # Block edits
        text_widget.bind("<KeyPress>", on_key)
        text_widget.bind("<<Paste>>", block)
        text_widget.bind("<<Cut>>", block)
        text_widget.bind("<<Clear>>", block)
        text_widget.bind("<Control-v>", block)
        text_widget.bind("<Control-x>", block)



    def _update_wrap_lengths_for(self, scroller: "ScrolledFrame", items: list):
        try:
            if scroller is None:
                return
            width = scroller.canvas.winfo_width()
            wrap = max(320, width - 90)
            for item in items:
                try:
                    lbl = item.get("content")
                    if lbl is not None:
                        lbl.configure(wraplength=wrap)
                    txt = item.get("text")
                    if txt is not None:
                        # re-fit height because wrapping changes with width
                        self._recalc_text_height(txt)
                        self._make_text_readonly(txt)
                except Exception:
                    pass
        except Exception:
            pass

    def _update_wrap_lengths(self):
        # update all tabs
        try:
            self._update_wrap_lengths_for(self.log_scroller, self.log_items)
            if getattr(self, "log_scroller_notify", None) is not None:
                self._update_wrap_lengths_for(self.log_scroller_notify, self.log_items_notify)
            if getattr(self, "log_scroller_keyword", None) is not None:
                self._update_wrap_lengths_for(self.log_scroller_keyword, self.log_items_keyword)
        except Exception:
            pass

    def _append_card(self, scroller, items: list, scope: str, speaker: str, content: str, ts_str: str, is_system: bool = False, is_hit: bool = False, ts_epoch: float = None):
        """Append a chat log card to the given scroller.
        Uses a Text widget so we can support adjustable font size / line spacing and keyword highlighting.
        """
        scope_norm = self.normalize_scope(scope)
        style = self.SCOPE_STYLE.get(scope_norm, self.SCOPE_STYLE.get("[チャネル]")) or {"bg":"#F3F4F6","border":"#64748B"}
        bg = style["bg"]
        border = style["border"]

        if scroller is None:
            return
        parent = scroller.inner

        # Compactness
        fs = int(getattr(self, "log_font_size", 10) or 10)
        if fs < 8: fs = 8
        if fs > 16: fs = 16
        header_fs = max(8, fs - 1)
        ls = int(getattr(self, "log_line_spacing", 0) or 0)
        if ls < 0: ls = 0
        if ls > 10: ls = 10
        gap = int(getattr(self, "log_card_gap", 2))
        if gap < 0: gap = 0
        if gap > 20: gap = 20
        display_format = str(getattr(self, "log_display_format", "full") or "full").strip().lower()
        if display_format not in ("full", "simple"):
            display_format = "full"


        card = ttk.Frame(parent, style="Card.TFrame")
        card.pack(fill="x", padx=10, pady=gap)
        card.configure(borderwidth=2)

        try:
            card.configure(style="Card.TFrame")
        except Exception:
            pass

        # Border simulation
        border_fr = tk.Frame(card, bg=border)
        border_fr.place(relx=0, rely=0, relwidth=1, relheight=1)
        inner = tk.Frame(card, bg=bg)
        inner.pack(fill="both", expand=True)

        # Header / content layout depends on display format
        lbl_time = None
        lbl_scope = None
        lbl_speaker = None

        if display_format == "full":
            header = tk.Frame(inner, bg=bg)
            header.pack(fill="x", padx=10, pady=(4, 0))

            lbl_time = tk.Label(header, text=self._fmt_time_label(ts_str), bg=bg, fg="#475569", font=("Yu Gothic UI", header_fs))
            lbl_time.pack(side="left")
            lbl_scope = tk.Label(header, text=self._fmt_scope_label(scope_norm), bg=bg, fg="#0F172A", font=("Yu Gothic UI", header_fs, "bold"))
            lbl_scope.pack(side="left", padx=(8, 0))

            speaker_fg = "#0F172A" if not is_system else "#334155"
            lbl_speaker = tk.Label(header, text=f"{speaker}:", bg=bg, fg=speaker_fg, font=("Yu Gothic UI", header_fs, "bold"))
            lbl_speaker.pack(side="left", padx=(8, 0))

            # Content (separate line)
            content_parent = inner
            content_pack = {"fill": "x", "padx": 10, "pady": (2, 6)}
        else:
            # Simple: speaker + content only (time/scope omitted).
            row = tk.Frame(inner, bg=bg)
            row.pack(fill="x", padx=10, pady=(4, 6))
            speaker_fg = "#0F172A" if not is_system else "#334155"
            lbl_speaker = tk.Label(row, text=f"{speaker}:", bg=bg, fg=speaker_fg, font=("Yu Gothic UI", header_fs, "bold"))
            lbl_speaker.pack(side="left")
            content_parent = row
            content_pack = {"side": "left", "fill": "x", "expand": True, "padx": (0, 0), "pady": 0}

        # Content Text widget
        content_text = tk.Text(content_parent, bg=bg, fg="#0F172A", bd=0, highlightthickness=0, wrap="word", padx=0, pady=0)
        content_text.pack(**content_pack)
        content_text.insert("1.0", content)

        # Base tag controls font + spacing
        content_text.tag_configure("base", font=("Yu Gothic UI", fs), spacing1=0, spacing2=0, spacing3=ls)
        content_text.tag_add("base", "1.0", "end")

        if self._content_has_keyword(content):
            self._apply_keyword_highlight(content_text, content)

        content_text.update_idletasks()
        self._recalc_text_height(content_text)
        self._make_text_readonly(content_text)

        header_labels = []
        if lbl_time is not None:
            header_labels.append((lbl_time, False))
        if lbl_scope is not None:
            header_labels.append((lbl_scope, True))
        if lbl_speaker is not None:
            header_labels.append((lbl_speaker, True))

        ts_val = (ts_epoch if ts_epoch is not None else time.time())

        items.append({
            "frame": card,
            "header_labels": header_labels,
            "lbl_time": lbl_time,
            "lbl_scope": lbl_scope,
            "lbl_speaker": lbl_speaker,
            "text": content_text,
            "scope": scope_norm,
            "speaker": speaker,
            "content_raw": content,
            "ts_str": ts_str,
            "is_system": bool(is_system),
            "ts": float(ts_val),
        })

        # Trim
        while len(items) > self.max_log_items:
            old = items.pop(0)
            try:
                old["frame"].destroy()
            except Exception:
                pass

        scroller.scroll_to_bottom()

    def refresh_log_styles(self):
        """Re-apply log font size / spacing to existing cards."""
        fs = int(getattr(self, "log_font_size", 10) or 10)
        if fs < 8: fs = 8
        if fs > 16: fs = 16
        header_fs = max(8, fs - 1)
        ls = int(getattr(self, "log_line_spacing", 0) or 0)
        if ls < 0: ls = 0
        if ls > 10: ls = 10
        gap = int(getattr(self, "log_card_gap", 2))
        if gap < 0: gap = 0
        if gap > 20: gap = 20


        def _apply(items: list):
            for it in items:
                try:
                    for lbl, is_bold in it.get("header_labels", []):
                        if is_bold:
                            lbl.configure(font=("Yu Gothic UI", header_fs, "bold"))
                        else:
                            lbl.configure(font=("Yu Gothic UI", header_fs))
                    # Card gap + label formatting refresh
                    fr = it.get("frame")
                    if fr is not None:
                        try:
                            fr.pack_configure(pady=gap)
                        except Exception:
                            pass
                    lt = it.get("lbl_time")
                    if lt is not None:
                        try:
                            lt.configure(text=self._fmt_time_label(it.get("ts_str", "")))
                        except Exception:
                            pass
                    lsco = it.get("lbl_scope")
                    if lsco is not None:
                        try:
                            lsco.configure(text=self._fmt_scope_label(it.get("scope", "")))
                        except Exception:
                            pass

                    tw = it.get("text")
                    if isinstance(tw, tk.Text):
                        tw.configure(font=("Yu Gothic UI", fs))
                        tw.tag_configure("base", font=("Yu Gothic UI", fs), spacing1=0, spacing2=0, spacing3=ls)
                        # Keep highlight font size in sync
                        try:
                            tw.tag_configure("kw", font=("Yu Gothic UI", fs, "bold"))
                        except Exception:
                            pass
                        try:
                            tw.tag_add("base", "1.0", "end")
                        except Exception:
                            pass
                        tw.update_idletasks()
                        self._recalc_text_height(tw)
                except Exception:
                    continue

        try:
            _apply(getattr(self, "log_items", []))
            _apply(getattr(self, "log_items_notify", []))
            _apply(getattr(self, "log_items_keyword", []))
        except Exception:
            pass
    def rebuild_log_cards(self):
        """Rebuild existing cards (useful when switching display format)."""
        def _rebuild(scroller, items: list):
            if scroller is None or items is None:
                return
            # Snapshot
            entries = []
            for it in list(items):
                entries.append({
                    "scope": it.get("scope", "[チャネル]"),
                    "speaker": it.get("speaker", ""),
                    "content": it.get("content_raw", ""),
                    "ts_str": it.get("ts_str", ""),
                    "is_system": bool(it.get("is_system", False)),
                    "ts": it.get("ts", None),
                })
            # Clear UI
            for it in list(items):
                try:
                    fr = it.get("frame")
                    if fr is not None:
                        fr.destroy()
                except Exception:
                    pass
            items.clear()
            # Re-create
            for e in entries:
                try:
                    self._append_card(scroller, items, e["scope"], e["speaker"], e["content"], e["ts_str"], is_system=e["is_system"], ts_epoch=e.get("ts"))
                except Exception:
                    pass
            try:
                scroller.scroll_to_bottom()
            except Exception:
                pass

        _rebuild(getattr(self, "log_scroller", None), getattr(self, "log_items", None))
        _rebuild(getattr(self, "log_scroller_notify", None), getattr(self, "log_items_notify", None))
        _rebuild(getattr(self, "log_scroller_keyword", None), getattr(self, "log_items_keyword", None))

    def add_system_log(self, text: str):
        t = time.strftime("%H:%M:%S")
        ts_epoch = time.time()
        self.message_buffer.append((t, "[チャネル]", "System", text, ts_epoch, False, False))
        self._append_card(self.log_scroller, self.log_items, "[チャネル]", "System", text, t, is_system=True, ts_epoch=ts_epoch)

    # ---------- Thread-safe UI updates ----------
    def process_ui_queue(self):
        try:
            while True:
                kind, payload = self.ui_queue.get_nowait()
                if kind == "preview":
                    self.safe_update_ui(payload)
                elif kind == "log":
                    self.safe_add_log(payload)
        except queue.Empty:
            pass
        finally:
            try:
                self.root.after(30, self.process_ui_queue)
            except Exception:
                pass


    def enqueue_ui(self, kind: str, payload):
        """Bounded UI queue: drop oldest when full to avoid memory growth."""
        try:
            self.ui_queue.put_nowait((kind, payload))
        except queue.Full:
            try:
                _ = self.ui_queue.get_nowait()
            except Exception:
                pass
            try:
                self.ui_queue.put_nowait((kind, payload))
            except Exception:
                pass

    def _periodic_gc(self):
        """Occasional full GC to keep long-running sessions tidy."""
        try:
            gc.collect()
        except Exception:
            pass
        try:
            self.root.after(600_000, self._periodic_gc)  # every 10 minutes
        except Exception:
            pass
    
    def _world_retention_tick(self):
        """定期的に『全ログ』から古いワールド発言を削除します。

        通知タブ・キーワードタブは保持します。
        """
        try:
            if not getattr(self, "world_retention_enabled", True):
                return
            minutes = float(getattr(self, "world_retention_minutes", 5) or 0)
            if minutes <= 0:
                return
            cutoff = time.time() - (minutes * 60.0)

            # Purge from UI (all logs only)
            self._purge_old_world_cards(cutoff_epoch=cutoff)

            # Purge from copy buffer (keep items that are notify/keyword hits)
            try:
                new_buf = deque(maxlen=self.max_buffer_items)
                for it in list(self.message_buffer):
                    try:
                        t, scope, speaker, content, ts, is_notify, hit_kw = it
                    except Exception:
                        # fallback (old format)
                        try:
                            t, scope, speaker, content = it[0], it[1], it[2], it[3]
                        except Exception:
                            continue
                        ts = time.time()
                        is_notify = False
                        hit_kw = False

                    if (scope == "[ワールド]") and (ts < cutoff) and (not is_notify) and (not hit_kw):
                        continue
                    new_buf.append((t, scope, speaker, content, ts, bool(is_notify), bool(hit_kw)))
                self.message_buffer = new_buf
            except Exception:
                pass
        except Exception:
            pass
        finally:
            try:
                self.root.after(15_000, self._world_retention_tick)
            except Exception:
                pass

    def _purge_old_world_cards(self, cutoff_epoch: float):
        """全ログタブのワールド発言カードのみ、cutoffより古いものを破棄する。"""
        try:
            if not hasattr(self, "log_items") or not getattr(self, "log_items", None):
                return
            items = self.log_items
            i = 0
            changed = False
            while i < len(items):
                it = items[i]
                try:
                    scope = it.get("scope", "")
                    ts = float(it.get("ts", 0.0) or 0.0)
                    is_system = bool(it.get("is_system", False))
                except Exception:
                    scope, ts, is_system = "", 0.0, False

                if (not is_system) and (scope == "[ワールド]") and ts and (ts < cutoff_epoch):
                    try:
                        it.get("frame").destroy()
                    except Exception:
                        pass
                    try:
                        items.pop(i)
                    except Exception:
                        i += 1
                    changed = True
                    continue
                i += 1

            if changed and getattr(self, "log_scroller", None) is not None:
                try:
                    self.log_scroller.refresh_scrollregion()
                except Exception:
                    pass
        except Exception:
            pass

    def safe_update_ui(self, pil_img: Image.Image):
        try:
            img_tk = ImageTk.PhotoImage(pil_img)
            self.preview_label.config(image=img_tk)
            self.preview_label.image = img_tk
        except Exception:
            pass




    def _select_log_tab(self, kind: str, initial: bool = False):
        """Select log tab (all/notify/keyword). Clears blink when opened."""
        try:
            if not hasattr(self, "_log_tab_frames"):
                return
            kind = kind if kind in self._log_tab_frames else "all"
            self.log_tab_cur = kind

            # Raise frame
            try:
                self._log_tab_frames[kind].tkraise()
            except Exception:
                pass

            # Clear alert for the opened tab
            if not initial:
                if kind in ("notify", "keyword"):
                    if self._tab_alert.get(kind):
                        self._tab_alert[kind] = False

            self._refresh_tab_titles()
        except Exception:
            pass

    def _refresh_tab_titles(self):
        """Update tab colors (blink by background color)."""
        try:
            if not hasattr(self, "_log_tab_buttons"):
                return

            cur = getattr(self, "log_tab_cur", "all")
            blink_on = bool(getattr(self, "_tab_blink_on", False))

            for kind, btn in self._log_tab_buttons.items():
                if kind == cur:
                    bg = self._tab_colors["selected"]
                    border = self._tab_colors["border_sel"]
                else:
                    # normal tab
                    bg = self._tab_colors["normal"]
                    border = self._tab_colors["border"]

                    # blink if alerted
                    if kind in ("notify", "keyword") and self._tab_alert.get(kind):
                        if kind == "notify":
                            bg = self._tab_colors["notify_b"] if blink_on else self._tab_colors["notify_a"]
                        elif kind == "keyword":
                            bg = self._tab_colors["kw_b"] if blink_on else self._tab_colors["kw_a"]

                try:
                    btn.config(bg=bg, fg=self._tab_colors["text"])
                    btn.config(highlightbackground=border)
                except Exception:
                    pass

            # stop blink job if no alerts
            if not any(self._tab_alert.values()):
                if self._tab_blink_job is not None:
                    try:
                        self.root.after_cancel(self._tab_blink_job)
                    except Exception:
                        pass
                self._tab_blink_job = None
        except Exception:
            pass

    def _ensure_tab_blink(self):
        try:
            if self._tab_blink_job is None:
                self._tab_blink_on = False
                self._tab_blink_job = self.root.after(420, self._tab_blink_tick)
        except Exception:
            pass

    def _tab_blink_tick(self):
        try:
            self._tab_blink_on = not self._tab_blink_on
            self._refresh_tab_titles()
            if any(self._tab_alert.values()):
                self._tab_blink_job = self.root.after(420, self._tab_blink_tick)
            else:
                self._tab_blink_job = None
        except Exception:
            self._tab_blink_job = None

    def _mark_tab_alert(self, kind: str):
        """kind: 'notify' or 'keyword'"""
        try:
            if kind not in ("notify", "keyword"):
                return

            cur = getattr(self, "log_tab_cur", "all")
            # If user is already on that tab, do not highlight
            if cur == kind:
                return

            self._tab_alert[kind] = True
            self._ensure_tab_blink()
            self._refresh_tab_titles()
        except Exception:
            pass


    def safe_add_log(self, data):
        try:
            speaker, scope, content, *is_sys = data
            is_system = is_sys[0] if is_sys else False
            t = time.strftime("%H:%M:%S")
            scope_norm = self.normalize_scope(scope)
            # play sound for chat messages only (not system)
            notify_scope = False
            hit_kw = False
            if not is_system:
                # notify sound
                self.maybe_play_notification(scope_norm, speaker, content)
                # routing targets
                notify_scope = self._is_scope_notify(scope_norm)
                hit_kw = self._is_keyword_hit(speaker, content)

            ts_epoch = time.time()
            self.message_buffer.append((t, scope_norm, speaker, content, ts_epoch, bool(notify_scope), bool(hit_kw)))
            self._append_card(self.log_scroller, self.log_items, scope_norm, speaker, content, t, is_system=is_system, ts_epoch=ts_epoch)

            # notify tab (scope)
            if (not is_system) and notify_scope and (getattr(self, "log_scroller_notify", None) is not None):
                self._append_card(self.log_scroller_notify, self.log_items_notify, scope_norm, speaker, content, t, is_system=is_system, ts_epoch=ts_epoch)
                self._mark_tab_alert('notify')

            # keyword-hit tab
            if (not is_system) and hit_kw and (getattr(self, "log_scroller_keyword", None) is not None):
                self._append_card(self.log_scroller_keyword, self.log_items_keyword, scope_norm, speaker, content, t, is_system=is_system, ts_epoch=ts_epoch)
                self._mark_tab_alert('keyword')
        except Exception:
            pass

    # ---------- Monitoring ----------
    def start_monitoring(self):
        if not self.area:
            messagebox.showwarning("警告", "先に「設定」から範囲を指定してください。")
            return
        if TESSERACT_PATH == "NOT_FOUND":
            messagebox.showwarning("警告", "Tesseractが見つかりません。インストールを確認してください。")
            return
        if self.monitoring:
            return

        self.monitoring = True
        self.stop_event.clear()
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self._set_status("監視を開始しました。")
        threading.Thread(target=self.monitor_loop, daemon=True).start()

    def stop_monitoring(self):
        if not self.monitoring:
            try:
                self.btn_start.config(state="normal")
                self.btn_stop.config(state="disabled")
            except Exception:
                pass
            return

        self.monitoring = False
        self.stop_event.set()
        try:
            self.btn_start.config(state="normal")
            self.btn_stop.config(state="disabled")
        except Exception:
            pass
        self._set_status("停止しました。")

    def _on_destroy(self, _evt=None):
        try:
            if getattr(self.app, 'settings_win', None) is self:
                self.app.settings_win = None
        except Exception:
            pass

    def on_close(self):
        try:
            self.stop_monitoring()
        except Exception:
            pass
        try:
            self.save_settings()
        except Exception:
            pass
        self.root.destroy()

    # ---------- OCR / parsing ----------
    def clean_speaker_line(self, line: str) -> str:
        """
        発言者行（1行目）の先頭アイコンは画像処理でマスクするため、
        ここでは文字列の先頭から「記号っぽい連続」を落とし、
        さらに「名前にスペースは入らない」前提で、誤認識の空白以降を除去する。
        """
        s = (line or "").strip()
        s = re.sub(r"^[^a-zA-Z0-9\u3040-\u30ff\u4e00-\u9faf]+", "", s).strip()
        # OCRが名前の後ろにスペースを入れることがあるので、最初のトークンだけ残す
        s = re.split(r"\s+", s, maxsplit=1)[0].strip()
        return s

    def make_preview(self, frame_bgr, binarized_gray, rects=None, max_w=800, max_h=190):
        """Create preview image (binarized only) with detected bubble rectangles."""
        try:
            # show ONLY binarized image (compact width)
            bin_rgb = cv2.cvtColor(binarized_gray, cv2.COLOR_GRAY2BGR)
            combined = bin_rgb.copy()

            # draw detected message rectangles (green)
            if rects:
                try:
                    for (x, y, rw, rh) in rects:
                        x0, y0 = int(x), int(y)
                        x1, y1 = int(x + rw), int(y + rh)
                        cv2.rectangle(combined, (x0, y0), (x1, y1), (0, 255, 0), 2)
                except Exception:
                    pass

            h, w = combined.shape[:2]
            max_w, max_h = int(max_w), int(max_h)
            scale = min(max_w / max(1, w), max_h / max(1, h))
            scale = max(scale, 0.01)
            resized = cv2.resize(combined, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_AREA)
            img_rgb = cv2.cvtColor(resized, cv2.COLOR_BGR2RGB)
            return Image.fromarray(img_rgb)
        except Exception:
            return None





    def allowed_scopes(self):
        # the 4 scopes we currently support (color mapping / UI)
        return ("[ワールド]", "[ギルド]", "[パーティ]", "[チャネル]")

    def is_plausible_name(self, name: str) -> bool:
        if not name:
            return False
        name = name.strip()
        if len(name) < 1 or len(name) > 24:
            return False

        # Must contain at least one Japanese kana/kanji OR alphabet
        if not re.search(r"[A-Za-z\u3040-\u30ff\u4e00-\u9faf]", name):
            return False

        # Avoid names that are almost all digits
        digits = sum(ch.isdigit() for ch in name)
        if len(name) >= 6 and digits / max(1, len(name)) > 0.6:
            return False

        # Reject obvious garbage: very long, mostly uppercase/digits, no spaces
        if len(name) >= 12 and re.fullmatch(r"[A-Za-z0-9\-_]+", name):
            has_lower = any("a" <= ch <= "z" for ch in name)
            digits_ratio = digits / max(1, len(name))
            if (not has_lower) and digits_ratio >= 0.2:
                return False

        return True

    def normalize_content_text(self, content: str) -> str:
        """Normalize chat content for display/dedupe stability.
        - Convert fullwidth spaces
        - If 4+ consecutive spaces appear, truncate from there (likely OCR garbage tail)
        - Collapse whitespace
        """
        if not content:
            return ""
        s = content.replace("\u3000", " ")

        # If 4+ consecutive spaces appear, it can be OCR garbage *or* just alignment spacing.
        # Heuristic:
        # - If the suffix contains Japanese characters (or looks meaningful), keep it and just collapse spaces.
        # - Otherwise, truncate from that point.
        m = re.search(r"[ \t]{4,}", s)
        if m:
            suffix = s[m.end():].strip()
            if suffix and re.search(r"[\u3040-\u30ff\u4e00-\u9faf]", suffix):
                # keep but collapse later
                pass
            else:
                s = s[:m.start()].rstrip()

        s = re.sub(r"\s+", " ", s).strip()
        # OCR誤読で本文先頭に "1]" が残る場合があるため除去
        if s.startswith("1]"):
            s = s[2:].lstrip()
        return s

    def is_plausible_content(self, content: str) -> bool:
        if not content:
            return False
        content = content.strip()
        if len(content) < 1:
            return False

        # Reject obvious garbage: long alnum-only
        if len(content) >= 10 and re.fullmatch(r"[A-Za-z0-9\-_==\+\*/\\\.,:;]+", content):
            return False

        # If it contains at least one Japanese char, accept
        if re.search(r"[\u3040-\u30ff\u4e00-\u9faf]", content):
            return True

        # Otherwise allow short content (like "OK", "gg") but reject long
        return len(content) <= 6

    def parse_text_multi(self, raw: str):
        """
        Parse OCR text into messages.
        Robust to bracket misreads such as [ ] -> | 1 I and to spaces inside scope like "ワール ド".
        """
        lines = [l.strip() for l in (raw or "").splitlines() if l.strip()]
        msgs = []
        if len(lines) < 2:
            return msgs

        # --- scope detection (tolerant) ---
        bracket_chars = r"\[\]【】（）\(\)〔〕｛｝\{\}\|｜Iil1"
        dash_chars = r"ー—‐―\-一"
        open_like = r"\[【\(\{〔｛\|｜Iil1"
        close_like = r"\]】\)\}〕｝\|｜Iil1"

        def detect_scope_and_content(line: str):
            if not line:
                return None, None
            # Quick compact for detection
            compact = re.sub(r"\s+", "", line)
            # normalize dash-like & common bracket look-alikes (only for detection)
            compact = compact.translate(str.maketrans({
                "—": "ー", "‐": "ー", "―": "ー", "-": "ー", "一": "ー",
                "［": "[", "］": "]", "【": "[", "】": "]",
                "｜": "|",
            }))
            # pick scope (prefer prefix match to avoid false positives)
            # NOTE: bracket may be misread as |/1/I etc, and spaces may appear inside the token.
            scope = None
            token_pat = None

            # pick scope from the prefix (avoid false positives from the message body)
            # World: safe to accept without brackets (body rarely starts with it)
            m_w = re.match(rf"^[{bracket_chars}\s]*(?:ワ\s*[{dash_chars}]?\s*ル\s*ド|WORLD)", line, flags=re.IGNORECASE)
            if m_w or "WORLD" in compact.upper():
                scope = "[ワールド]"
                token_pat = rf"(?:ワ\s*[{dash_chars}]?\s*ル\s*ド|WORLD)"
            else:
                def _marker_ok(_line, _m_end):
                    l2 = _line.lstrip()
                    if not l2:
                        return False
                    # Require an opening marker at the start of the line (or its common OCR look-alikes),
                    # and also require a closing marker right after the token.
                    if not re.match(r"[\[【\(\{〔｛\|｜Iil1]", l2):
                        return False
                    rest = _line[_m_end:].lstrip()
                    return bool(rest) and re.match(r"[\]】\)\}〕｝\|｜Iil1]", rest)


                mg = re.match(rf"^[{bracket_chars}\s]*(?:ギ\s*ル\s*ド|キ\s*ル\s*ド|GUILD)", line, flags=re.IGNORECASE)
                if mg and _marker_ok(line, mg.end()):
                    scope = "[ギルド]"
                    token_pat = r"(?:ギ\s*ル\s*ド|キ\s*ル\s*ド|GUILD)"
                else:
                    mp = re.match(rf"^[{bracket_chars}\s]*(?:パ\s*[{dash_chars}]?\s*テ\s*(?:ィ|イ)|PARTY)", line, flags=re.IGNORECASE)
                    if mp and _marker_ok(line, mp.end()):
                        scope = "[パーティ]"
                        token_pat = rf"(?:パ\s*[{dash_chars}]?\s*テ\s*(?:ィ|イ)|PARTY)"
                    else:
                        mc = re.match(rf"^[{bracket_chars}\s]*(?:チャ\s*ネ\s*ル|チャン\s*ネ\s*ル|チヤ\s*ネ\s*ル|CHANNEL)", line, flags=re.IGNORECASE)
                        if mc and _marker_ok(line, mc.end()):
                            scope = "[チャネル]"
                            token_pat = r"(?:チャ\s*ネ\s*ル|チャン\s*ネ\s*ル|チヤ\s*ネ\s*ル|CHANNEL)"

            if not scope or not token_pat:
                return None, None

            # Remove prefix like [ワールド] / |ワールド| / 1ワールド1 etc
            post_bracket_chars = bracket_chars.replace("1", "")
            prefix_pat = rf"^[{bracket_chars}\s]*({token_pat})[{post_bracket_chars}\s]*"
            content = re.sub(prefix_pat, "", line).strip()

            # If closing bracket was misread as '1' and glued to leading digits (e.g. "ワールド118"),
            # drop ONLY the first '1' when there is no real closing bracket and there are 3+ digits.
            try:
                has_real_close = bool(re.search(rf"({token_pat})\s*[\]\】]", compact))
            except Exception:
                has_real_close = False

            if (not has_real_close) and re.match(r"^[1１][0-9０-９]{2,}", content):
                content = content[1:].lstrip()

            # Fallback: if not removed well, cut after first token occurrence
            if content == line.strip():
                m = re.search(token_pat, line)
                if m:
                    content = line[m.end():].strip()
                    content = re.sub(rf"^[{bracket_chars}\s]+", "", content).strip()

            return scope, content

        def is_scope_line(line: str) -> bool:
            sc, _ = detect_scope_and_content(line)
            return sc is not None

        # Main scan: find scope line(s)
        for i, line in enumerate(lines):
            scope, content = detect_scope_and_content(line)
            if scope is None:
                continue

            # name line is usually previous line
            if i > 0:
                name = self.clean_speaker_line(lines[i - 1]).strip()
            else:
                # rare: name+scope in same line
                pre = re.split(r"\[|【|\||｜", line, maxsplit=1)[0].strip()
                name = self.clean_speaker_line(pre).strip()

            if (not name) or (not self.is_plausible_name(name)):
                # 先頭が欠けたバブル等で本文が名前扱いになるのを防ぐ
                continue

            scope_norm = self.normalize_scope(scope)
            if scope_norm not in self.allowed_scopes():
                continue

            # Append following wrapped lines until next scope block (rare)
            j = i + 1
            while j < len(lines):
                nxt = lines[j]

                # stop if next is a scope line (or looks like speaker line followed by scope line)
                if is_scope_line(nxt):
                    break
                if j + 1 < len(lines) and is_scope_line(lines[j + 1]):
                    break

                content += (" " if content else "") + nxt
                j += 1

            content = content.strip()
            content = self.normalize_content_text(content)
            if self.is_plausible_content(content):
                msgs.append({"name": name, "scope": scope_norm, "content": content})

        return msgs




    def parse_text_fallback_single(self, raw: str):
        """
        Fallback parser for a single chat bubble.
        Used when line breaks are missing (merged into 1 line) or when parse_text_multi can't parse.
        IMPORTANT: If the bubble does not contain a scope prefix like [ワールド]/[ギルド]/[パーティ]/[チャネル],
        this returns [] (do not log) as requested.
        """
        lines = [l.strip() for l in (raw or "").splitlines() if l.strip()]
        if not lines:
            return []

        # Prefer scope detection anchored at individual lines (scope usually begins line 2)
        name = ""
        scope = ""
        content = ""

        bracket_chars = r"\[\]【】（）\(\)〔〕｛｝\{\}\|｜Iil1"
        dash_chars = r"ー—‐―\-一"
        open_like = r"\[【\(\{〔｛\|｜Iil1"
        close_like = r"\]】\)\}〕｝\|｜Iil1"

        def _detect_scope_line(line: str):
            if not line:
                return None, None
            compact = re.sub(r"\s+", "", line)
            l2 = line.lstrip()
            if not re.match(rf"[{open_like}]", l2):
                return None, None


            if re.search(rf"^[{bracket_chars}\s]*ワ\s*[{dash_chars}]?\s*ル\s*ド", line):
                tok_pat = rf"ワ\s*[{dash_chars}]?\s*ル\s*ド"
                sc = "[ワールド]"
            elif re.search(rf"^[{bracket_chars}\s]*(?:ギ\s*ル\s*ド|キ\s*ル\s*ド)", line):
                tok_pat = r"ギ\s*ル\s*ド|キ\s*ル\s*ド"
                sc = "[ギルド]"
            elif re.search(rf"^[{bracket_chars}\s]*パ\s*[{dash_chars}]?\s*テ\s*(?:ィ|イ)", line):
                tok_pat = rf"パ\s*[{dash_chars}]?\s*テ\s*(?:ィ|イ)"
                sc = "[パーティ]"
            elif re.search(rf"^[{bracket_chars}\s]*(?:チャ\s*ネ\s*ル|チャン\s*ネ\s*ル|チヤ\s*ネ\s*ル)", line) or "CHANNEL" in compact.upper():
                tok_pat = r"チャ\s*ネ\s*ル|チャン\s*ネ\s*ル|チヤ\s*ネ\s*ル|CHANNEL"
                sc = "[チャネル]"
            else:
                return None, None

            # Remove prefix token with surrounding bracket-like chars
            prefix_pat = rf"^[{open_like}\s]*({tok_pat})[{close_like}\s]*"
            ct = re.sub(prefix_pat, "", line).strip()
            return sc, ct

        for i in range(1, len(lines)):
            sc, ct = _detect_scope_line(lines[i])
            if sc:
                cand_name = self.clean_speaker_line(lines[i - 1]).strip()
                if (not cand_name) or (not self.is_plausible_name(cand_name)):
                    return []
                name = cand_name
                scope = sc
                if len(lines) > i + 1:
                    ct = (ct + " " + " ".join(lines[i + 1:])).strip()
                content = ct
                break

        if scope:
            content = self.normalize_content_text(content)
            if not content or not self.is_plausible_content(content):
                return []
            return [{"name": name, "scope": scope, "content": content}]

        # Join to one line for robust search (when line breaks are missing)
        joined = " ".join(lines)

        # Find scope token anywhere, but require bracket-like marker near it OR that it appears very early.
        # This avoids false positives where the message content just mentions 'ワールド' etc.
        token_re = re.compile(
            rf"(?P<prefix>[{bracket_chars}\s]{{0,6}})?"
            rf"(?P<tok>ワ\s*[{dash_chars}]?\s*ル\s*ド|ギ\s*ル\s*ド|キ\s*ル\s*ド|パ\s*[{dash_chars}]?\s*テ\s*(?:ィ|イ)|チャ\s*ネ\s*ル|チャン\s*ネ\s*ル|WORLD|CHANNEL)"
            rf"(?P<suffix>[{bracket_chars}\s]{{0,6}})?",
            re.IGNORECASE
        )

        m = token_re.search(joined)
        if not m:
            return []  # no scope -> do not log

                # Require scope marker context.
        # Important: do NOT treat plain 'ギルド' etc in message body as a scope just because it follows a previous bracket.
        start = m.start()
        open_like = r"\[【\(\{〔｛\|｜Iil1"
        close_like = r"\]】\)\}〕｝\|｜Iil1"
        pre_g = (m.group("prefix") or "")
        suf_g = (m.group("suffix") or "")
        ch_before = joined[start - 1] if start > 0 else ""
        ch_after = joined[m.end()] if m.end() < len(joined) else ""

        prefix_has_open = bool(re.search(rf"[{open_like}]", pre_g)) or bool(re.match(rf"[{open_like}]", ch_before))
        suffix_has_close = bool(re.search(rf"[{close_like}]", suf_g)) or bool(re.match(rf"[{close_like}]", ch_after))

        # Require explicit scope marker context.
        # Do NOT treat plain 'ギルド' etc at line start as a scope unless we see a bracket-like close marker.
        prefix_ok = prefix_has_open and suffix_has_close

        if not prefix_ok:
            return []

        tok = m.group("tok") or ""
        tok_up = tok.upper()
        if "WORLD" in tok_up or re.search(r"ワ\s*.+?ル\s*ド", tok):
            scope = "[ワールド]"
        elif re.search(r"ギ\s*ル\s*ド|キ\s*ル\s*ド", tok):
            scope = "[ギルド]"
        elif re.search(r"パ\s*.+?テ\s*(?:ィ|イ)", tok):
            scope = "[パーティ]"
        else:
            scope = "[チャネル]"

        # Determine name: text before the scope marker (trim brackets if any)
        pre = joined[:m.start()].strip()
        pre = re.sub(rf"[{bracket_chars}\s]+$", "", pre).strip()
        name = self.clean_speaker_line(pre).strip()
        if (not name) or (not self.is_plausible_name(name)):
            return []

        # Content: text after scope marker (also strip bracket-like at start)
        post = joined[m.end():].strip()
        post = re.sub(rf"^[{bracket_chars}\s]+", "", post).strip()
        content = post

        # If still empty, try the remaining lines after the first one
        if not content and len(lines) >= 2:
            # keep detected scope
            name = name or self.clean_speaker_line(lines[0]).strip()
            content = " ".join(lines[1:]).strip()

        if not content or not self.is_plausible_content(content):
            return []
        if not name:
            name = "?"
        return [{"name": name, "scope": scope, "content": content}]

    # ---------- OCR / Dedupe helpers ----------
    def normalize_for_dedupe(self, s: str) -> str:
        if not s:
            return ""
        s = s.replace("\u3000", " ")
        m = re.search(r"[ \t]{4,}", s)
        if m:
            s = s[:m.start()].rstrip()
        s = re.sub(r"\s+", " ", s)
        s = s.strip()
        # OCR誤読で先頭に '1]' が残るケースを正規化
        if s.startswith("1]"):
            s = s[2:].lstrip()
        # common punctuation normalize
        s = s.replace("，", ",").replace("．", ".").replace("：", ":").replace("！", "!").replace("？", "?")
        # 大文字小文字は同一視（O/o, U/u など）
        try:
            s = s.casefold()
        except Exception:
            s = s.lower()
        return s


    def _kata_to_hira_char(self, ch: str) -> str:
        o = ord(ch)
        if 0x30A1 <= o <= 0x30F6:
            return chr(o - 0x60)
        return ch

    def _hira_to_kata_char(self, ch: str) -> str:
        o = ord(ch)
        if 0x3041 <= o <= 0x3096:
            return chr(o + 0x60)
        return ch

    def _is_kana_char(self, ch: str) -> bool:
        o = ord(ch)
        return (0x3040 <= o <= 0x309F) or (0x30A0 <= o <= 0x30FF)

    def _kana_to_hira(self, s: str) -> str:
        out = []
        for ch in s:
            out.append(self._kata_to_hira_char(ch))
        return "".join(out)

    def normalize_for_keyword_match(self, s: str) -> str:
        """Normalize text for keyword matching (space-insensitive, optional kana folding)."""
        if not s:
            return ""
        try:
            s = unicodedata.normalize("NFKC", s)
        except Exception:
            pass
        s = self.normalize_for_dedupe(s)
        if getattr(self, "keyword_ignore_spaces", True):
            s = re.sub(r"[ \t\u3000]+", "", s)
        if getattr(self, "keyword_allow_kana_variants", False):
            s = self._kana_to_hira(s)
        return s

    def build_keyword_regex(self, kw: str) -> str:
        """Build a regex that matches kw even if spaces are inserted between characters.
        If kana-variant option is enabled, hiragana/katakana are treated as equivalent."""
        kw = kw or ""
        try:
            kw = unicodedata.normalize("NFKC", kw)
        except Exception:
            pass
        kw = kw.strip()
        if getattr(self, "keyword_ignore_spaces", True):
            kw = re.sub(r"[ \t\u3000]+", "", kw)
        if not kw:
            return ""
        parts = []
        use_kana = bool(getattr(self, "keyword_allow_kana_variants", False))
        for ch in kw:
            if use_kana and self._is_kana_char(ch):
                hira = self._kata_to_hira_char(ch)
                kata = self._hira_to_kata_char(hira)
                if hira != kata:
                    parts.append(f"[{re.escape(hira)}{re.escape(kata)}]")
                else:
                    parts.append(re.escape(ch))
            else:
                parts.append(re.escape(ch))
        sep = r"[ \t\u3000]*" if getattr(self, "keyword_ignore_spaces", True) else ""
        return sep.join(parts) if sep else "".join(parts)

    def is_duplicate(self, name: str, scope: str, content_norm: str) -> bool:
        try:
            name_cmp = self.normalize_for_dedupe(name)
            scope_cmp = self.normalize_scope(scope)
        except Exception:
            name_cmp = (name or "")
            scope_cmp = self.normalize_scope(scope)

        # Note: content_norm is expected to be already normalized with normalize_for_dedupe()

        # Exact match first
        key = f"{name_cmp}|{scope_cmp}|{content_norm}"
        if key in self.last_keys:
            return True

        # Similarity-based dedupe (optional)
        if not self.use_similarity_dedupe:
            return False

        # Very short strings are too unstable for similarity; require exact
        if len(content_norm) < 6:
            return False

        threshold = float(self.similarity_threshold)
        # Compare only within recent history, same speaker+scope
        for prev in reversed(self.last_msgs):
            if prev.get("name_cmp") != name_cmp or prev.get("scope_cmp") != scope_cmp:
                continue
            prev_text = prev.get("content_norm", "")
            if not prev_text:
                continue
            ratio = difflib.SequenceMatcher(None, content_norm, prev_text).ratio()
            if ratio >= threshold:
                return True
        return False

    def remember_message(self, name: str, scope: str, content_norm: str):
        name_cmp = self.normalize_for_dedupe(name)
        scope_cmp = self.normalize_scope(scope)
        key = f"{name_cmp}|{scope_cmp}|{content_norm}"
        self.last_keys.append(key)
        self.last_msgs.append({"name_cmp": name_cmp, "scope_cmp": scope_cmp, "content_norm": content_norm})
        # cap history
        while len(self.last_keys) > int(self.dedupe_history):
            self.last_keys.pop(0)
        while len(self.last_msgs) > int(self.dedupe_history):
            self.last_msgs.pop(0)




    def preprocess_for_ocr(self, frame_bgr: np.ndarray):
        """
        OCR前処理（軽量版）:
        - 必要に応じて拡大（ocr_scale）
        - 二値化（しきい値 current_threshold）
        Returns: (binarized_for_ocr, binarized_for_preview, scale)
        """
        gray = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2GRAY)

        proc = gray
        try:
            scale = float(getattr(self, "ocr_scale", 1.0))
        except Exception:
            scale = 1.0

        if scale < 1.0:
            scale = 1.0

        if scale > 1.0:
            proc = cv2.resize(proc, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)

        thr = int(self.current_threshold)
        _, bin_ocr = cv2.threshold(proc, thr, 255, cv2.THRESH_BINARY)

        # preview should match original size for make_preview()
        if scale > 1.0:
            bin_prev = cv2.resize(bin_ocr, (gray.shape[1], gray.shape[0]), interpolation=cv2.INTER_AREA)
        else:
            bin_prev = bin_ocr.copy()

        return bin_ocr, bin_prev, scale


    def _expand_rects_for_bubble(self, rects, W, H):
        """
        Expand tight text rectangles slightly so that cropping includes name/scope and avoids cutting edges.
        """
        try:
            if not rects:
                return []
            pad_l = max(8, int(getattr(self, "icon_mask_w", 28) * 0.6))
            pad_r = 10
            pad_t = 6
            pad_b = 6
            out = []
            for (x, y, rw, rh) in rects:
                x0 = max(0, int(x) - pad_l)
                y0 = max(0, int(y) - pad_t)
                x1 = min(W, int(x + rw) + pad_r)
                y1 = min(H, int(y + rh) + pad_b)
                if x1 - x0 >= 50 and y1 - y0 >= 22:
                    out.append((x0, y0, x1 - x0, y1 - y0))
            return out
        except Exception:
            return rects or []

    def _score_rects_by_text_density(self, rects, inv_text, W, H):
        """
        Score candidate rect sets. Higher score = better (more rectangles with real text inside, fewer huge blobs).
        inv_text: text pixels should be >0 (e.g. 255 - bin_prev)
        """
        try:
            if not rects:
                return -1e9
            total = 0
            dens_sum = 0.0
            total_area = 0.0
            huge = False
            for (x, y, rw, rh) in rects:
                x0 = max(0, int(x)); y0 = max(0, int(y))
                x1 = min(W, x0 + int(rw)); y1 = min(H, y0 + int(rh))
                if x1 <= x0 or y1 <= y0:
                    continue
                area = float((x1 - x0) * (y1 - y0))
                total_area += area
                if area > 0.70 * (W * H):
                    huge = True
                patch = inv_text[y0:y1, x0:x1]
                nz = float(cv2.countNonZero(patch))
                dens = nz / area
                dens_sum += dens
                total += 1
            if total <= 0:
                return -1e9
            avg_dens = dens_sum / total
            area_ratio = total_area / float(W * H)
            score = (total * (avg_dens + 0.002)) / (1.0 + 1.7 * area_ratio)
            if huge:
                score *= 0.25
            return score
        except Exception:
            return -1e9

    def detect_chat_bubble_rects_auto(self, frame_bgr: np.ndarray, bin_prev: np.ndarray, vmin_override=None, smax_override=None):
        """
        Robust bubble detection:
        - Color-based (white-ish low saturation) detection may fail when the game background is bright.
        - Text-density detection (from binarized image) is more stable for bright scenes.
        This method computes both and selects the better one by text density score.
        """
        try:
            if frame_bgr is None:
                return []
            H, W = frame_bgr.shape[:2]

            # Text-based candidates (tight -> expand slightly)
            rects_text = []
            if bin_prev is not None:
                try:
                    rects_text = self.detect_message_rects(bin_prev)
                    rects_text = self._expand_rects_for_bubble(rects_text, W, H)
                except Exception:
                    rects_text = []

            # Color-based candidates
            rects_color = self.detect_chat_bubble_rects(frame_bgr, vmin_override=vmin_override, smax_override=smax_override)

            if bin_prev is None:
                return rects_color or rects_text

            inv = 255 - bin_prev  # text pixels should be bright in inv
            sc_text = self._score_rects_by_text_density(rects_text, inv, W, H)
            sc_color = self._score_rects_by_text_density(rects_color, inv, W, H)

            # Filter obviously-empty color rects (background blobs)
            if rects_color:
                try:
                    filtered = []
                    for (x, y, rw, rh) in rects_color:
                        x0 = max(0, int(x)); y0 = max(0, int(y))
                        x1 = min(W, x0 + int(rw)); y1 = min(H, y0 + int(rh))
                        if x1 - x0 < 60 or y1 - y0 < 22:
                            continue
                        area = float((x1 - x0) * (y1 - y0))
                        dens = float(cv2.countNonZero(inv[y0:y1, x0:x1])) / max(1.0, area)
                        if dens >= 0.0012:
                            filtered.append((x0, y0, x1 - x0, y1 - y0))
                    if filtered:
                        sc_color_f = self._score_rects_by_text_density(filtered, inv, W, H)
                        if sc_color_f > sc_color:
                            rects_color, sc_color = filtered, sc_color_f
                except Exception:
                    pass

            # Choose better set
            if sc_text > sc_color * 1.05:
                return rects_text
            if sc_color > sc_text * 1.05:
                return rects_color

            # Close call: prefer the one with more rects (likely per-bubble)
            if len(rects_text) >= len(rects_color):
                return rects_text
            return rects_color
        except Exception:
            return []


    def detect_chat_bubble_rects(self, frame_bgr: np.ndarray, vmin_override=None, smax_override=None):
            """
            Detect chat bubble rectangles by color (bright + low saturation regions).

            IMPORTANT:
              - This should return *per chat bubble* rectangles (one per message), not a single giant ROI.
              - We intentionally use a *wide but short* closing kernel to avoid merging vertically stacked bubbles.

            Returns: list[(x,y,w,h)] in original (capture) coordinates.
            """
            try:
                if frame_bgr is None:
                    return []

                H, W = frame_bgr.shape[:2]
                if H < 30 or W < 80:
                    return []

                hsv = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2HSV)
                vmin = int(vmin_override) if vmin_override is not None else int(self.bubble_v_min)
                smax = int(smax_override) if smax_override is not None else int(self.bubble_s_max)
                vmin = max(80, min(250, vmin))
                smax = max(0, min(160, smax))

                # White-ish region: low saturation, high value (bubble background)
                mask = cv2.inRange(hsv, (0, 0, vmin), (179, smax, 255))

                # Morphology tuned for bubbles:
                # - open: remove small noise
                # - close: fill holes INSIDE bubbles but avoid bridging BETWEEN bubbles (short height kernel)
                open_k = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))

                # wide but short close kernel (avoid vertical merging)
                kw = max(7, min(31, int(W * 0.035)))
                kh = max(3, min(9, int(H * 0.010)))
                if kw % 2 == 0: kw += 1
                if kh % 2 == 0: kh += 1
                close_k = cv2.getStructuringElement(cv2.MORPH_RECT, (kw, kh))

                mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, open_k, iterations=1)
                mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, close_k, iterations=1)

                contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

                rects = []
                min_area = max(120, int(getattr(self, "rect_min_area", 120)))
                max_area = int(0.80 * W * H)  # ignore giant rects

                for c in contours:
                    x, y, rw, rh = cv2.boundingRect(c)
                    area = rw * rh
                    if area < min_area:
                        continue
                    if area > max_area:
                        continue
                    if rw < 80 or rh < 28:
                        continue

                    # additional guard: reject nearly full-width/height boxes (often ROI)
                    if rw > 0.98 * W and rh > 0.70 * H:
                        continue

                    pad = 2
                    x0 = max(0, x - pad)
                    y0 = max(0, y - pad)
                    x1 = min(W, x + rw + pad)
                    y1 = min(H, y + rh + pad)
                    rects.append((x0, y0, x1 - x0, y1 - y0))

                rects.sort(key=lambda r: (r[1], r[0]))

                # Merge ONLY if rectangles heavily overlap (split edges)
                merged = []
                for (x, y, rw, rh) in rects:
                    if not merged:
                        merged.append([x, y, rw, rh])
                        continue
                    mx, my, mw, mh = merged[-1]
                    # IoU-like overlap check
                    ix0 = max(mx, x); iy0 = max(my, y)
                    ix1 = min(mx + mw, x + rw); iy1 = min(my + mh, y + rh)
                    iw = max(0, ix1 - ix0); ih = max(0, iy1 - iy0)
                    inter = iw * ih
                    union = (mw * mh) + (rw * rh) - inter
                    iou = (inter / union) if union > 0 else 0.0
                    if iou > 0.35:
                        nx = min(mx, x)
                        ny = min(my, y)
                        nr = max(mx + mw, x + rw)
                        nb = max(my + mh, y + rh)
                        merged[-1] = [nx, ny, nr - nx, nb - ny]
                    else:
                        merged.append([x, y, rw, rh])

                # Safety: if we ended up with one huge rect, treat as failure
                if len(merged) == 1:
                    x, y, rw, rh = merged[0]
                    if (rw * rh) > (0.65 * W * H):
                        return []

                return [tuple(r) for r in merged]
            except Exception:
                return []

    def detect_message_rects(self, bin_img: np.ndarray):
        """
        Detect per-message rectangles from a binarized image (bg=255, text=0).
        Returns list[(x,y,w,h)] in the SAME coordinate as bin_img.
        """
        try:
            if bin_img is None:
                return []
            h, w = bin_img.shape[:2]
            if h < 30 or w < 80:
                return []

            inv = 255 - bin_img

            kw = max(10, min(200, int(self.rect_kernel_w)))
            kh = max(6, min(80, int(self.rect_kernel_h)))

            # 1) connect characters into lines
            k1 = cv2.getStructuringElement(cv2.MORPH_RECT, (kw, 3))
            dil1 = cv2.dilate(inv, k1, iterations=1)

            # 2) connect name-line + message-line
            k2 = cv2.getStructuringElement(cv2.MORPH_RECT, (3, kh))
            dil2 = cv2.dilate(dil1, k2, iterations=1)

            contours, _ = cv2.findContours(dil2, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            rects = []
            min_area = int(self.rect_min_area)

            for c in contours:
                x, y, rw, rh = cv2.boundingRect(c)
                if rw * rh < min_area:
                    continue
                if rh < 18 or rw < 80:
                    continue
                rects.append((x, y, rw, rh))

            if not rects:
                return []

            rects.sort(key=lambda r: (r[1], r[0]))

            # merge close rects
            merged = []
            gap = max(0, min(60, int(self.rect_merge_gap)))
            for (x, y, rw, rh) in rects:
                if not merged:
                    merged.append([x, y, rw, rh])
                    continue
                mx, my, mw, mh = merged[-1]
                overlap = min(mx + mw, x + rw) - max(mx, x)
                overlap_ok = overlap > 0.35 * min(mw, rw)
                if y <= my + mh + gap and overlap_ok:
                    nx = min(mx, x); ny = min(my, y)
                    nr = max(mx + mw, x + rw); nb = max(my + mh, y + rh)
                    merged[-1] = [nx, ny, nr - nx, nb - ny]
                else:
                    merged.append([x, y, rw, rh])
            return [tuple(r) for r in merged]
        except Exception:
            return []

    def apply_icon_mask_by_rects(self, bin_ocr: np.ndarray, rects: list, scale: float):
        """
        Mask top-left icon area for each detected message rect.
        rects are in original coord (bin_prev). Apply on bin_ocr using 'scale'.
        """
        if not self.icon_mask_enabled:
            return bin_ocr
        try:
            if bin_ocr is None or not rects:
                return bin_ocr
            s = float(scale) if scale else 1.0
            mw_s = max(6, int(self.icon_mask_w * s))
            mh_s = max(6, int(self.icon_mask_h * s))

            H, W = bin_ocr.shape[:2]

            def _estimate_firstline_mask_height(crop_bin, mh_default):
                """Estimate a safe mask height that stays within the first line (icon/name line).

                We look for a horizontal 'gap' (few dark pixels) between the first and second text lines.
                If found, we clamp the mask height to that gap so we don't erase the scope label on line 2.
                """
                try:
                    h = int(crop_bin.shape[0])
                    if h < 16:
                        return mh_default

                    # Count dark pixels (text) per row in the top portion
                    row = (crop_bin < 128).sum(axis=1).astype(int)
                    scan_end = max(12, min(h, int(h * 0.75)))
                    top = row[:scan_end]
                    maxv = int(top.max()) if top.size else 0
                    if maxv <= 0:
                        return mh_default

                    gap_thresh = max(3, int(maxv * 0.12))

                    # Find where text starts
                    start = 0
                    for i in range(min(scan_end, h)):
                        if row[i] > gap_thresh * 2:
                            start = i
                            break

                    gap_y = None
                    # Look for a small-gap region that is followed by another text region (2nd line)
                    for yy in range(min(h - 6, start + 8), scan_end - 3):
                        if row[yy] <= gap_thresh and row[yy + 1] <= gap_thresh:
                            if row[yy + 3:scan_end].max() > gap_thresh * 2:
                                gap_y = yy
                                break

                    if gap_y is None:
                        return mh_default

                    # Clamp: keep at least a little height, but stop before the gap
                    return max(6, min(mh_default, int(gap_y)))
                except Exception:
                    return mh_default

            for (x, y, rw, rh) in rects:
                xs = int(x * s); ys = int(y * s)
                x0 = max(0, min(W - 1, xs))
                y0 = max(0, min(H - 1, ys))

                # Determine a per-bubble safe mask height so we don't wipe the scope label
                x2 = max(0, min(W, int((x + rw) * s)))
                y2 = max(0, min(H, int((y + rh) * s)))
                mh_eff = mh_s
                if x2 > x0 + 12 and y2 > y0 + 16:
                    crop = bin_ocr[y0:y2, x0:x2]
                    mh_eff = _estimate_firstline_mask_height(crop, mh_s)

                x1 = max(0, min(W, x0 + mw_s))
                y1 = max(0, min(H, y0 + mh_eff))
                if x1 > x0 and y1 > y0:
                    bin_ocr[y0:y1, x0:x1] = 255
            return bin_ocr
        except Exception:
            return bin_ocr

    def monitor_loop(self):
        while self.monitoring and not self.stop_event.is_set():
            try:
                shot = pyautogui.screenshot(region=self.area)
                frame = cv2.cvtColor(np.array(shot), cv2.COLOR_RGB2BGR)
                bin_ocr, bin_prev, scale = self.preprocess_for_ocr(frame)

                if self.use_bubble_rects:
                    rects = self.detect_chat_bubble_rects_auto(frame, bin_prev)
                    try:
                        H0, W0 = bin_prev.shape[:2]
                        if (not rects) or (len(rects) == 1 and rects[0][2] * rects[0][3] > 0.85 * (W0 * H0)):
                            rects = self.detect_message_rects(bin_prev)
                    except Exception:
                        pass
                else:
                    rects = self.detect_message_rects(bin_prev)
                self.last_rects = rects
                # Apply icon mask per-bubble (prevents icon -> kanji noise)
                mask_rects = rects
                bin_ocr = self.apply_icon_mask_by_rects(bin_ocr, mask_rects, scale)

                prev = None


                if not bool(getattr(self, "preview_collapsed", False)):


                    prev = self.make_preview(frame, bin_prev, rects, max_w=getattr(self, 'preview_target_w', 800), max_h=190)
                if prev is not None:
                    self.enqueue_ui("preview", prev)

                # --- OCR per chat bubble (each green rectangle) ---
                all_msgs = []
                if rects:
                    for (rx, ry, rw, rh) in rects:
                        try:
                            rx = int(max(0, rx)); ry = int(max(0, ry))
                            rw = int(max(1, rw)); rh = int(max(1, rh))
                            rx2 = int(min(frame.shape[1], rx + rw))
                            ry2 = int(min(frame.shape[0], ry + rh))
                            if rx2 - rx < 40 or ry2 - ry < 20:
                                continue
                            bubble = frame[ry:ry2, rx:rx2]
                            b_bin_ocr, _, b_scale = self.preprocess_for_ocr(bubble)
                            if self.icon_mask_enabled:
                                mw = int(max(0, min(b_bin_ocr.shape[1], int(self.icon_mask_w * b_scale))))
                                mh = int(max(0, min(b_bin_ocr.shape[0], int(self.icon_mask_h * b_scale))))
                                if mw > 0 and mh > 0:
                                    b_bin_ocr[0:mh, 0:mw] = 255
                            b_raw = pytesseract.image_to_string(
                                Image.fromarray(b_bin_ocr),
                                lang="jpn",
                                config="--oem 1 --psm 6"
                            )
                            b_msgs = self.parse_text_multi(b_raw)
                            if not b_msgs:
                                b_msgs = self.parse_text_fallback_single(b_raw)
                            all_msgs.extend(b_msgs)
                        except Exception:
                            continue
                else:
                    raw = pytesseract.image_to_string(
                        Image.fromarray(bin_ocr),
                        lang="jpn",
                        config="--oem 1 --psm 6"
                    )
                    all_msgs = self.parse_text_multi(raw)

                # If requested (after Clear), seed dedupe from the current screen
                # without outputting logs. This prevents the same on-screen messages from
                # flooding the log right after clearing.
                if self.seed_dedupe_on_next and all_msgs:
                    try:
                        for mm in all_msgs:
                            n0 = (mm.get("name") or "").strip()
                            sc0 = self.normalize_scope(mm.get("scope") or "")
                            c0 = (mm.get("content") or "").strip()
                            cn0 = self.normalize_for_dedupe(c0)
                            if n0 and cn0:
                                self.remember_message(n0, sc0, cn0)
                    except Exception:
                        pass
                    self.seed_dedupe_on_next = False
                    # Skip printing this cycle
                    self.stop_event.wait(float(self.current_interval))
                    continue

                for m in all_msgs:
                    name = (m.get("name") or "").strip()
                    scope_norm = self.normalize_scope(m.get("scope") or "")
                    content = (m.get("content") or "").strip()

                    content_norm = self.normalize_for_dedupe(content)
                    if not content_norm:
                        continue

                    if not self.is_duplicate(name, scope_norm, content_norm):
                        self.enqueue_ui("log", (name, scope_norm, content))
                        self.remember_message(name, scope_norm, content_norm)

            except Exception as loop_err:
                self.enqueue_ui("log", ("System", "[チャネル]", f"エラー: {loop_err}", True))

            self.stop_event.wait(float(self.current_interval))



# --- Compatibility safeguard (v78): ensure ChatMonitorApp.open_settings exists ---
# 途中の編集で open_settings がクラス外に出てしまうと、起動時に AttributeError になるため保険を入れる。
try:
    getattr(ChatMonitorApp, "open_settings")
except Exception:
    def _fallback_open_settings(self):
        """Open settings window (fallback)."""
        try:
            win = getattr(self, 'settings_win', None)
            if win is not None and win.winfo_exists():
                try:
                    win.deiconify()
                except Exception:
                    pass
                try:
                    win.lift()
                    win.focus_force()
                except Exception:
                    pass
                return
        except Exception:
            self.settings_win = None
        try:
            self.settings_win = SettingsWindow(self)
            if hasattr(self, 'apply_window_attributes'):
                try:
                    self.apply_window_attributes()
                except Exception:
                    pass
        except Exception:
            try:
                self.settings_win = SettingsWindow(self)
            except Exception:
                pass
    ChatMonitorApp.open_settings = _fallback_open_settings

if __name__ == "__main__":
    root = tk.Tk()
    app = ChatMonitorApp(root)
    root.mainloop()
