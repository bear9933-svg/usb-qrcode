"""
QR / 條碼掃描器 + 任意 Office 游標填入
不需指定檔案：只要 Excel 或 Word 開著，滑鼠點哪格就填哪格。
需要：opencv-python, pyzbar, Pillow, pywin32
選用：openpyxl, xlwings
"""

import cv2
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import numpy as np
import os
import json
import threading
import time
import subprocess
import sys
from datetime import datetime

# ── 套件檢查 ──────────────────────────────────────────────
try:
    from pyzbar import pyzbar
    from pyzbar.pyzbar import ZBarSymbol
    PYZBAR_OK = True
except ImportError:
    PYZBAR_OK = False
    ZBarSymbol = None

try:
    import openpyxl
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

try:
    import xlwings as xw
    XLWINGS_OK = True
except ImportError:
    XLWINGS_OK = False

try:
    from PIL import Image, ImageTk
    PIL_OK = True
except ImportError:
    PIL_OK = False

# ── pywin32：抓任意 Office 視窗游標 ──────────────────────
try:
    import win32com.client
    import win32gui
    import win32process
    import win32con
    import pythoncom
    WIN32_OK = True
except ImportError:
    WIN32_OK = False
    pythoncom = None

# ── 掃描碼種 ──────────────────────────────────────────────
SCAN_SYMBOLS = None
if ZBarSymbol is not None:
    _wanted = ["CODE128"]
    SCAN_SYMBOLS = [getattr(ZBarSymbol, n) for n in _wanted if hasattr(ZBarSymbol, n)]

def _suppress_zbar_stderr():
    devnull = open("nul" if sys.platform=="win32" else "/dev/null", "w")
    old_fd  = os.dup(2)
    os.dup2(devnull.fileno(), 2)
    return old_fd, devnull

def _restore_stderr(old_fd, devnull):
    os.dup2(old_fd, 2)
    os.close(old_fd)
    devnull.close()

# ── Office COM 游標偵測 ───────────────────────────────────
def _com_init():
    """確保目前執行緒的 COM 已初始化（STA）。"""
    if pythoncom:
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass


def _com_release():
    """釋放目前執行緒的 COM 物件並 CoUninitialize。
    必須在「建立 COM 物件的同一執行緒」呼叫才有效（STA 規定）。
    """
    if not WIN32_OK:
        return
    for prog_id in (
        "Excel.Application", "Excel.Application.16",
        "Excel.Application.15", "Excel.Application.14",
        "Word.Application",   "Word.Application.16",
        "Word.Application.15","Word.Application.14",
    ):
        try:
            obj = win32com.client.GetActiveObject(prog_id)
            del obj
        except Exception:
            pass
    if pythoncom:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


# ── 單一常駐 COM 執行緒 ────────────────────────────────────
# 問題根源：每次 _poll_cursor 都開新執行緒呼叫 COM，
# 每條執行緒的 STA apartment 從未釋放，累積後造成 ROT 殘留，
# 重新啟動時 GetActiveObject 拿到 stale 物件而失敗。
# 解法：所有 COM 操作統一在同一條長駐執行緒執行，
# 只做一次 CoInitialize / CoUninitialize，完全避免 apartment 洩漏。
import queue as _queue

_com_thread      = None          # 常駐 COM 執行緒
_com_req_q       = _queue.Queue()  # 請求佇列：放 (fn, result_holder)
_com_stop_event  = threading.Event()


def _com_worker():
    """常駐 COM 執行緒主迴圈：統一做 init/call/uninit，避免 apartment 洩漏。"""
    _com_init()
    while not _com_stop_event.is_set():
        try:
            fn, holder = _com_req_q.get(timeout=0.1)
            try:
                holder["result"] = fn()
            except Exception as e:
                holder["result"] = None
                holder["error"]  = e
            finally:
                holder["done"].set()
        except _queue.Empty:
            continue
    # 執行緒結束前正確釋放本執行緒的 COM
    _com_release()


def _com_thread_start():
    """啟動常駐 COM 執行緒（只啟一次）。"""
    global _com_thread
    _com_stop_event.clear()
    if _com_thread is None or not _com_thread.is_alive():
        _com_thread = threading.Thread(target=_com_worker, daemon=True,
                                       name="COM-worker")
        _com_thread.start()


def _com_thread_stop():
    """停止常駐 COM 執行緒，等待它正確釋放 COM 後才返回。"""
    _com_stop_event.set()
    if _com_thread and _com_thread.is_alive():
        _com_thread.join(timeout=2.0)


def _com_call(fn, timeout=1.5):
    """在常駐 COM 執行緒執行 fn()，回傳結果；逾時或失敗回傳 None。"""
    if not WIN32_OK:
        return None
    holder = {"result": None, "error": None, "done": threading.Event()}
    _com_req_q.put((fn, holder))
    if holder["done"].wait(timeout=timeout):
        return holder["result"]
    return None  # 逾時


def _get_office_cursor():
    """
    回傳目前 Office 選取位置的 dict，或 None。
    透過常駐 COM 執行緒呼叫，確保同一 STA apartment，不會有 stale 連線問題。
    """
    def _fn():
        for prog_id in ("Excel.Application", "Excel.Application.16",
                        "Excel.Application.15", "Excel.Application.14"):
            try:
                xl = win32com.client.GetActiveObject(prog_id)
                wb = xl.ActiveWorkbook
                ws = xl.ActiveSheet
                if wb is None or ws is None:
                    continue
                sel = xl.Selection
                r   = sel.Row
                c   = sel.Column
                ltr = _col_num_to_letter(c)
                return {
                    "app":   "excel",
                    "book":  wb.Name,
                    "sheet": ws.Name,
                    "row":   r,
                    "col":   c,
                    "addr":  f"{ltr}{r}",
                }
            except Exception:
                continue
        for prog_id in ("Word.Application", "Word.Application.16",
                        "Word.Application.15", "Word.Application.14"):
            try:
                wd  = win32com.client.GetActiveObject(prog_id)
                doc = wd.ActiveDocument
                if doc is None:
                    continue
                sel = wd.Selection
                return {
                    "app":  "word",
                    "doc":  doc.Name,
                    "para": sel.Paragraphs(1).Range.Start,
                }
            except Exception:
                continue
        return None
    return _com_call(_fn)


def _diagnose_office_com() -> list[str]:
    """
    回傳診斷訊息清單，幫助判斷為何無法連上 Office COM。
    """
    lines = []
    if not WIN32_OK:
        lines.append("❌ pywin32 未安裝或 import 失敗")
        lines.append("   → 執行：pip install pywin32")
        lines.append("   → 再執行：python Scripts/pywin32_postinstall.py -install")
        return lines

    lines.append("✔ pywin32 import 成功")
    _com_init()

    # 檢查 Excel
    xl_found = False
    for prog_id in ("Excel.Application", "Excel.Application.16",
                    "Excel.Application.15", "Excel.Application.14"):
        try:
            xl = win32com.client.GetActiveObject(prog_id)
            xl_found = True
            wb = xl.ActiveWorkbook
            ws = xl.ActiveSheet
            if wb is None:
                lines.append(f"⚠ Excel ({prog_id}) 找到，但沒有活頁簿（請開啟一個 xlsx）")
            elif ws is None:
                lines.append(f"⚠ Excel ({prog_id}) 找到，但 ActiveSheet 為 None")
            else:
                lines.append(f"✔ Excel ({prog_id})：{wb.Name} / {ws.Name}")
                try:
                    sel = xl.Selection
                    lines.append(f"   游標：Row={sel.Row}, Col={sel.Column}")
                except Exception as e:
                    lines.append(f"   ⚠ Selection 讀取失敗：{e}")
            break
        except Exception as e:
            err = str(e)
            if "Invalid class string" in err or "class not registered" in err.lower():
                lines.append(f"❌ {prog_id} 未安裝於此系統")
            elif "Operation unavailable" in err or "-2147221021" in err:
                lines.append(f"⚠ {prog_id} 未在執行中（ROT 查無此項）")
            # 繼續嘗試下一個版本

    if not xl_found:
        lines.append("❌ 找不到執行中的 Excel")
        lines.append("   可能原因：")
        lines.append("   1. Excel 未開啟，或開啟後沒有活頁簿")
        lines.append("   2. Excel 以系統管理員執行，但本程式用一般權限（或反之）")
        lines.append("   3. pywin32 安裝後未執行 pywin32_postinstall")

    # 檢查 Word
    wd_found = False
    for prog_id in ("Word.Application", "Word.Application.16",
                    "Word.Application.15", "Word.Application.14"):
        try:
            wd = win32com.client.GetActiveObject(prog_id)
            wd_found = True
            doc = wd.ActiveDocument
            if doc is None:
                lines.append(f"⚠ Word ({prog_id}) 找到，但沒有開啟文件")
            else:
                lines.append(f"✔ Word ({prog_id})：{doc.Name}")
            break
        except Exception:
            continue

    if not wd_found:
        lines.append("❌ 找不到執行中的 Word（Word 未開啟或版本不符）")

    return lines

def _col_num_to_letter(n):
    ltr = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        ltr = chr(65 + rem) + ltr
    return ltr

def _inject_to_office(data: str, cursor: dict) -> tuple[bool, str]:
    """
    把 data 寫入 cursor 描述的位置（透過常駐 COM 執行緒，避免 apartment 洩漏）。
    """
    if not WIN32_OK:
        return False, "需要 pywin32"
    def _fn():
        if cursor["app"] == "excel":
            for prog_id in ("Excel.Application", "Excel.Application.16",
                            "Excel.Application.15", "Excel.Application.14"):
                try:
                    xl   = win32com.client.GetActiveObject(prog_id)
                    ws   = xl.ActiveSheet
                    row  = cursor["row"]
                    col  = cursor["col"]
                    cell = ws.Cells(row, col)
                    if data.isdigit() and data.startswith("0"):
                        cell.NumberFormat = "@"
                    cell.Value = data
                    return True, ""
                except Exception:
                    continue
            return False, "無法連上 Excel（請確認 Excel 仍開啟）"
        elif cursor["app"] == "word":
            for prog_id in ("Word.Application", "Word.Application.16",
                            "Word.Application.15", "Word.Application.14"):
                try:
                    wd  = win32com.client.GetActiveObject(prog_id)
                    sel = wd.Selection
                    sel.TypeText(data)
                    sel.TypeParagraph()
                    return True, ""
                except Exception:
                    continue
            return False, "無法連上 Word（請確認 Word 仍開啟）"
        return False, "未知應用程式"
    result = _com_call(_fn, timeout=3.0)
    return result if result else (False, "COM 呼叫逾時或失敗")


def _move_excel_down(row: int, col: int):
    """精確移到 row+1 同欄，不受 Excel Enter 設定影響。
    用獨立執行緒直接呼叫 COM，不進共用佇列排隊，確保延遲時間準確。
    """
    def _fn():
        if not WIN32_OK:
            return
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            for prog_id in ("Excel.Application", "Excel.Application.16",
                            "Excel.Application.15", "Excel.Application.14"):
                try:
                    xl = win32com.client.GetActiveObject(prog_id)
                    xl.ActiveSheet.Cells(row + 1, col).Select()
                    return
                except Exception:
                    continue
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
    threading.Thread(target=_fn, daemon=True).start()

# ── 色彩主題 ──────────────────────────────────────────────
C = {
    "bg":              "#f0f2f8",
    "panel":           "#ffffff",
    "topbar":          "#1e293b",
    "border":          "#dde1ea",
    "accent":          "#2563eb",
    "accent2":         "#16a34a",
    "warn":            "#d97706",
    "text":            "#1e293b",
    "muted":           "#64748b",
    "preview_bg":      "#1a1a2e",
    "btn_hover":       "#1d4ed8",
    "btn_green":       "#16a34a",
    "btn_green_hover": "#15803d",
    "purple":          "#7c3aed",
    "purple_hover":    "#6d28d9",
}

_UI  = "Microsoft JhengHei UI"
_NUM = "Consolas"

F = {
    "title":      (_UI,  14, "bold"),
    "card_title": (_UI,   9, "bold"),
    "body":       (_UI,   9),
    "hint":       (_UI,   9),
    "val_sm":     (_NUM,  9),
    "val_md":     (_NUM,  9, "bold"),
    "btn_main":   (_UI,  11, "bold"),
    "btn_sm":     (_UI,   9),
    "status":     (_UI,   9),
}

DISPLAY_W, DISPLAY_H = 640, 480
SETTINGS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "qr_scanner_settings.json")


# ══════════════════════════════════════════════════════════
class QRExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📷 QR掃描 + Excel填入工具")
        self.root.configure(bg=C["bg"])
        self.root.resizable(True, True)
        self.root.minsize(1020, 620)

        # 鏡頭
        self.cap           = None
        self.running       = False
        self.camera_index  = tk.IntVar(value=0)

        # 掃碼
        self.scan_var      = tk.BooleanVar(value=False)
        self._scan_results = []
        self._last_code    = ""
        self._last_code_ts = 0.0

        # 填入目標（COM 游標模式）
        self.inject_mode_var   = tk.BooleanVar(value=True)   # 預設開啟填入
        self.inject_follow_var = tk.BooleanVar(value=False)  # 預設不自動下移
        self.follow_delay_var  = tk.DoubleVar(value=0.5)     # 下移延遲秒數
        self._cursor_info      = None
        self._cursor_track_job = None
        self._inject_history   = []   # 用於撤銷

        # 掃描節流設定
        self.scan_interval_var = tk.DoubleVar(value=5.0)  # 同碼冷卻秒數（預設5秒）
        self.scan_maxcount_var = tk.IntVar(value=0)       # 0=不限，>0=每碼上限次數
        self._code_count: dict = {}                       # {key: 掃到次數}
        self._code_ts:    dict = {}                       # {key: 上次掃到時間戳}

        # 篩選 / 擷取模式
        # filter_mode: "all" | "mac" | "sn"
        self.filter_mode_var   = tk.StringVar(value="all")
        # extract_mode: "full" | "macN"（MAC 後 N 碼，N 由 mac_digits_var 控制）
        self.extract_mode_var  = tk.StringVar(value="full")
        # MAC 後幾碼（2 / 4 / 6 / 8 / 12）
        self.mac_digits_var    = tk.IntVar(value=4)

        # 全部掃交替填入（MAC → SN → MAC → SN ...）
        # alt_fill_var: True = 啟用交替填入
        self.alt_fill_var      = tk.BooleanVar(value=False)
        # 交替狀態：0 = 下一筆填 MAC（游標欄），1 = 下一筆填 SN（游標欄+1）
        self._alt_fill_step    = 0   # 0=等待MAC, 1=等待SN
        self._alt_fill_base_row = None  # 本輪起始列
        self._alt_fill_base_col = None  # 本輪起始欄（MAC 欄）
        self._alt_fill_locked  = False  # MAC 填完後短暫鎖定，防止 SN 搶先

        # 解析度
        self._res_options = [
            ("2560 × 1440 (2K)", 2560, 1440),
            ("1920 × 1080 (FHD)", 1920, 1080),
            ("1280 × 720 (HD)",   1280,  720),
            ("640 × 480 (SD)",     640,  480),
        ]
        self.res_var = tk.StringVar(value=self._res_options[0][0])  # 預設 2K

        # 數位放大
        self.zoom_var = tk.DoubleVar(value=1.0)   # 1.0 ~ 4.0
        self._zoom_cx = 0.5   # 放大中心 X（0~1，相對於原始畫面）
        self._zoom_cy = 0.5   # 放大中心 Y

        # 焦距固定（關閉自動對焦）
        self.focus_var     = tk.IntVar(value=0)     # 0 ~ 255
        self.autofocus_var = tk.BooleanVar(value=True)   # True = 預設開啟自動對焦

        # 閒置自動鎖焦：避免 AF 在無對焦目標時漂移造成掃碼失焦
        self._last_activity_ts   = time.time()   # 最後有掃碼或操作的時間
        self._idle_lock_seconds  = 30.0          # 閒置超過 N 秒後自動鎖定焦距
        self._idle_af_locked     = False         # 目前是否處於閒置鎖焦狀態
        self._idle_locked_focus  = 0             # 鎖定時儲存的焦距值
        self._idle_lock_job      = None          # 閒置鎖焦計時器
        # 黃金焦距：每次成功掃碼時記錄的清晰焦距，閒置鎖焦優先用這個值
        self._golden_focus       = None          # None = 尚未學習到清晰焦距
        self._golden_focus_count = 0             # 學習次數（累積愈多愈可靠）

        # 光線補償
        self.auto_exposure_var  = tk.BooleanVar(value=True)   # 自動曝光補償開關
        self.exposure_var       = tk.IntVar(value=-6)          # 手動曝光值（-13 ~ 0）
        self.brightness_var     = tk.IntVar(value=128)         # 亮度（0 ~ 255）
        self._ae_job            = None                         # 自動曝光補償計時器

        # 條碼強化模式（Code128 等長窄線性條碼）
        self.barcode_mode_var = tk.BooleanVar(value=True)   # 預設開啟：強化細條紋 Code128 辨識率

        # 成功填入 Toast 通知狀態
        self._toast_win     = None
        self._toast_job     = None
        self.toast_enabled_var = tk.BooleanVar(value=True)   # 開/關 Toast
        self._toast_pinned  = True     # Toast 是否釘選（不自動消失）
        self._toast_x       = None     # 記憶位置 X
        self._toast_y       = None     # 記憶位置 Y

        # 背景掃描執行緒
        self._scan_thread_running = False
        self._latest_frame        = None   # 主執行緒寫入，掃描執行緒讀取
        self._latest_codes        = []     # 掃描執行緒寫入，主執行緒讀取
        self._frame_lock          = threading.Lock()
        self._codes_lock          = threading.Lock()
        self._new_frame_event     = threading.Event()  # 有新幀才喚醒掃描執行緒

        # 效能控制
        self._frame_interval = 30        # 正常更新間隔 ms（約 33 FPS）
        self._minimized      = False     # 視窗是否最小化

        # 狀態
        self.status_var  = tk.StringVar(value="初始化中...")

        self._load_settings()   # ← 載入上次儲存的設定
        if not hasattr(self, "_compact_pos"):
            self._compact_pos = None  # 尚無記錄時，啟動後靠右上角
        self._build_ui()
        # UI 建立後，同步顯示標籤（_load_settings 在 UI 前執行，標籤尚不存在）
        self._on_zoom_change()
        self._on_delay_change()
        self._on_focus_change()   # 同步焦距滑桿 enable/disable 狀態
        self._auto_install_check()
        self._start_camera()
        # 預設開啟填入模式（UI 建好後執行才能更新按鈕狀態）
        self.root.after(100, self._ensure_inject_on)
        # 套用釘選狀態
        if self._pinned:
            self.root.after(150, lambda: (
                self.root.wm_attributes("-topmost", True),
                self.pin_btn.config(bg=C["accent2"],
                                    activebackground=C["btn_green_hover"])
            ))
        # 啟動後自動切換簡易模式並靠右
        self.root.after(200, self._startup_compact_right)

    # ── 設定儲存 / 載入 ───────────────────────────────────
    def _save_settings(self):
        data = {
            "camera_index":      self.camera_index.get(),
            "res_var":           self.res_var.get(),
            "scan_interval":     self.scan_interval_var.get(),
            "filter_mode":       self.filter_mode_var.get(),
            "extract_mode":      self.extract_mode_var.get(),
            "mac_digits":        self.mac_digits_var.get(),
            "alt_fill":          self.alt_fill_var.get(),
            "inject_follow":     self.inject_follow_var.get(),
            "follow_delay":      self.follow_delay_var.get(),
            "zoom":              self.zoom_var.get(),
            "zoom_cx":           self._zoom_cx,
            "zoom_cy":           self._zoom_cy,
            "focus":             self.focus_var.get(),
            "autofocus":         self.autofocus_var.get(),
            "auto_exposure":     self.auto_exposure_var.get(),
            "exposure":          self.exposure_var.get(),
            "brightness":        self.brightness_var.get(),
            "barcode_mode":      self.barcode_mode_var.get(),
            "toast_enabled":     self.toast_enabled_var.get(),
            "toast_pinned":      self._toast_pinned,
            "toast_x":           self._toast_x,
            "toast_y":           self._toast_y,
            "pinned":            self._pinned,
            "compact_pos":       getattr(self, "_compact_pos", None),
        }
        # 檔案寫入搬到背景執行緒，避免磁碟 I/O 阻塞主執行緒 cap.read() 節奏
        def _write():
            try:
                with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                self.root.after(0, lambda: self.status_var.set("✔ 設定已儲存"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("儲存失敗", str(e)))
        threading.Thread(target=_write, daemon=True).start()

    def _load_settings(self):
        if not os.path.exists(SETTINGS_PATH):
            return
        try:
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.camera_index.set(     data.get("camera_index",  0))
            self.res_var.set(          data.get("res_var",        self._res_options[0][0]))
            self.scan_interval_var.set(data.get("scan_interval",  5.0))
            self.filter_mode_var.set(  data.get("filter_mode",    "all"))
            self.extract_mode_var.set( data.get("extract_mode",   "full"))
            self.mac_digits_var.set(   data.get("mac_digits",     4))
            self.alt_fill_var.set(     data.get("alt_fill",       False))
            self.inject_follow_var.set(data.get("inject_follow",  False))
            self.follow_delay_var.set( data.get("follow_delay",   0.5))
            self.zoom_var.set(         data.get("zoom",           1.0))
            self._zoom_cx =            data.get("zoom_cx",        0.5)
            self._zoom_cy =            data.get("zoom_cy",        0.5)
            self.focus_var.set(        data.get("focus",          0))
            self.autofocus_var.set(    data.get("autofocus",      True))
            self.auto_exposure_var.set(data.get("auto_exposure",  True))
            self.exposure_var.set(     data.get("exposure",       -6))
            self.brightness_var.set(   data.get("brightness",     128))
            self.barcode_mode_var.set( data.get("barcode_mode",   False))
            self.toast_enabled_var.set(data.get("toast_enabled",  True))
            self._toast_pinned =       data.get("toast_pinned",   True)
            self._toast_x =            data.get("toast_x",        None)
            self._toast_y =            data.get("toast_y",        None)
            self._pinned =             data.get("pinned",         True)
            cp = data.get("compact_pos", None)
            # 舊格式只有 (x, y) 兩個值，捨棄讓程式重新計算大小
            self._compact_pos = cp if (cp and len(cp) == 4) else None
        except Exception:
            pass   # 設定檔損毀就用預設值

    # ── 自動安裝檢查 ──────────────────────────────────────
    def _auto_install_check(self):
        missing = []
        if not PYZBAR_OK:    missing.append("pyzbar")
        if not OPENPYXL_OK:  missing.append("openpyxl")
        if not PIL_OK:       missing.append("Pillow")
        if missing:
            msg = "缺少套件：" + "、".join(missing) + "\n是否立即自動安裝？"
            if messagebox.askyesno("套件安裝", msg):
                self._install_packages(missing)

    def _install_packages(self, pkgs):
        self.status_var.set("⏳ 安裝套件中，請稍候...")
        def _do():
            for pkg in pkgs:
                try:
                    subprocess.check_call(
                        [sys.executable, "-m", "pip", "install", pkg],
                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                except Exception as e:
                    self.root.after(0, lambda e=e: messagebox.showerror(
                        "安裝失敗", str(e)))
            self.root.after(0, lambda: messagebox.showinfo(
                "安裝完成", "套件已安裝，請重新啟動程式"))
            self.root.after(0, lambda: self.status_var.set("✔ 安裝完成，請重新啟動"))
        threading.Thread(target=_do, daemon=True).start()

    # ── 建立 UI ───────────────────────────────────────────
    def _build_ui(self):
        # ── 頂列 ──────────────────────────────────────────
        self._topbar = tk.Frame(self.root, bg=C["topbar"], height=48)
        self._topbar.pack(fill="x", side="top")
        self._topbar.pack_propagate(False)

        self._topbar_title = tk.Label(self._topbar, text="📷  QR掃描 + Excel自動填入",
                 font=F["title"], fg="white",
                 bg=C["topbar"])
        self._topbar_title.pack(side="left", padx=16, pady=10)

        # 簡易模式切換（最右邊，先 pack）
        self._compact_mode = False
        self.compact_btn = tk.Button(
            self._topbar, text="⊟  簡易",
            font=F["hint"], bg="#475569", fg="white",
            activebackground="#334155", activeforeground="white",
            bd=0, padx=10, pady=4, cursor="hand2",
            command=self._toggle_compact)
        self.compact_btn.pack(side="right", padx=(0, 8), pady=10)

        # 釘選（always on top）—— 注意：_pinned 已由 _load_settings 設定，勿覆蓋
        # 首次執行（設定檔不存在）時預設為釘選
        if not hasattr(self, "_pinned"):
            self._pinned = True
        self.pin_btn = tk.Button(
            self._topbar, text="📌",
            font=F["hint"], bg="#475569", fg="white",
            activebackground="#334155", activeforeground="white",
            bd=0, padx=8, pady=4, cursor="hand2",
            command=self._toggle_pin)
        self.pin_btn.pack(side="right", padx=(0, 4), pady=10)

        # 簡易模式專用：重新啟動按鈕（放 topbar 左側，切換時才顯示）
        self._topbar_restart_btn = tk.Button(
            self._topbar, text="🔄  重新啟動",
            font=F["hint"], bg="#b91c1c", fg="white",
            activebackground="#991b1b", activeforeground="white",
            bd=0, padx=12, pady=4, cursor="hand2",
            command=self._restart_app)
        # 預設不 pack，切換到簡易模式時才顯示

        # ── 鏡頭 / 解析度工具列（topbar 下方獨立一行）────────
        self._topbar_ctrl = tk.Frame(self.root, bg="#0f172a")
        self._topbar_ctrl.pack(fill="x", side="top")
        ctrl = self._topbar_ctrl

        tk.Label(ctrl, text="鏡頭", font=F["hint"],
                 fg="#94a3b8", bg="#0f172a").pack(side="left", padx=(14, 4), pady=6)
        tk.Spinbox(ctrl, from_=0, to=4,
                   textvariable=self.camera_index,
                   width=3, font=F["val_sm"],
                   bg="#334155", fg="white", relief="flat", bd=1,
                   command=self._restart_camera).pack(side="left", pady=6)
        tk.Button(ctrl, text="重新連線",
                  font=F["hint"], bg=C["accent"], fg="white",
                  activebackground=C["btn_hover"], activeforeground="white",
                  bd=0, padx=10, pady=3, cursor="hand2",
                  command=self._restart_camera).pack(side="left", padx=(8, 20), pady=6)

        tk.Label(ctrl, text="解析度", font=F["hint"],
                 fg="#94a3b8", bg="#0f172a").pack(side="left", padx=(0, 4), pady=6)
        res_names = [r[0] for r in self._res_options]
        self._res_combo = ttk.Combobox(ctrl, textvariable=self.res_var,
                                       values=res_names, state="readonly",
                                       width=18, font=F["hint"])
        self._res_combo.pack(side="left", pady=6)
        tk.Button(ctrl, text="套用",
                  font=F["hint"], bg=C["btn_green"], fg="white",
                  activebackground=C["btn_green_hover"], activeforeground="white",
                  bd=0, padx=12, pady=3, cursor="hand2",
                  command=self._restart_camera).pack(side="left", padx=(8, 0), pady=6)

        # ── 自動對焦控制 ──────────────────────────────────
        self._af_chk = tk.Checkbutton(
            ctrl, text="自動對焦",
            variable=self.autofocus_var,
            font=F["hint"], fg="#94a3b8",
            bg="#0f172a", selectcolor="#334155",
            activebackground="#0f172a", activeforeground="white",
            command=self._on_focus_change)
        self._af_chk.pack(side="left", padx=(16, 0), pady=6)

        tk.Button(
            ctrl, text="🔄 重新啟動",
            font=F["hint"], bg="#b91c1c", fg="white",
            activebackground="#991b1b", activeforeground="white",
            bd=0, padx=10, pady=3, cursor="hand2",
            command=self._restart_app
        ).pack(side="left", padx=(8, 0), pady=6)

        tk.Checkbutton(
            ctrl, text="🔲 條碼強化",
            variable=self.barcode_mode_var,
            font=F["hint"], fg="#fbbf24",
            bg="#0f172a", selectcolor="#334155",
            activebackground="#0f172a", activeforeground="white"
        ).pack(side="left", padx=(12, 0), pady=6)

        tk.Checkbutton(
            ctrl, text="🔔 掃描提示",
            variable=self.toast_enabled_var,
            font=F["hint"], fg="#7dd3fc",
            bg="#0f172a", selectcolor="#334155",
            activebackground="#0f172a", activeforeground="white"
        ).pack(side="left", padx=(12, 0), pady=6)

        # ── 主體 ──────────────────────────────────────────
        self._body = tk.Frame(self.root, bg=C["bg"])
        self._body.pack(fill="both", expand=True, padx=10, pady=8)

        # 左：預覽
        self._left = tk.Frame(self._body, bg=C["bg"])
        self._left.pack(side="left", fill="both", expand=True, padx=(0, 8))
        left = self._left

        self._preview_wrap = tk.Frame(left, bg=C["preview_bg"],
                                highlightbackground=C["border"],
                                highlightthickness=1)
        self._preview_wrap.pack()
        self.video_label = tk.Label(self._preview_wrap, bg=C["preview_bg"],
                                    width=DISPLAY_W, height=DISPLAY_H)
        self.video_label.pack()
        self.video_label.bind("<Button-1>", self._on_video_click)
        self.video_label.bind("<Button-3>", self._zoom_reset)
        self.root.bind("<Configure>", self._on_window_resize)

        # 狀態列
        self._status_bar = tk.Frame(left, bg=C["panel"],
                              highlightbackground=C["border"],
                              highlightthickness=1)
        self._status_bar.pack(fill="x", pady=(6, 0))
        tk.Label(self._status_bar, textvariable=self.status_var,
                 font=F["status"], fg=C["accent2"],
                 bg=C["panel"]).pack(side="left", padx=12, pady=5)

        # 簡易模式專用按鈕列（獨立 Frame，平時隱藏，不受 body 高度限制）
        self._compact_btn_bar = tk.Frame(self.root, bg=C["bg"])
        # 預設不 pack，切換時才顯示

        self._compact_scan_btn = tk.Button(
            self._compact_btn_bar, text="▶  開始掃描",
            font=F["btn_main"],
            bg=C["purple"], fg="white",
            activebackground=C["purple_hover"], activeforeground="white",
            bd=0, pady=6, cursor="hand2",
            command=self._toggle_scan)
        self._compact_scan_btn.pack(fill="x")

        # 右：控制欄（可捲動）
        self._right_outer = tk.Frame(self._body, bg=C["bg"], width=360)
        self._right_outer.pack(side="right", fill="y")
        self._right_outer.pack_propagate(False)

        canvas = tk.Canvas(self._right_outer, bg=C["bg"],
                           width=360, highlightthickness=0)
        vsb = ttk.Scrollbar(self._right_outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        self.sf = tk.Frame(canvas, bg=C["bg"])
        sf_id = canvas.create_window((0, 0), window=self.sf, anchor="nw")

        self.sf.bind("<Configure>",
                     lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(sf_id, width=e.width))
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        self._build_scan_card()
        self._build_zoom_card()          # ← 數位放大
        self._build_light_card()         # ← 光線補償
        self._build_scan_settings_card()
        self._build_filter_card()        # ← 新增篩選卡片
        self._build_excel_card()
        self._build_history_card()

        # ── 底列 ──────────────────────────────────────────
        self._btn_bar = tk.Frame(self.root, bg=C["panel"],
                           highlightbackground=C["border"],
                           highlightthickness=1)
        self._btn_bar.pack(side="bottom", fill="x")

        self.scan_main_btn = tk.Button(
            self._btn_bar, text="▶  開始掃描",
            font=F["btn_main"],
            bg=C["purple"], fg="white",
            activebackground=C["purple_hover"], activeforeground="white",
            bd=0, padx=28, pady=10, cursor="hand2",
            command=self._toggle_scan)
        self.scan_main_btn.pack(side="left", padx=(14, 8), pady=8)

        self._pause_btn = tk.Button(self._btn_bar, text="⏸  暫停預覽",
                  font=F["btn_sm"],
                  bg=C["bg"], fg=C["text"],
                  activebackground=C["border"],
                  highlightbackground=C["border"], highlightthickness=1,
                  bd=0, padx=12, pady=10, cursor="hand2",
                  command=self._toggle_camera)
        self._pause_btn.pack(side="left", pady=8)

        tk.Button(self._btn_bar, text="💾  儲存設定",
                  font=F["btn_sm"],
                  bg=C["bg"], fg=C["text"],
                  activebackground=C["border"],
                  highlightbackground=C["border"], highlightthickness=1,
                  bd=0, padx=12, pady=10, cursor="hand2",
                  command=self._save_settings
                  ).pack(side="right", padx=(0, 14), pady=8)

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.bind("<Escape>", lambda e: self._on_close())
        self.root.bind("<r>", self._zoom_reset)
        self.root.bind("<R>", self._zoom_reset)

    # ── 掃碼卡片 ──────────────────────────────────────────
    def _build_scan_card(self):
        card = self._make_card(pady_top=0)
        title_row = tk.Frame(card, bg=C["panel"])
        title_row.pack(fill="x", padx=12, pady=(10, 4))
        tk.Label(title_row, text="📷  QR / 條碼掃描",
                 font=F["card_title"], fg=C["purple"],
                 bg=C["panel"]).pack(side="left")
        self.scan_state_lbl = tk.Label(title_row, text="● 關閉",
                                       font=F["hint"], fg=C["muted"],
                                       bg=C["panel"])
        self.scan_state_lbl.pack(side="right")

        if not PYZBAR_OK:
            tk.Button(card, text="⬇ 安裝 pyzbar",
                      font=F["btn_sm"], bg=C["warn"], fg="white",
                      activebackground="#b45309", bd=0,
                      padx=10, pady=4, cursor="hand2",
                      command=lambda: self._install_packages(["pyzbar"])
                      ).pack(anchor="w", padx=12, pady=(0, 4))

        # 掃碼紀錄清單
        tk.Label(card, text="掃碼紀錄（最新在上）",
                 font=F["hint"], fg=C["muted"],
                 bg=C["panel"]).pack(anchor="w", padx=12, pady=(4, 2))

        list_frame = tk.Frame(card, bg=C["panel"])
        list_frame.pack(fill="x", padx=12, pady=(0, 4))
        sb = ttk.Scrollbar(list_frame, orient="vertical")
        self.scan_listbox = tk.Listbox(
            list_frame, height=6,
            font=F["hint"], fg=C["text"], bg="#f8fafc",
            selectbackground=C["accent"], selectforeground="white",
            relief="flat", bd=0,
            highlightbackground=C["border"], highlightthickness=1,
            yscrollcommand=sb.set)
        sb.config(command=self.scan_listbox.yview)
        sb.pack(side="right", fill="y")
        self.scan_listbox.pack(side="left", fill="x", expand=True)

        btn_row = tk.Frame(card, bg=C["panel"])
        btn_row.pack(fill="x", padx=12, pady=(0, 10))
        tk.Button(btn_row, text="📋 複製",
                  font=F["btn_sm"], bg=C["accent"], fg="white",
                  activebackground=C["btn_hover"], bd=0,
                  padx=10, pady=3, cursor="hand2",
                  command=self._copy_selected).pack(side="left", padx=(0, 6))
        tk.Button(btn_row, text="🗑 清除",
                  font=F["btn_sm"], bg=C["bg"], fg=C["text"],
                  activebackground=C["border"],
                  highlightbackground=C["border"], highlightthickness=1,
                  bd=0, padx=10, pady=3, cursor="hand2",
                  command=self._clear_scan).pack(side="left")
        tk.Button(btn_row, text="📤 匯出CSV",
                  font=F["btn_sm"], bg=C["bg"], fg=C["text"],
                  activebackground=C["border"],
                  highlightbackground=C["border"], highlightthickness=1,
                  bd=0, padx=10, pady=3, cursor="hand2",
                  command=self._export_csv).pack(side="right")

    # ── 游標填入卡片 ──────────────────────────────────────
    def _build_excel_card(self):
        card = self._make_card()

        # 標題 + 開關（注意：side="right" 的元件必須先 pack）
        title_row = tk.Frame(card, bg=C["panel"])
        title_row.pack(fill="x", padx=12, pady=(10, 4))
        self.inject_btn = tk.Button(
            title_row, text="○ 關閉",
            font=F["hint"], bg="#94a3b8", fg="white",
            activebackground=C["border"], bd=0,
            padx=10, pady=2, cursor="hand2",
            command=self._toggle_inject_mode)
        self.inject_btn.pack(side="right")
        tk.Label(title_row, text="🖱  游標填入（Excel / Word）",
                 font=F["card_title"], fg=C["accent2"],
                 bg=C["panel"]).pack(side="left")

        # pywin32 缺失提示
        if not WIN32_OK:
            tk.Button(card, text="⬇ 安裝 pywin32（必要）",
                      font=F["btn_sm"], bg=C["warn"], fg="white",
                      activebackground="#b45309", bd=0,
                      padx=10, pady=4, cursor="hand2",
                      command=lambda: self._install_packages(["pywin32"])
                      ).pack(anchor="w", padx=12, pady=(0, 4))

        # 使用說明
        tk.Label(card,
                 text="不需開啟指定檔案。\n"
                      "在 Excel/Word 點好儲存格後，\n"
                      "掃描條碼即自動填入游標位置。",
                 font=F["hint"], fg=C["muted"], bg=C["panel"],
                 justify="left").pack(anchor="w", padx=12, pady=(0, 6))

        # 游標狀態顯示
        self.cursor_lbl = tk.Label(
            card, text="⏳ 尚未偵測到 Office 視窗",
            font=F["hint"], fg=C["muted"], bg=C["panel"],
            justify="left", anchor="w", wraplength=280)
        self.cursor_lbl.pack(fill="x", padx=12, pady=(0, 6))

        # 填入後自動下移（Excel）+ 延遲設定
        follow_row = tk.Frame(card, bg=C["panel"])
        follow_row.pack(fill="x", padx=12, pady=(2, 2))
        tk.Checkbutton(
            follow_row, text="填入後自動下移",
            variable=self.inject_follow_var,
            font=F["hint"], fg=C["text"], bg=C["panel"],
            activebackground=C["panel"], cursor="hand2"
        ).pack(side="left")

        delay_row = tk.Frame(card, bg=C["panel"])
        delay_row.pack(fill="x", padx=20, pady=(0, 6))
        tk.Label(delay_row, text="延遲",
                 font=F["hint"], fg=C["muted"],
                 bg=C["panel"]).pack(side="left")
        tk.Scale(delay_row, from_=0.0, to=3.0, resolution=0.1,
                 variable=self.follow_delay_var,
                 orient="horizontal", length=150,
                 bg=C["panel"], highlightthickness=0,
                 troughcolor=C["border"], sliderrelief="flat",
                 command=self._on_delay_change,
                 ).pack(side="left", padx=(4, 2))
        self.delay_lbl = tk.Label(delay_row, text="0.5 秒",
                                  font=F["val_sm"], fg=C["accent"],
                                  bg=C["panel"], width=5)
        self.delay_lbl.pack(side="left")

        # 撤銷按鈕
        xl_btn_row = tk.Frame(card, bg=C["panel"])
        xl_btn_row.pack(fill="x", padx=12, pady=(0, 6))
        tk.Button(xl_btn_row, text="↩ 撤銷上一筆",
                  font=F["btn_sm"], bg=C["bg"], fg=C["text"],
                  activebackground=C["border"],
                  highlightbackground=C["border"], highlightthickness=1,
                  bd=0, padx=8, pady=3, cursor="hand2",
                  command=self._inject_undo).pack(side="left")

        # 診斷按鈕
        tk.Button(xl_btn_row, text="🔍 診斷",
                  font=F["btn_sm"], bg=C["warn"], fg="white",
                  activebackground="#b45309", bd=0,
                  padx=8, pady=3, cursor="hand2",
                  command=self._show_diagnose).pack(side="right", pady=(0, 0))

        tk.Label(card,
                 text="若顯示「未偵測到」，請按「診斷」\n確認 Excel/Word 已開啟並有活頁簿",
                 font=F["hint"], fg=C["muted"], bg=C["panel"],
                 justify="left").pack(anchor="w", padx=12, pady=(0, 10))

    # ── 篩選 / 擷取卡片 ───────────────────────────────────
    def _build_filter_card(self):
        card = self._make_card()
        tk.Label(card, text="🎯  條碼篩選 & 擷取",
                 font=F["card_title"], fg=C["purple"],
                 bg=C["panel"]).pack(anchor="w", padx=12, pady=(10, 6))

        # ── 篩選模式 ──────────────────────────────────────
        tk.Label(card, text="只掃哪條碼：",
                 font=F["hint"], fg=C["muted"],
                 bg=C["panel"]).pack(anchor="w", padx=12, pady=(0, 2))

        fm_row = tk.Frame(card, bg=C["panel"])
        fm_row.pack(fill="x", padx=16, pady=(0, 6))
        for val, lbl in [("all", "全部掃"),
                         ("mac", "只掃 MAC"),
                         ("sn",  "只掃 SN")]:
            tk.Radiobutton(
                fm_row, text=lbl,
                variable=self.filter_mode_var, value=val,
                font=F["hint"], fg=C["text"], bg=C["panel"],
                activebackground=C["panel"],
                selectcolor=C["panel"],
                cursor="hand2",
                command=self._on_filter_change,
            ).pack(side="left", padx=(0, 10))

        # 提示：如何辨識 MAC / SN
        self.filter_hint_lbl = tk.Label(
            card, text="",
            font=F["hint"], fg=C["accent"], bg=C["panel"],
            justify="left", wraplength=280)
        self.filter_hint_lbl.pack(fill="x", anchor="w", padx=12, pady=(0, 6))

        # ── 擷取模式 ──────────────────────────────────────
        sep = tk.Frame(card, bg=C["border"], height=1)
        sep.pack(fill="x", padx=12, pady=(2, 6))

        tk.Label(card, text="填入什麼內容：",
                 font=F["hint"], fg=C["muted"],
                 bg=C["panel"]).pack(anchor="w", padx=12, pady=(0, 2))

        em_row = tk.Frame(card, bg=C["panel"])
        em_row.pack(fill="x", padx=16, pady=(0, 4))
        self._extract_full_rb = tk.Radiobutton(
            em_row, text="完整內容",
            variable=self.extract_mode_var, value="full",
            font=F["hint"], fg=C["text"], bg=C["panel"],
            activebackground=C["panel"], selectcolor=C["panel"],
            cursor="hand2",
            command=self._on_extract_mode_change,
        )
        self._extract_full_rb.pack(side="left", padx=(0, 10))
        self._extract_macN_rb = tk.Radiobutton(
            em_row, text="MAC 後",
            variable=self.extract_mode_var, value="macN",
            font=F["hint"], fg=C["text"], bg=C["panel"],
            activebackground=C["panel"], selectcolor=C["panel"],
            cursor="hand2",
            command=self._on_extract_mode_change,
        )
        self._extract_macN_rb.pack(side="left", padx=(0, 2))

        # 位數微調器（2 / 4 / 6 / 8 / 12）
        self._mac_digits_spin = tk.Spinbox(
            em_row, from_=1, to=12, increment=1,
            textvariable=self.mac_digits_var,
            width=3, font=F["val_sm"],
            bg="#f1f5f9", fg=C["text"],
            relief="flat", bd=1,
            highlightbackground=C["border"], highlightthickness=1,
            command=self._on_digits_spin_change,
        )
        self._mac_digits_spin.pack(side="left", padx=(0, 2))
        tk.Label(em_row, text="碼",
                 font=F["hint"], fg=C["text"],
                 bg=C["panel"]).pack(side="left")

        # 預覽框：顯示上一筆原始值 → 擷取後值
        self.extract_preview_lbl = tk.Label(
            card, text="擷取預覽：（尚無掃碼紀錄）",
            font=F["val_sm"], fg=C["muted"], bg=C["panel"],
            justify="left", anchor="w", wraplength=280)
        self.extract_preview_lbl.pack(fill="x", padx=12, pady=(0, 10))

        self._on_filter_change()   # 初始化提示文字

        # ── 「全部掃」交替填入選項 ────────────────────────
        sep2 = tk.Frame(card, bg=C["border"], height=1)
        sep2.pack(fill="x", padx=12, pady=(2, 6))

        self._alt_fill_chk = tk.Checkbutton(
            card,
            text="全部掃：交替填入 MAC → SN\n（每掃完一對自動下移一列）",
            variable=self.alt_fill_var,
            font=F["hint"], fg=C["purple"], bg=C["panel"],
            activebackground=C["panel"], selectcolor=C["panel"],
            cursor="hand2", justify="left",
            wraplength=280,
            command=self._on_alt_fill_change,
        )
        self._alt_fill_chk.pack(anchor="w", padx=12, pady=(0, 4), fill="x")

        self._alt_fill_hint_lbl = tk.Label(
            card,
            text="",
            font=F["hint"], fg=C["muted"], bg=C["panel"],
            justify="left", wraplength=280)
        self._alt_fill_hint_lbl.pack(fill="x", anchor="w", padx=12, pady=(0, 8))
        self._on_alt_fill_change()   # 初始化提示

    def _on_filter_change(self, *_):
        m = self.filter_mode_var.get()
        hints = {
            "all": "偵測所有條碼，不篩選。",
            "mac": "只接受包含「MAC:」前綴或長度符合 MAC 的條碼。",
            "sn":  "只掃 SN（非 MAC 格式的條碼，如 21912A1393004010970）。",
        }
        self.filter_hint_lbl.config(text=hints.get(m, ""))
        # 只有「只掃 MAC」才能選 MAC 後 N 碼；其他模式強制完整內容並鎖定
        if hasattr(self, "_extract_full_rb") and hasattr(self, "_extract_macN_rb"):
            if m == "mac":
                self._extract_full_rb.config(state="normal")
                self._extract_macN_rb.config(state="normal")
                if hasattr(self, "_mac_digits_spin"):
                    self._mac_digits_spin.config(state="normal")
            else:
                self.extract_mode_var.set("full")
                self._extract_full_rb.config(state="disabled")
                self._extract_macN_rb.config(state="disabled")
                if hasattr(self, "_mac_digits_spin"):
                    self._mac_digits_spin.config(state="disabled")

    def _on_alt_fill_change(self, *_):
        """交替填入選項切換時更新提示、重置狀態。"""
        enabled = self.alt_fill_var.get()
        # 重置交替狀態
        self._alt_fill_step     = 0
        self._alt_fill_base_row = None
        self._alt_fill_base_col = None
        self._alt_fill_locked   = False
        if hasattr(self, "_alt_fill_hint_lbl"):
            if enabled:
                self._alt_fill_hint_lbl.config(
                    text="游標欄=MAC、下一欄=SN，\n掃完一對自動下移。\n"
                         "（需搭配「全部掃」模式）",
                    fg=C["purple"])
            else:
                self._alt_fill_hint_lbl.config(text="", fg=C["muted"])

    def _reset_alt_fill(self):
        """外部可呼叫：手動重置交替狀態（例如開始新一批掃描）。"""
        self._alt_fill_step     = 0
        self._alt_fill_base_row = None
        self._alt_fill_base_col = None
        self._alt_fill_locked   = False
        self.status_var.set("🔄 交替填入狀態已重置")

    def _on_extract_mode_change(self, *_):
        """radio button 切換時更新預覽（不強制切換模式）。"""
        self._update_extract_preview()

    def _on_digits_spin_change(self, *_):
        """點 spinbox 時：自動切換 radio 到 macN，並更新預覽。"""
        # 點 spinbox 不會自動切 radio，幫它切
        self.extract_mode_var.set("macN")
        self._update_extract_preview()

    def _update_extract_preview(self):
        """更新擷取預覽標籤。"""
        if self._scan_results:
            last_entry = self._scan_results[-1]
            # _scan_results 元素為 dict（有 "raw" 欄位）或純字串（相容舊版）
            last_raw = last_entry.get("raw", "") if isinstance(last_entry, dict) else last_entry
            extracted = self._apply_extract(last_raw)
            mode = self.extract_mode_var.get()
            if mode == "full":
                self.extract_preview_lbl.config(
                    text=f"擷取預覽（完整）：\n{last_raw}",
                    fg=C["accent"])
            else:
                n = self.mac_digits_var.get()
                self.extract_preview_lbl.config(
                    text=f"擷取預覽（後 {n} 碼）：\n{last_raw}  →  {extracted}",
                    fg=C["accent"])

    def _apply_filter(self, raw_data: str) -> bool:
        """
        依 filter_mode_var 決定這筆資料是否通過篩選。
        True = 接受，False = 忽略。
        """
        mode = self.filter_mode_var.get()
        if mode == "all":
            return True
        d = raw_data.strip()
        if mode == "mac":
            # 接受：以 "MAC:" 開頭
            if d.upper().startswith("MAC:"):
                return True
            # 接受：整串去掉分隔符後剛好 12 位 hex（純 MAC 位址，無其他內容）
            clean = d.replace(":", "").replace("-", "").replace(" ", "")
            if len(clean) == 12 and all(c in "0123456789ABCDEFabcdef" for c in clean):
                return True
            # 接受：標準 MAC 位址格式 XX:XX:XX:XX:XX:XX 或 XX-XX-XX-XX-XX-XX
            import re
            if re.search(r'^([0-9A-Fa-f]{2}[:\-]){5}[0-9A-Fa-f]{2}$', d):
                return True
            return False
        if mode == "sn":
            # SN 條碼本身不含 "SN:" 前綴，直接判斷：不是 MAC 格式就接受
            import re
            # 排除 MAC 格式（12位hex，含或不含分隔符）
            clean = d.replace(":", "").replace("-", "").replace(" ", "")
            is_mac = (len(clean) == 12 and
                      all(c in "0123456789ABCDEFabcdef" for c in clean))
            if is_mac:
                return False
            if re.search(r'^([0-9A-Fa-f]{2}[:\-]){5}[0-9A-Fa-f]{2}$', d):
                return False
            # 接受：英數字串（可含 - _ .），長度 ≥ 1
            return True
        return True

    def _apply_extract(self, raw_data: str) -> str:
        """
        依 extract_mode_var 決定最終填入的字串。
        full  → 原始字串
        macN  → MAC 後 N 碼（大寫，無分隔符），N 由 mac_digits_var 控制
        """
        mode = self.extract_mode_var.get()
        if mode == "full":
            return raw_data
        if mode in ("macN", "mac4"):   # 相容舊設定值 "mac4"
            import re
            n = self.mac_digits_var.get()
            d = raw_data.strip()
            # 去掉 "MAC:" 前綴
            if d.upper().startswith("MAC:"):
                d = d[4:]
            # 優先：找標準 MAC 位址格式（XX:XX:XX:XX:XX:XX 或 XX-XX-XX-XX-XX-XX）
            m = re.search(r'([0-9A-Fa-f]{2}[:\-]){5}([0-9A-Fa-f]{2})', d)
            if m:
                clean = re.sub(r'[:\-]', '', m.group(0)).upper()
                return clean[-n:]
            # 次選：找連續 12 位 hex
            m = re.search(r'[0-9A-Fa-f]{12}', d)
            if m:
                return m.group(0)[-n:].upper()
            # 最後：去掉所有分隔符取最後 N 碼
            clean = d.replace(":", "").replace("-", "").replace(" ", "").upper()
            return clean[-n:] if len(clean) >= n else clean
        return raw_data

    # ── 數位放大卡片 ──────────────────────────────────────
    def _build_zoom_card(self):
        card = self._make_card()
        tk.Label(card, text="🔍  數位放大鏡",
                 font=F["card_title"], fg=C["text"],
                 bg=C["panel"]).pack(anchor="w", padx=12, pady=(10, 4))

        slider_row = tk.Frame(card, bg=C["panel"])
        slider_row.pack(fill="x", padx=12, pady=(0, 2))
        tk.Label(slider_row, text="1×",
                 font=F["hint"], fg=C["muted"],
                 bg=C["panel"]).pack(side="left")
        tk.Scale(slider_row, from_=1.0, to=4.0, resolution=0.5,
                 variable=self.zoom_var,
                 orient="horizontal", length=130,
                 bg=C["panel"], highlightthickness=0,
                 troughcolor=C["border"], sliderrelief="flat",
                 showvalue=False,
                 command=self._on_zoom_change
                 ).pack(side="left", padx=(4, 4))
        tk.Label(slider_row, text="4×",
                 font=F["hint"], fg=C["muted"],
                 bg=C["panel"]).pack(side="left")

        self.zoom_lbl = tk.Label(card,
                                 text="目前：1.0×",
                                 font=F["val_sm"], fg=C["accent"],
                                 bg=C["panel"])
        self.zoom_lbl.pack(anchor="w", padx=12, pady=(0, 4))

        tk.Label(card,
                 text="◈ 點擊畫面移動中心　置重中心 [R]",
                 font=F["hint"], fg=C["muted"],
                 bg=C["panel"]).pack(anchor="w", padx=12, pady=(0, 10))

    def _on_zoom_change(self, *_):
        v = self.zoom_var.get()
        self.zoom_lbl.config(text=f"目前：{v:.1f}×")

    # ── 光線補償卡片 ──────────────────────────────────────
    def _build_light_card(self):
        card = self._make_card()
        tk.Label(card, text="💡  光線補償",
                 font=F["card_title"], fg=C["warn"],
                 bg=C["panel"]).pack(anchor="w", padx=12, pady=(10, 6))

        # 自動曝光補償開關
        ae_row = tk.Frame(card, bg=C["panel"])
        ae_row.pack(fill="x", padx=12, pady=(0, 4))
        tk.Checkbutton(
            ae_row, text="自動曝光補償（光線不足時穩定 AF）",
            variable=self.auto_exposure_var,
            font=F["hint"], fg=C["text"], bg=C["panel"],
            activebackground=C["panel"], cursor="hand2",
            command=self._on_light_change,
        ).pack(side="left")

        # 亮度滑桿
        br_row = tk.Frame(card, bg=C["panel"])
        br_row.pack(fill="x", padx=12, pady=(0, 2))
        tk.Label(br_row, text="亮度",
                 font=F["hint"], fg=C["muted"],
                 bg=C["panel"], width=4).pack(side="left")
        tk.Scale(br_row, from_=0, to=255, resolution=1,
                 variable=self.brightness_var,
                 orient="horizontal", length=180,
                 bg=C["panel"], highlightthickness=0,
                 troughcolor=C["border"], sliderrelief="flat",
                 showvalue=False,
                 command=self._on_light_change,
                 ).pack(side="left", padx=(4, 4))
        self.brightness_lbl = tk.Label(br_row, text="128",
                                       font=F["val_sm"], fg=C["accent"],
                                       bg=C["panel"], width=4)
        self.brightness_lbl.pack(side="left")

        # 曝光滑桿（手動模式才有效）
        ex_row = tk.Frame(card, bg=C["panel"])
        ex_row.pack(fill="x", padx=12, pady=(0, 2))
        tk.Label(ex_row, text="曝光",
                 font=F["hint"], fg=C["muted"],
                 bg=C["panel"], width=4).pack(side="left")
        self._exposure_scale = tk.Scale(
            ex_row, from_=-13, to=0, resolution=1,
            variable=self.exposure_var,
            orient="horizontal", length=180,
            bg=C["panel"], highlightthickness=0,
            troughcolor=C["border"], sliderrelief="flat",
            showvalue=False,
            command=self._on_light_change,
        )
        self._exposure_scale.pack(side="left", padx=(4, 4))
        self.exposure_lbl = tk.Label(ex_row, text="-6",
                                     font=F["val_sm"], fg=C["accent"],
                                     bg=C["panel"], width=4)
        self.exposure_lbl.pack(side="left")

        tk.Label(card,
                 text="光線昏暗時調高亮度；自動補償會偵測畫面亮度自動調整曝光",
                 font=F["hint"], fg=C["muted"], bg=C["panel"],
                 justify="left", wraplength=280,
                 ).pack(anchor="w", padx=12, pady=(2, 10))

        self._on_light_change()   # 初始化標籤

    def _on_light_change(self, *_):
        """套用光線設定到鏡頭，並更新標籤。"""
        bv = self.brightness_var.get()
        ev = self.exposure_var.get()
        self.brightness_lbl.config(text=str(bv))
        self.exposure_lbl.config(text=str(ev))
        self._apply_light()
        # 自動補償開啟時啟動偵測迴圈，關閉時停止
        if self.auto_exposure_var.get():
            self._start_ae_loop()
        else:
            self._stop_ae_loop()

    def _apply_light(self):
        """把目前亮度/曝光設定寫入鏡頭。"""
        if not self.cap or not self.cap.isOpened():
            return
        bv = self.brightness_var.get()
        self.cap.set(cv2.CAP_PROP_BRIGHTNESS, bv)
        if not self.auto_exposure_var.get():
            # 手動曝光：關閉鏡頭自動曝光，寫入指定值
            self.cap.set(cv2.CAP_PROP_AUTO_EXPOSURE, 1)   # 1 = manual (DSHOW)
            self.cap.set(cv2.CAP_PROP_EXPOSURE, self.exposure_var.get())
        else:
            # 自動曝光：交給鏡頭，只控制亮度偏移
            self.cap.set(cv2.CAP_PROP_AUTO_EXPOSURE, 3)   # 3 = auto (DSHOW)

    def _start_ae_loop(self):
        """啟動自動曝光補償迴圈（每 1.5 秒偵測畫面亮度，自動微調曝光）。"""
        self._stop_ae_loop()
        self._ae_loop()

    def _stop_ae_loop(self):
        if self._ae_job:
            self.root.after_cancel(self._ae_job)
            self._ae_job = None

    # ── 焦距巡邏（防止驅動偷改 autofocus）─────────────────
    def _start_focus_guardian(self):
        """啟動焦距巡邏：每 3 秒主動確認驅動的 autofocus 狀態，若被重置就補回來。"""
        self._stop_focus_guardian()
        self._focus_guardian_job = None
        self._focus_guardian()

    def _stop_focus_guardian(self):
        job = getattr(self, "_focus_guardian_job", None)
        if job:
            self.root.after_cancel(job)
        self._focus_guardian_job = None
        # 同時停止閒置鎖焦計時器
        idle_job = getattr(self, "_idle_lock_job", None)
        if idle_job:
            self.root.after_cancel(idle_job)
        self._idle_lock_job = None

    def _focus_guardian(self):
        """定期巡邏：不管有無掃碼動作，每 1.5 秒主動確保焦距正確。
        固定焦距模式：讀回 AF 狀態 + 補送 FOCUS；偵測清晰度驟降時立即補送。
        自動對焦模式：不干預，讓驅動自行處理。
        """
        if not self.cap or not self.cap.isOpened() or not self.running:
            return

        if not self.autofocus_var.get():
            # ── 固定焦距模式 ────────────────────────────────
            # 1. 讀回驅動 AF 狀態，被偷改就壓回來
            actual_af = self.cap.get(cv2.CAP_PROP_AUTOFOCUS)
            if actual_af != 0:
                self.cap.set(cv2.CAP_PROP_AUTOFOCUS, 0)
                self.cap.set(cv2.CAP_PROP_FOCUS, self.focus_var.get())
                self.status_var.set(
                    f"🔒 焦距巡邏：AF 被重置，已補回 {self.focus_var.get()}")
            else:
                # 無論如何都主動送一次，防止驅動靜默漂移
                self.cap.set(cv2.CAP_PROP_FOCUS, self.focus_var.get())

            # 2. 清晰度驟降偵測：Laplacian 方差跌破上次的 40% → 立即再補送
            frame = None
            with self._frame_lock:
                frame = self._latest_frame
            if frame is not None:
                gray  = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                score = cv2.Laplacian(gray, cv2.CV_64F).var()
                prev  = getattr(self, "_guardian_last_score", score)
                self._guardian_last_score = score
                if prev > 10 and score < prev * 0.4:
                    # 清晰度驟降，立即再補送一次焦距
                    self.cap.set(cv2.CAP_PROP_AUTOFOCUS, 0)
                    self.cap.set(cv2.CAP_PROP_FOCUS, self.focus_var.get())
                    self.status_var.set(
                        f"🔒 清晰度驟降({score:.0f}<{prev:.0f}×40%)，焦距已補回")

        else:
            # ── 自動對焦模式 ────────────────────────────────
            # 閒置鎖焦狀態下：持續補送鎖定的焦距，防止驅動偷改回 AF
            if getattr(self, "_idle_af_locked", False):
                actual_af = self.cap.get(cv2.CAP_PROP_AUTOFOCUS)
                if actual_af != 0:
                    self.cap.set(cv2.CAP_PROP_AUTOFOCUS, 0)
                self.cap.set(cv2.CAP_PROP_FOCUS, self._idle_locked_focus)

        self._focus_guardian_job = self.root.after(1500, self._focus_guardian)

    # ── 閒置自動鎖焦 ──────────────────────────────────────
    def _mark_activity(self):
        """每次成功掃到條碼時呼叫：
        1. 記錄當下焦距為「黃金焦距」（此時畫面清晰，值最可靠）
        2. 重置閒置計時
        3. 若目前處於閒置鎖焦，解鎖讓 AF 微調後再鎖回黃金焦距
        """
        # ── 學習黃金焦距 ──────────────────────────────────
        if self.cap and self.cap.isOpened() and self.autofocus_var.get():
            current_focus = int(self.cap.get(cv2.CAP_PROP_FOCUS))
            # 若 AF 目前已鎖定（閒置狀態），讀到的是鎖定值，不算新學習
            if not self._idle_af_locked and current_focus > 0:
                if self._golden_focus is None:
                    # 第一次學習：直接採用
                    self._golden_focus = current_focus
                else:
                    # 後續學習：做指數移動平均（EMA α=0.3），讓值逐漸穩定
                    # 避免單次異常值把黃金焦距帶歪
                    alpha = 0.3
                    self._golden_focus = int(
                        alpha * current_focus + (1 - alpha) * self._golden_focus)
                self._golden_focus_count += 1
                self.status_var.set(
                    f"📐 黃金焦距已更新 = {self._golden_focus}"
                    f"（第 {self._golden_focus_count} 次學習）")

        self._last_activity_ts = time.time()
        if self._idle_af_locked:
            self._idle_unlock_focus()

    def _idle_unlock_focus(self):
        """解除閒置鎖焦：若使用者原本開啟 AF，短暫恢復 AF 微調後再鎖回黃金焦距。"""
        if not self._idle_af_locked:
            return
        self._idle_af_locked = False
        if self.autofocus_var.get():
            if self.cap and self.cap.isOpened():
                self.cap.set(cv2.CAP_PROP_AUTOFOCUS, 1)
            # 1.2 秒後鎖回黃金焦距（給 AF 足夠時間微調）
            if self._idle_lock_job:
                self.root.after_cancel(self._idle_lock_job)
            self._idle_lock_job = self.root.after(1200, self._idle_lock_focus_now)

    def _idle_lock_focus_now(self):
        """鎖定焦距：優先使用黃金焦距，沒有黃金焦距才讀當下值。"""
        if not self.cap or not self.cap.isOpened():
            return
        if self.autofocus_var.get():
            if self._golden_focus is not None:
                # 有黃金焦距：直接鎖定（最可靠，掃碼時確認清晰的值）
                lock_val = self._golden_focus
                src = f"黃金焦距（學習 {self._golden_focus_count} 次）"
            else:
                # 尚未學習：退而求其次，讀當下 AF 值
                lock_val = int(self.cap.get(cv2.CAP_PROP_FOCUS))
                src = "當下 AF 值（尚未學習黃金焦距）"
            self._idle_locked_focus = lock_val
            self.cap.set(cv2.CAP_PROP_AUTOFOCUS, 0)
            self.cap.set(cv2.CAP_PROP_FOCUS, lock_val)
            self._idle_af_locked = True
            self.status_var.set(f"🔒 閒置鎖焦 = {lock_val}　來源：{src}")
        # 固定焦距模式不需額外動作（已由 guardian 維持）

    def _idle_check_and_lock(self):
        """每隔 10 秒檢查一次閒置狀態，超過 idle_lock_seconds 就鎖焦。"""
        if not self.running:
            return
        elapsed = time.time() - self._last_activity_ts
        if (not self._idle_af_locked
                and self.autofocus_var.get()
                and elapsed >= self._idle_lock_seconds):
            self._idle_lock_focus_now()
            hint = "掃碼時自動解鎖並微調" if self._golden_focus is not None \
                   else "請先掃一次條碼讓程式學習清晰焦距"
            self.status_var.set(
                f"🔒 閒置 {elapsed:.0f}s：已鎖焦防漂移（{hint}）")
        self._idle_lock_job = self.root.after(10_000, self._idle_check_and_lock)

    def _ae_loop(self):
        """偵測最新幀的平均亮度，若過暗/過亮則微調曝光，穩定 AF 基準。"""
        if not self.auto_exposure_var.get():
            # 自動曝光關閉：亮度不調，但焦距補送仍由 _focus_guardian 負責，直接結束
            return
        frame = None
        with self._frame_lock:
            frame = self._latest_frame
        brightness_changed = False
        if frame is not None and self.cap and self.cap.isOpened():
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            mean = float(gray.mean())
            if mean < 100:
                new_br = min(255, self.brightness_var.get() + 8)
                self.brightness_var.set(new_br)
                self.brightness_lbl.config(text=str(new_br))
                self.cap.set(cv2.CAP_PROP_BRIGHTNESS, new_br)
                brightness_changed = True
            elif mean > 160:
                new_br = max(0, self.brightness_var.get() - 8)
                self.brightness_var.set(new_br)
                self.brightness_lbl.config(text=str(new_br))
                self.cap.set(cv2.CAP_PROP_BRIGHTNESS, new_br)
                brightness_changed = True
        # DSHOW 驅動可能在任何時候偷改 autofocus，不只亮度調整時
        # 每輪都補送焦距指令，確保長時間運作下固定焦距不漂移
        self.root.after(200, self._apply_focus)
        self._ae_job = self.root.after(1500, self._ae_loop)

    def _on_video_click(self, event):
        """點擊預覽畫面，將放大中心移到該點。"""
        z = self.zoom_var.get()
        if z <= 1.0:
            return
        # 計算點擊位置在原始畫面上的比例
        px = event.x / DISPLAY_W
        py = event.y / DISPLAY_H
        half = 0.5 / z
        # 限制中心不超出邊界
        cx = max(half, min(1.0 - half, self._zoom_cx + (px - 0.5) / z))
        cy = max(half, min(1.0 - half, self._zoom_cy + (py - 0.5) / z))
        self._zoom_cx = cx
        self._zoom_cy = cy

    def _zoom_reset(self, event=None):
        """右鍵或快捷鍵 R：重置中心到畫面中央。"""
        self._zoom_cx = 0.5
        self._zoom_cy = 0.5

    # ── 掃描設定卡片 ──────────────────────────────────────
    def _build_scan_settings_card(self):
        card = self._make_card()
        tk.Label(card, text="⚙️  掃描節流設定",
                 font=F["card_title"], fg=C["text"],
                 bg=C["panel"]).pack(anchor="w", padx=12, pady=(10, 6))

        # 冷卻時間
        iv_row = tk.Frame(card, bg=C["panel"])
        iv_row.pack(fill="x", padx=12, pady=(0, 4))
        tk.Label(iv_row, text="同碼冷卻",
                 font=F["hint"], fg=C["muted"],
                 bg=C["panel"], width=6).pack(side="left")
        tk.Scale(iv_row, from_=0.5, to=10.0, resolution=0.5,
                 variable=self.scan_interval_var,
                 orient="horizontal", length=140,
                 bg=C["panel"], highlightthickness=0,
                 troughcolor=C["border"], sliderrelief="flat"
                 ).pack(side="left", padx=(4, 4))
        self.iv_lbl = tk.Label(iv_row, font=F["val_sm"],
                               fg=C["accent"], bg=C["panel"], width=5)
        self.iv_lbl.pack(side="left")
        self.scan_interval_var.trace_add("write", self._on_interval_change)
        self._on_interval_change()

        tk.Label(card, text="同一個碼多久後才再次接受（秒）",
                 font=F["hint"], fg=C["muted"],
                 bg=C["panel"], wraplength=280, justify="left"
                 ).pack(anchor="w", padx=12, pady=(0, 10))

    # ── 歷史紀錄卡片 ──────────────────────────────────────
    def _build_history_card(self):
        card = self._make_card()
        tk.Label(card, text="📋  本次統計",
                 font=F["card_title"], fg=C["text"],
                 bg=C["panel"]).pack(anchor="w", padx=12, pady=(10, 4))

        self.stat_lbl = tk.Label(
            card, text="掃描次數：0\n填入筆數：0",
            font=F["hint"], fg=C["muted"], bg=C["panel"],
            justify="left", anchor="w")
        self.stat_lbl.pack(fill="x", padx=12, pady=(0, 10))
        self._scan_count = 0
        self._fill_count = 0

    # ── 工具方法 ──────────────────────────────────────────
    def _make_card(self, pady_top=8):
        outer = tk.Frame(self.sf, bg=C["bg"])
        outer.pack(fill="x", pady=(pady_top, 0), padx=4)
        f = tk.Frame(outer, bg=C["panel"],
                     highlightbackground=C["border"],
                     highlightthickness=1)
        f.pack(fill="x")
        return f

    def _col_letter_to_num(self, col_str):
        col_str = col_str.strip().upper()
        num = 0
        for ch in col_str:
            if ch.isalpha():
                num = num * 26 + (ord(ch) - ord('A') + 1)
        return max(1, num)

    def _update_stat(self):
        self.stat_lbl.config(
            text=f"掃描次數：{self._scan_count}\n填入筆數：{self._fill_count}")

    def _ensure_inject_on(self):
        """啟動時若 inject_mode 預設 True，同步按鈕狀態並開始追蹤游標。"""
        if self.inject_mode_var.get() and WIN32_OK:
            self.inject_btn.config(text="● 填入中", bg=C["accent2"],
                                   activebackground=C["btn_green_hover"])
            self._start_cursor_tracking()
            if not self.scan_var.get():
                self._toggle_scan()

    def _on_delay_change(self, *_):
        v = self.follow_delay_var.get()
        self.delay_lbl.config(text=f"{v:.1f} 秒")

    # ── 掃碼邏輯 ──────────────────────────────────────────
    def _toggle_scan(self):
        if not self.scan_var.get():
            if not PYZBAR_OK:
                messagebox.showwarning("提示", "請先安裝 pyzbar 套件")
                return
            self.scan_var.set(True)
            self.scan_main_btn.config(text="⏹  停止掃描", bg="#dc2626",
                                      activebackground="#b91c1c")
            self._compact_scan_btn.config(text="⏹  停止掃描", bg="#dc2626",
                                          activebackground="#b91c1c")
            self.scan_state_lbl.config(text="● 掃描中", fg=C["accent2"])
            self.status_var.set("📷 掃碼模式已開啟")
        else:
            self.scan_var.set(False)
            self.scan_main_btn.config(text="▶  開始掃描", bg=C["purple"],
                                      activebackground=C["purple_hover"])
            self._compact_scan_btn.config(text="▶  開始掃描", bg=C["purple"],
                                          activebackground=C["purple_hover"])
            self.scan_state_lbl.config(text="● 關閉", fg=C["muted"])
            self.status_var.set("掃碼模式已關閉")

    # ── 影像預處理：多管線提升辨識率 ──────────────────────
    def _preprocess_candidates(self, frame):
        """
        回傳多個預處理版本的灰階影像，供 pyzbar 依序嘗試。
        越前面的管線越快，越後面處理越重。
        條碼強化模式（barcode_mode_var=True）會額外加入針對長窄條碼的管線。
        """
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        candidates = []

        # 1. 原始灰階（最快，適合清晰場景）
        candidates.append(gray)

        # 2. CLAHE：自適應直方圖均衡化，改善逆光 / 曝光不均
        clahe = cv2.createCLAHE(clipLimit=2.5, tileGridSize=(8, 8))
        candidates.append(clahe.apply(gray))

        # 3. 輕微銳化（unsharp mask），補償鏡頭柔焦
        blur  = cv2.GaussianBlur(gray, (0, 0), 3)
        sharp = cv2.addWeighted(gray, 1.6, blur, -0.6, 0)
        candidates.append(sharp)

        # 4. 自適應二值化（光線不均時效果最好）
        adp = cv2.adaptiveThreshold(
            gray, 255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY, 15, 4)
        candidates.append(adp)

        # 5. 放大 1.5×（小型 / 遠距條碼）
        h, w = gray.shape
        big = cv2.resize(gray, (int(w * 1.5), int(h * 1.5)),
                         interpolation=cv2.INTER_CUBIC)
        candidates.append(big)

        # ── 條碼強化模式：針對長窄線性條碼（Code128/Code39 等）──
        if getattr(self, "barcode_mode_var", None) and self.barcode_mode_var.get():

            # 6. 強銳化（強補焦距模糊）
            blur2  = cv2.GaussianBlur(gray, (0, 0), 2)
            sharp2 = cv2.addWeighted(gray, 2.2, blur2, -1.2, 0)
            candidates.append(sharp2)

            # 7. Otsu 二值化（全局最佳閾值）
            _, otsu = cv2.threshold(gray, 0, 255,
                                    cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            candidates.append(otsu)

            # 8. 強銳化後 Otsu（模糊條碼最有效）
            _, otsu_sharp = cv2.threshold(sharp2, 0, 255,
                                          cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            candidates.append(otsu_sharp)

            # 9. Gamma 提亮（γ=0.5）：暗部細節增強，適合燈光不足 / 標籤泛黃
            lut = np.array([int((i / 255.0) ** 0.5 * 255) for i in range(256)],
                           dtype=np.uint8)
            gamma_bright = cv2.LUT(gray, lut)
            candidates.append(gamma_bright)

            # 10. Gamma 提亮後 Otsu：低對比標籤（反光/泛黃）辨識率最高
            _, otsu_gamma = cv2.threshold(gamma_bright, 0, 255,
                                          cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            candidates.append(otsu_gamma)

            # 11. 水平形態學梯度：強調條碼垂直條紋邊緣（Code128 條紋方向）
            kernel_h = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 1))
            morph_grad = cv2.morphologyEx(gray, cv2.MORPH_GRADIENT, kernel_h)
            _, morph_bin = cv2.threshold(morph_grad, 0, 255,
                                         cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            candidates.append(morph_bin)

            # 12. CLAHE 強參數（clipLimit=4）：極端逆光 / 局部過曝
            clahe2 = cv2.createCLAHE(clipLimit=4.0, tileGridSize=(4, 4))
            clahe2_img = clahe2.apply(gray)
            candidates.append(clahe2_img)

            # 13. 放大 2× + 銳化（細條紋放大後補強邊緣，比直接放大灰階更清晰）
            big2 = cv2.resize(sharp2, (w * 2, h * 2),
                              interpolation=cv2.INTER_CUBIC)
            candidates.append(big2)

            # 14. 放大 2× + Otsu（最重但對極細條紋辨識率最高）
            _, otsu_big = cv2.threshold(big2, 0, 255,
                                        cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            candidates.append(otsu_big)

            # 15. 放大 2× + 自適應二值化（條碼局部光線不均時最有效）
            big2_gray = cv2.resize(gray, (w * 2, h * 2),
                                   interpolation=cv2.INTER_CUBIC)
            adp_big = cv2.adaptiveThreshold(
                big2_gray, 255,
                cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                cv2.THRESH_BINARY, 21, 5)
            candidates.append(adp_big)

        return candidates

    def _decode_best(self, frame, scale: float = 1.0):
        """
        依序嘗試多個預處理管線，找到就 early exit。
        scale：掃描幀相對原圖的縮放比（<1 表示已縮小），座標會自動換算回原圖。
        若有 filter 模式（非 all），只有通過 filter 的碼才算「找到」可 early exit。
        各管線若有放大，會在 img 的 .zoom_factor 屬性標記（用 numpy view 無法附加，
        改用 list of (img, zoom) tuple）。
        """
        seen = {}
        filter_mode = self.filter_mode_var.get()

        # 建立帶放大係數的候選管線清單
        # 格式：[(img, zoom)]  zoom=1.0 表示不需要座標還原
        raw_candidates = self._preprocess_candidates(frame)
        h0, w0 = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY).shape
        candidates_with_zoom = []
        for img in raw_candidates:
            ih, iw = img.shape[:2]
            zoom = iw / w0   # 計算相對原圖的放大比例
            candidates_with_zoom.append((img, zoom))

        old_fd, devnull = _suppress_zbar_stderr()
        try:
            for idx, (img, zoom) in enumerate(candidates_with_zoom):
                found = pyzbar.decode(img, symbols=SCAN_SYMBOLS)
                for c in found:
                    key = (c.type, c.data)
                    if key not in seen:
                        # 若此管線有放大，座標要縮回原圖尺寸
                        if zoom != 1.0:
                            c = c._replace(polygon=[
                                type(pt)(int(pt.x / zoom), int(pt.y / zoom))
                                for pt in c.polygon
                            ])
                        # 若掃描幀本身相對原圖有縮小，再放大座標
                        if scale != 1.0:
                            inv = 1.0 / scale
                            c = c._replace(polygon=[
                                type(pt)(int(pt.x * inv), int(pt.y * inv))
                                for pt in c.polygon
                            ])
                        seen[key] = c

                # early exit：前 5 管線（非放大版）找到就停
                if seen and idx < 5:
                    if filter_mode == "all":
                        break
                    else:
                        data_list = [c.data.decode("utf-8", errors="replace")
                                     for c in seen.values()]
                        if any(self._apply_filter(d) for d in data_list):
                            break
        finally:
            _restore_stderr(old_fd, devnull)
        # 多條碼時依面積排序（面積大 = 條紋較寬、較清晰，優先處理）
        codes = list(seen.values())
        def _area(c):
            pts = c.polygon
            if not pts:
                return 0
            xs = [p.x for p in pts]
            ys = [p.y for p in pts]
            return (max(xs) - min(xs)) * (max(ys) - min(ys))
        codes.sort(key=_area, reverse=True)
        return codes

    def _draw_codes(self, out, codes):
        """把條碼框線和標籤畫在 out 上（無 zoom 偏移）。"""
        for code in codes:
            pts = np.array(code.polygon, dtype=np.int32)
            data = code.data.decode("utf-8", errors="replace")
            kind = code.type
            x, y = pts[:, 0].min(), pts[:, 1].min()
            cv2.polylines(out, [pts], True, (0, 220, 80), 2)
            for pt in pts:
                cv2.circle(out, tuple(pt), 5, (0, 220, 80), -1)
            label = f"[{kind}] {data[:36]}{'…' if len(data) > 36 else ''}"
            (tw, th), _ = cv2.getTextSize(label, cv2.FONT_HERSHEY_SIMPLEX, 0.52, 2)
            cv2.rectangle(out, (x, y - th - 10), (x + tw + 6, y + 2), (0, 0, 0), -1)
            cv2.putText(out, label, (x + 3, y - 4),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.52, (0, 255, 100), 2)
        return out

    def _draw_codes_on_crop(self, crop_frame, codes, x1, y1, z):
        """把原圖座標的條碼框線投影到 zoom crop 後的顯示畫面。"""
        for code in codes:
            pts_orig = np.array(code.polygon, dtype=np.float32)
            # 減去 crop 偏移，再乘以顯示縮放比例（crop 後 resize 到 DISPLAY_W/H）
            crop_h, crop_w = crop_frame.shape[:2]
            dw = getattr(self, "_compact_dw", DISPLAY_W) if self._compact_mode else DISPLAY_W
            dh = getattr(self, "_compact_dh", DISPLAY_H) if self._compact_mode else DISPLAY_H
            sx = dw / crop_w
            sy = dh / crop_h
            pts = pts_orig.copy()
            pts[:, 0] = (pts_orig[:, 0] - x1)
            pts[:, 1] = (pts_orig[:, 1] - y1)
            pts = pts.astype(np.int32)
            data = code.data.decode("utf-8", errors="replace")
            kind = code.type
            x, y = pts[:, 0].min(), pts[:, 1].min()
            # 過濾掉不在 crop 範圍內的框
            if x < 0 or y < 0:
                continue
            cv2.polylines(crop_frame, [pts], True, (0, 220, 80), 2)
            for pt in pts:
                cv2.circle(crop_frame, tuple(pt), 5, (0, 220, 80), -1)
            label = f"[{kind}] {data[:36]}{'…' if len(data) > 36 else ''}"
            (tw, th), _ = cv2.getTextSize(label, cv2.FONT_HERSHEY_SIMPLEX, 0.52, 2)
            cv2.rectangle(crop_frame, (x, y - th - 10), (x + tw + 6, y + 2), (0, 0, 0), -1)
            cv2.putText(crop_frame, label, (x + 3, y - 4),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.52, (0, 255, 100), 2)
        return crop_frame

    def _process_codes(self, codes):
        """處理掃到的條碼：冷卻判斷、篩選、填入 Excel。"""
        import time
        now      = time.time()
        interval = self.scan_interval_var.get()
        maxcount = self.scan_maxcount_var.get()

        for code in codes:
            data = code.data.decode("utf-8", errors="replace")
            kind = code.type

            if not self._apply_filter(data):
                # 有掃到但不符合篩選條件，顯示提示而非完全沉默
                mode = self.filter_mode_var.get()
                if mode != "all":
                    self.root.after(0, lambda d=data, m=mode: self.status_var.set(
                        f"⚠ 掃到但不符合「{m.upper()}」格式：{d[:30]}"))
                continue

            cursor = self._cursor_info
            if cursor and cursor.get("app") == "excel":
                cell_tag = f"{cursor['sheet']}!{cursor['addr']}"
            elif cursor and cursor.get("app") == "word":
                cell_tag = f"word:{cursor.get('para', '?')}"
            else:
                cell_tag = "nocursor"

            # ── 冷卻 key：只看條碼本身（與儲存格無關）────────────────
            # 修正：原先 key 包含 cell_tag，導致下移後新格冷卻重算，
            # 或同一格連掃因 cell_tag 不變卻在冷卻期內無法觸發。
            cooldown_key = f"{kind}:{data}"
            # maxcount key：條碼+儲存格，確保「同碼填不同格」各自計次
            count_key    = f"{kind}:{data}@{cell_tag}"

            last_ts = self._code_ts.get(cooldown_key, 0.0)
            if (now - last_ts) < interval:
                remain = interval - (now - last_ts)
                self.root.after(0, lambda r=remain, d=data: self.status_var.set(
                    f"⏳ 冷卻中 {r:.1f}s：{d[:30]}"))
                continue

            count_so_far = self._code_count.get(count_key, 0)
            if maxcount > 0 and count_so_far >= maxcount:
                continue

            self._last_code    = cooldown_key
            self._last_code_ts = now
            self._code_ts[cooldown_key] = now
            self._code_count[count_key] = count_so_far + 1

            # 掃到條碼：重置閒置計時，解除閒置鎖焦（讓 AF 先對焦再鎖定）
            self.root.after(0, self._mark_activity)

            final_data = self._apply_extract(data)

            preview = f"原始：{data[:30]}\n填入：{final_data}"
            self.root.after(0, lambda p=preview:
                self.extract_preview_lbl.config(text=f"擷取預覽：\n{p}", fg=C["accent2"]))

            ts = datetime.now().strftime("%H:%M:%S")
            entry = {"time": ts, "type": kind, "data": final_data, "raw": data}
            self._scan_results.append(entry)
            self._scan_count += 1
            self.root.after(0, lambda e=entry: self._add_scan_result(e))
            if self.inject_mode_var.get():
                # 判斷是否啟用「全部掃交替填入」
                if (self.alt_fill_var.get()
                        and self.filter_mode_var.get() == "all"
                        and self._cursor_info
                        and self._cursor_info.get("app") == "excel"):
                    self.root.after(0, lambda d=final_data: self._do_inject_alt(d))
                else:
                    self.root.after(0, lambda d=final_data: self._do_inject(d))

    # 保留舊方法名稱以免其他地方有呼叫（已不使用）
    def _scan_and_draw(self, frame):
        codes = self._decode_best(frame)
        out = frame.copy()
        if codes:
            self._draw_codes(out, codes)
            self._process_codes(codes)
        return out

    def _add_scan_result(self, entry):
        txt = f"{entry['time']}  [{entry['type']}]\n{entry['data']}"
        self.scan_listbox.insert(0, txt)
        self.scan_listbox.itemconfig(0, fg=C["accent2"])
        if self.scan_listbox.size() > 100:
            self.scan_listbox.delete(100, tk.END)
        self.status_var.set(f"📷 掃到：{entry['data'][:50]}")
        self._update_stat()

    def _copy_selected(self):
        sel = self.scan_listbox.curselection()
        if not sel:
            messagebox.showinfo("提示", "請先在清單中選取一筆紀錄")
            return
        raw   = self.scan_listbox.get(sel[0])
        lines = raw.split("\n")
        data  = lines[1] if len(lines) > 1 else raw
        self.root.clipboard_clear()
        self.root.clipboard_append(data)
        self.status_var.set(f"✔ 已複製：{data[:50]}")

    def _clear_scan(self):
        if messagebox.askyesno("確認", "確定清除所有掃碼紀錄？"):
            self.scan_listbox.delete(0, tk.END)
            self._scan_results.clear()
            self._last_code = ""
            self._scan_count = 0
            self._update_stat()
            self.status_var.set("掃碼紀錄已清除")

    def _export_csv(self):
        if not self._scan_results:
            messagebox.showinfo("提示", "目前沒有掃碼紀錄")
            return
        path = filedialog.asksaveasfilename(
            title="匯出掃碼紀錄",
            defaultextension=".csv",
            filetypes=[("CSV 檔案", "*.csv"), ("所有檔案", "*.*")],
            initialfile=f"scan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8-sig") as f:
                f.write("時間,類型,資料\n")
                for r in self._scan_results:
                    data = r["data"].replace('"', '""')
                    f.write(f'{r["time"]},{r["type"]},"{data}"\n')
            self.status_var.set(f"✔ 已匯出 {len(self._scan_results)} 筆 → {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("匯出失敗", str(e))

    # ── 掃描設定回呼 ──────────────────────────────────────
    def _on_interval_change(self, *_):
        v = self.scan_interval_var.get()
        self.iv_lbl.config(text=f"{v:.1f} 秒")

    def _reset_code_count(self):
        self._code_count.clear()
        self._code_ts.clear()
        self._last_code    = ""
        self._last_code_ts = 0.0
        self.status_var.set("🔄 掃描計數已重置")

    # ── 診斷 Office COM ───────────────────────────────────
    def _show_diagnose(self):
        lines = _diagnose_office_com()
        msg   = "\n".join(lines)
        messagebox.showinfo("Office COM 診斷", msg)

    # ── 游標填入模式 ──────────────────────────────────────
    def _toggle_inject_mode(self):
        if not WIN32_OK:
            messagebox.showinfo("提示", "請先安裝 pywin32 套件")
            return
        on = not self.inject_mode_var.get()
        self.inject_mode_var.set(on)
        if on:
            self.inject_btn.config(text="● 填入中", bg=C["accent2"],
                                   activebackground=C["btn_green_hover"])
            self._start_cursor_tracking()
            if not self.scan_var.get():
                self._toggle_scan()
        else:
            self.inject_btn.config(text="○ 關閉", bg="#94a3b8",
                                   activebackground=C["border"])
            self._stop_cursor_tracking()

    def _start_cursor_tracking(self):
        self._stop_cursor_tracking()
        self._poll_cursor()

    def _stop_cursor_tracking(self):
        if self._cursor_track_job:
            self.root.after_cancel(self._cursor_track_job)
            self._cursor_track_job = None

    def _poll_cursor(self):
        """每 300ms 偵測 Office 游標（COM 呼叫在背景執行緒，不阻塞鏡頭）。"""
        if not self.inject_mode_var.get():
            return

        def _bg():
            info = _get_office_cursor()
            self.root.after(0, lambda: _ui(info))

        def _ui(info):
            self._cursor_info = info
            if info:
                if info["app"] == "excel":
                    self.cursor_lbl.config(
                        text=f"📊 Excel  {info['book']}\n"
                             f"   工作表：{info['sheet']}  儲存格：{info['addr']}",
                        fg=C["accent2"])
                elif info["app"] == "word":
                    self.cursor_lbl.config(
                        text=f"📝 Word  {info['doc']}\n"
                             f"   游標位置：{info['para']}",
                        fg=C["accent"])
            else:
                self.cursor_lbl.config(
                    text="⚠ 未偵測到 Excel/Word\n   請先開啟並點選儲存格",
                    fg=C["warn"])
            self._cursor_track_job = self.root.after(300, self._poll_cursor)

        threading.Thread(target=_bg, daemon=True).start()

    def _do_inject(self, data: str):
        """把掃到的 data 注入到目前游標位置（COM 呼叫移至背景執行緒，不阻塞鏡頭）。"""
        if not self.inject_mode_var.get():
            return
        info = self._cursor_info
        if not info:
            self.status_var.set("⚠ 找不到 Office 游標，請先點選儲存格")
            return

        # 把耗時的 COM 呼叫搬到背景執行緒，主執行緒不等待，cap.read() 節奏不受影響
        follow    = self.inject_follow_var.get()
        info_snap = dict(info)   # 快照，避免背景執行緒讀到已變更的 _cursor_info

        def _bg():
            ok, err = _inject_to_office(data, info_snap)
            # 結果回主執行緒更新 UI
            self.root.after(0, lambda: _ui(ok, err))

        def _ui(ok, err):
            if ok:
                if info_snap["app"] == "excel":
                    addr     = info_snap["addr"]
                    row, col = info_snap["row"], info_snap["col"]
                    next_row = row + 1
                    # 在 COM 寫入完成後才讀 delay_ms，確保延遲從填入完成時起算
                    delay_ms = int(self.follow_delay_var.get() * 1000)
                    self.cursor_lbl.config(
                        text=f"✔ 已填入 {info_snap['sheet']}!{addr}：{data[:20]}"
                             + (f"\n   {self.follow_delay_var.get():.1f}秒後下移至"
                                f" {_col_num_to_letter(col)}{next_row}" if follow else ""),
                        fg=C["accent2"])
                    self._fill_count += 1
                    self._update_stat()
                    self._inject_history.append({
                        "app":   "excel",
                        "sheet": info_snap["sheet"],
                        "row":   row,
                        "col":   col,
                    })
                    msg1 = f"✔ 已填入 Excel\n{info_snap['sheet']}!{addr}\n{data[:28]}"
                    self.root.after(0, lambda m=msg1: self._show_success_overlay(m))
                    if follow:
                        self.root.after(delay_ms, lambda r=row, c=col: self._delayed_move(r, c))
                elif info_snap["app"] == "word":
                    self.cursor_lbl.config(
                        text=f"✔ 已插入 Word：{data[:30]}",
                        fg=C["accent"])
                    self._fill_count += 1
                    self._update_stat()
                    self._inject_history.append({"app": "word", "data": data})
                    msg2 = f"✔ 已插入 Word\n{data[:28]}"
                    self.root.after(0, lambda m=msg2: self._show_success_overlay(m))
            else:
                self.status_var.set(f"⚠ 填入失敗：{err}")

        threading.Thread(target=_bg, daemon=True).start()

    def _do_inject_alt(self, data: str):
        """
        交替填入模式：
          step=0（等 MAC）→ 填游標欄，記下 base_row / base_col，step→1
          step=1（等 SN ）→ 填 base_col+1 同列，然後下移一列回 step=0
        若游標偏離了上次 base_row（使用者手動移格），自動重置到游標目前位置。
        """
        if not self.inject_mode_var.get():
            return
        # MAC 剛填完的短暫鎖定期內，忽略任何新掃描（防止 SN 搶在移格前觸發）
        if self._alt_fill_locked:
            self.status_var.set("⏳ 等待移格完成，暫緩填入...")
            return
        info = self._cursor_info
        if not info or info.get("app") != "excel":
            self.status_var.set("⚠ 找不到 Excel 游標，請先點選儲存格")
            return

        cur_row = info["row"]
        cur_col = info["col"]

        # ── 計算本次要填的目標欄列 ──────────────────────────
        if self._alt_fill_step == 0:
            # 第一筆（MAC）：以游標目前位置為基準
            self._alt_fill_base_row = cur_row
            self._alt_fill_base_col = cur_col
            target_row = cur_row
            target_col = cur_col
        else:
            # 第二筆（SN）：若使用者已手動移走游標，重置以游標為新基準
            if (self._alt_fill_base_row is None
                    or cur_row != self._alt_fill_base_row):
                # 游標已被移動，重置並從頭開始（本筆當作 MAC）
                self._alt_fill_base_row = cur_row
                self._alt_fill_base_col = cur_col
                self._alt_fill_step     = 0
                target_row = cur_row
                target_col = cur_col
            else:
                target_row = self._alt_fill_base_row
                target_col = self._alt_fill_base_col + 1   # SN 填右邊一欄

        step_snap     = self._alt_fill_step
        base_row_snap = self._alt_fill_base_row
        base_col_snap = self._alt_fill_base_col
        info_snap     = dict(info)
        info_snap["row"]  = target_row
        info_snap["col"]  = target_col
        info_snap["addr"] = f"{_col_num_to_letter(target_col)}{target_row}"

        def _bg():
            ok, err = _inject_to_office(data, info_snap)
            self.root.after(0, lambda: _ui(ok, err))

        def _ui(ok, err):
            if ok:
                step_label = "MAC" if step_snap == 0 else "SN"
                addr       = info_snap["addr"]
                self.cursor_lbl.config(
                    text=f"✔ 交替填入（{step_label}）{info_snap['sheet']}!{addr}：{data[:20]}",
                    fg=C["accent2"])
                self._fill_count += 1
                self._update_stat()
                self._inject_history.append({
                    "app":   "excel",
                    "sheet": info_snap["sheet"],
                    "row":   target_row,
                    "col":   target_col,
                })
                msg = f"✔ 交替填入（{step_label}）\n{info_snap['sheet']}!{addr}\n{data[:28]}"
                self._show_success_overlay(msg)

                if step_snap == 0:
                    # MAC 填完 → 鎖定後等 SN
                    self._alt_fill_step   = 1
                    self._alt_fill_locked = True
                    # 延遲解鎖（給鏡頭和游標追蹤一點緩衝，預設 0.8 秒）
                    unlock_ms = max(800, int(self.follow_delay_var.get() * 1000))
                    self.root.after(unlock_ms, self._alt_fill_unlock)
                else:
                    # SN 填完 → 下移一列，游標回 MAC 欄
                    self._alt_fill_step     = 0
                    self._alt_fill_locked   = False
                    self._alt_fill_base_row = None
                    self._alt_fill_base_col = None
                    delay_ms = int(self.follow_delay_var.get() * 1000)
                    # 移到 base_col（MAC 欄）的下一列
                    next_row = base_row_snap + 1
                    self.root.after(delay_ms,
                        lambda r=base_row_snap, c=base_col_snap: _move_excel_down(r, c))
                    # 同步更新游標快取
                    if self._cursor_info:
                        self._cursor_info = dict(self._cursor_info)
                        self._cursor_info["row"]  = next_row
                        self._cursor_info["col"]  = base_col_snap
                        self._cursor_info["addr"] = (
                            f"{_col_num_to_letter(base_col_snap)}{next_row}")
            else:
                self.status_var.set(f"⚠ 交替填入失敗：{err}")

        threading.Thread(target=_bg, daemon=True).start()

    def _alt_fill_unlock(self):
        """MAC 填完後的鎖定期結束，開放接受下一筆（SN）掃描。"""
        self._alt_fill_locked = False
        self.status_var.set("✅ 可掃 SN 了")

    def _show_success_overlay(self, msg: str):
        """右下角 Toast 通知：可拖動、釘選、記憶位置、開關控制。"""
        # 開關關閉時直接跳過
        if not self.toast_enabled_var.get():
            return

        # 取消前一個自動消失計時器（若已釘選則保留視窗，僅更新內容）
        if getattr(self, "_toast_job", None):
            try:
                self.root.after_cancel(self._toast_job)
            except Exception:
                pass
            self._toast_job = None

        # 若已有釘選中的 Toast，只更新文字內容，不重建視窗
        if self._toast_pinned and getattr(self, "_toast_win", None):
            try:
                lines = msg.replace("✔ ", "").split("\n")
                self._toast_title_var.set(lines[0] if lines else msg)
                self._toast_body_var.set("\n".join(lines[1:]) if len(lines) > 1 else "")
                # 更新後不設自動消失（釘選中）
                return
            except Exception:
                pass   # 視窗已損毀，重建

        # 關掉舊視窗
        if getattr(self, "_toast_win", None):
            try:
                self._toast_win.destroy()
            except Exception:
                pass
            self._toast_win = None

        lines = msg.replace("✔ ", "").split("\n")
        title = lines[0] if lines else msg
        body  = "\n".join(lines[1:]) if len(lines) > 1 else ""

        # ── 建立 Toast 視窗 ────────────────────────────────
        t = tk.Toplevel(self.root)
        t.overrideredirect(True)
        t.attributes("-topmost", True)
        t.attributes("-alpha", 0.95)
        t.configure(bg="#15803d")
        self._toast_win = t
        self._toast_title_var = tk.StringVar(value=title)
        self._toast_body_var  = tk.StringVar(value=body)

        # 內層深色底（左邊留 4px 給綠邊條）
        inner = tk.Frame(t, bg="#1a2e1a", padx=0, pady=0)
        inner.pack(fill="both", expand=True, padx=(4, 0))

        # 頂列：大勾 + 文字 + 釘選按鈕 + 關閉按鈕
        top_row = tk.Frame(inner, bg="#1a2e1a")
        top_row.pack(fill="x", padx=(10, 6), pady=(10, 2))

        tk.Label(top_row, text="✔",
                 font=("Segoe UI", 18, "bold"),
                 fg="#4ade80", bg="#1a2e1a").pack(side="left", padx=(0, 8))

        tk.Label(top_row, textvariable=self._toast_title_var,
                 font=("Segoe UI", 11, "bold"),
                 fg="#ffffff", bg="#1a2e1a",
                 anchor="w").pack(side="left", fill="x", expand=True)

        # 釘選按鈕
        pin_color = "#4ade80" if self._toast_pinned else "#64748b"
        self._toast_pin_btn = tk.Label(
            top_row, text="📌",
            font=("Segoe UI", 10), fg=pin_color, bg="#1a2e1a",
            cursor="hand2")
        self._toast_pin_btn.pack(side="left", padx=(4, 2))
        self._toast_pin_btn.bind("<Button-1>", self._toast_toggle_pin)

        # 關閉按鈕
        close_btn = tk.Label(
            top_row, text="✕",
            font=("Segoe UI", 10, "bold"), fg="#94a3b8", bg="#1a2e1a",
            cursor="hand2")
        close_btn.pack(side="left", padx=(2, 0))
        close_btn.bind("<Button-1>", lambda e: self._toast_close())

        # 內容文字
        self._toast_body_lbl = tk.Label(
            inner, textvariable=self._toast_body_var,
            font=("Segoe UI", 9), fg="#86efac", bg="#1a2e1a",
            anchor="w", justify="left")
        self._toast_body_lbl.pack(fill="x", padx=(36, 10), pady=(0, 10))

        # ── 拖動支援 ──────────────────────────────────────
        self._toast_drag_x = 0
        self._toast_drag_y = 0

        def _drag_start(e):
            self._toast_drag_x = e.x_root - t.winfo_x()
            self._toast_drag_y = e.y_root - t.winfo_y()

        def _drag_move(e):
            nx = e.x_root - self._toast_drag_x
            ny = e.y_root - self._toast_drag_y
            t.geometry(f"+{nx}+{ny}")
            # 即時記憶位置
            self._toast_x = nx
            self._toast_y = ny

        for widget in (inner, top_row, self._toast_body_lbl):
            widget.bind("<ButtonPress-1>",   _drag_start)
            widget.bind("<B1-Motion>",        _drag_move)

        # ── 定位：記憶位置 → 右下角預設 ──────────────────
        t.update_idletasks()
        sw = t.winfo_screenwidth()
        sh = t.winfo_screenheight()
        tw = max(t.winfo_reqwidth(), 290)
        th = t.winfo_reqheight()
        if self._toast_x is not None and self._toast_y is not None:
            tx, ty = self._toast_x, self._toast_y
        else:
            tx = sw - tw - 18
            ty = sh - th - 52
        t.geometry(f"{tw}x{th}+{tx}+{ty}")

        # ── 自動消失（未釘選才啟動）──────────────────────
        if not self._toast_pinned:
            def _fade(alpha=0.95):
                if not getattr(self, "_toast_win", None):
                    return
                if self._toast_pinned:   # 拖動後釘選就停止淡出
                    return
                alpha -= 0.08
                if alpha <= 0.0:
                    self._toast_close()
                    return
                try:
                    self._toast_win.attributes("-alpha", alpha)
                except Exception:
                    return
                self._toast_job = self.root.after(35, lambda: _fade(alpha))
            self._toast_job = self.root.after(2000, _fade)

    def _toast_toggle_pin(self, _=None):
        """切換 Toast 釘選狀態。"""
        self._toast_pinned = not self._toast_pinned
        # 取消淡出計時器
        if self._toast_pinned and getattr(self, "_toast_job", None):
            try:
                self.root.after_cancel(self._toast_job)
            except Exception:
                pass
            self._toast_job = None
            # 恢復不透明
            if getattr(self, "_toast_win", None):
                try:
                    self._toast_win.attributes("-alpha", 0.95)
                except Exception:
                    pass
        # 更新按鈕顏色
        if getattr(self, "_toast_pin_btn", None):
            color = "#4ade80" if self._toast_pinned else "#64748b"
            try:
                self._toast_pin_btn.config(fg=color)
            except Exception:
                pass

    def _toast_close(self):
        """手動關閉 Toast。"""
        if getattr(self, "_toast_job", None):
            try:
                self.root.after_cancel(self._toast_job)
            except Exception:
                pass
            self._toast_job = None
        if getattr(self, "_toast_win", None):
            try:
                self._toast_win.destroy()
            except Exception:
                pass
            self._toast_win = None

    def _delayed_move(self, row: int, col: int):
        """延遲後執行移格，並更新本地游標快取。"""
        _move_excel_down(row, col)
        if self._cursor_info and self._cursor_info.get("row") == row:
            self._cursor_info = dict(self._cursor_info)
            new_row = row + 1
            self._cursor_info["row"]  = new_row
            self._cursor_info["addr"] = f"{_col_num_to_letter(col)}{new_row}"

    def _inject_undo(self):
        """撤銷上一筆填入（Excel：清空儲存格；Word：不支援）。"""
        if not self._inject_history:
            messagebox.showinfo("提示", "沒有可撤銷的紀錄")
            return
        rec = self._inject_history.pop()
        if rec["app"] == "excel":
            for prog_id in ("Excel.Application", "Excel.Application.16",
                            "Excel.Application.15", "Excel.Application.14"):
                try:
                    xl = win32com.client.GetActiveObject(prog_id)
                    xl.Worksheets(rec["sheet"]).Cells(rec["row"], rec["col"]).Value = None
                    self.cursor_lbl.config(
                        text=f"↩ 已撤銷 {rec['sheet']}!{_col_num_to_letter(rec['col'])}{rec['row']}",
                        fg=C["warn"])
                    if self._fill_count > 0:
                        self._fill_count -= 1
                    self._update_stat()
                    break
                except Exception as e:
                    self.status_var.set(f"⚠ 撤銷失敗：{e}")
        elif rec["app"] == "word":
            messagebox.showinfo("提示", "Word 模式不支援撤銷，請在 Word 中手動按 Ctrl+Z")

    # ── 自動找最佳焦距 ────────────────────────────────────
    def _auto_find_focus(self):
        """掃描 0~255 所有焦距，用 Laplacian 方差找最清晰的數值，自動套用。"""
        if not self.cap or not self.cap.isOpened():
            self.status_var.set("⚠ 鏡頭未連線")
            return
        if getattr(self, "_focus_scan_running", False):
            return
        self._focus_scan_running = True
        self._autofocus_before_scan = self.autofocus_var.get()  # 記住掃描前的狀態
        self.autofocus_var.set(False)
        self.cap.set(cv2.CAP_PROP_AUTOFOCUS, 0)
        self.status_var.set("🔍 自動焦距掃描中，請將條碼放在鏡頭前...")

        def _scan():
            best_val   = self.focus_var.get()
            best_score = -1.0
            # 分粗掃（步進 10）+ 細掃（最佳值附近 ±15，步進 1）
            steps = list(range(0, 256, 10))

            for fv in steps:
                if not self._focus_scan_running:
                    break
                self.cap.set(cv2.CAP_PROP_FOCUS, fv)
                # 等待鏡頭實際移動（PW310P 約需 150ms）
                time.sleep(0.18)
                # 丟棄 2 幀舊緩衝
                for _ in range(2):
                    self.cap.read()
                ok, frame = self.cap.read()
                if not ok:
                    continue
                gray  = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                score = cv2.Laplacian(gray, cv2.CV_64F).var()
                self.root.after(0, lambda f=fv, s=score, b=best_score: self.status_var.set(
                    f"🔍 掃描焦距 {f}/255　清晰度={s:.0f}　目前最佳={b:.0f}"))
                if score > best_score:
                    best_score = score
                    best_val   = fv

            # 細掃：最佳值 ±15
            fine_start = max(0,   best_val - 15)
            fine_end   = min(255, best_val + 15)
            for fv in range(fine_start, fine_end + 1):
                if not self._focus_scan_running:
                    break
                self.cap.set(cv2.CAP_PROP_FOCUS, fv)
                time.sleep(0.12)
                for _ in range(2):
                    self.cap.read()
                ok, frame = self.cap.read()
                if not ok:
                    continue
                gray  = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                score = cv2.Laplacian(gray, cv2.CV_64F).var()
                if score > best_score:
                    best_score = score
                    best_val   = fv

            # 套用最佳值
            self.cap.set(cv2.CAP_PROP_FOCUS, best_val)
            self.root.after(0, lambda: self.focus_var.set(best_val))
            self.root.after(0, lambda: self.status_var.set(
                f"✔ 最佳焦距 = {best_val}　清晰度 = {best_score:.0f}　已套用並儲存"))
            # 還原掃描前的自動對焦狀態
            prev_af = getattr(self, "_autofocus_before_scan", True)
            self.root.after(0, lambda: self.autofocus_var.set(prev_af))
            self.root.after(0, self._apply_focus)
            self.root.after(0, self._save_settings)
            self._focus_scan_running = False

        threading.Thread(target=_scan, daemon=True).start()

    def _apply_focus(self):
        """依目前設定套用自動對焦或固定焦距。"""
        if not self.cap or not self.cap.isOpened():
            return
        if self.autofocus_var.get():
            self.cap.set(cv2.CAP_PROP_AUTOFOCUS, 1)
        else:
            self.cap.set(cv2.CAP_PROP_AUTOFOCUS, 0)
            self.cap.set(cv2.CAP_PROP_FOCUS, self.focus_var.get())

    def _apply_focus_with_retry(self, retries=3, delay_ms=800):
        """鏡頭剛開啟時分多次延遲送出 AF 指令，確保驅動程式已就緒。
        第一次也延遲 delay_ms 再送，避免 DSHOW 驅動尚未完成初始化就被覆蓋。"""
        def _try(n):
            if not self.cap or not self.cap.isOpened():
                return
            self._apply_focus()
            if n > 1:
                self.root.after(delay_ms, lambda: _try(n - 1))
        # 第一次也延遲，驅動初始化約需 500~1000ms
        self.root.after(delay_ms, lambda: _try(retries))

    def _on_focus_change(self, *_):
        """滑桿或 checkbox 改變時即時套用到鏡頭。"""
        self._apply_focus()
        af = self.autofocus_var.get()
        fv = self.focus_var.get()
        if af:
            if hasattr(self, "_focus_slider"):
                self._focus_slider.config(state="disabled")
            self.status_var.set("🔍 自動對焦 已開啟")
        else:
            if hasattr(self, "_focus_slider"):
                self._focus_slider.config(state="normal")
            self.status_var.set(f"🔍 固定焦距 = {fv}")

    # ── 鏡頭控制 ──────────────────────────────────────────
    def _start_camera(self):
        idx = self.camera_index.get()
        sel = self.res_var.get()
        res = next((r for r in self._res_options if r[0] == sel), self._res_options[0])
        cam_w, cam_h = res[1], res[2]
        # 低階 CPU 保護：鏡頭輸入解析度上限 1280×720，超過無意義且很耗資源
        cam_w = min(cam_w, 1280)
        cam_h = min(cam_h, 720)
        _com_thread_start()   # 確保常駐 COM 執行緒已啟動（只啟一次）
        self.cap = cv2.VideoCapture(idx, cv2.CAP_DSHOW)
        # MJPG 格式：減少 USB 頻寬佔用，讓 AF 運算更穩定
        self.cap.set(cv2.CAP_PROP_FOURCC, cv2.VideoWriter_fourcc(*'MJPG'))
        self.cap.set(cv2.CAP_PROP_FRAME_WIDTH,  cam_w)
        self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, cam_h)
        # 保持 30fps：PW310P 的 AF 需要足夠幀率做對比度計算，太低會 hunting
        self.cap.set(cv2.CAP_PROP_FPS, 30)
        if self.cap.isOpened():
            self.running = True
            actual_w = int(self.cap.get(cv2.CAP_PROP_FRAME_WIDTH))
            actual_h = int(self.cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
            # 延遲重試套用對焦：DSHOW 驅動初始化需要時間，立刻送指令會被忽略
            # 導致第二次開啟時焦距設定失效（模糊）。分 3 次在 0/800/1600ms 各送一次確保生效。
            self._apply_focus_with_retry(retries=3, delay_ms=800)
            self._apply_light()   # 套用光線補償設定
            if self.auto_exposure_var.get():
                self.root.after(1000, self._start_ae_loop)   # 鏡頭穩定後才啟動
            # 焦距巡邏：2.5 秒後啟動（等 retry 全部送完後再開始監控）
            self.root.after(2500, self._start_focus_guardian)
            # 閒置鎖焦：3 秒後開始第一次檢查（給鏡頭暖機時間）
            self._last_activity_ts = time.time()
            self._idle_af_locked   = False
            if self._idle_lock_job:
                self.root.after_cancel(self._idle_lock_job)
            self._idle_lock_job = self.root.after(3000, self._idle_check_and_lock)
            self.status_var.set(f"● 鏡頭 {idx} 已連線　{actual_w}×{actual_h}")
            # AF 預熱：1 秒後開始自動觸發驅動焦距學習，繞過 30 秒硬體限制
            self.root.after(1000, self._af_warmup_cycle)
            self._start_scan_thread()
            self._update_frame()
        else:
            self.root.withdraw()
            messagebox.showerror("鏡頭未連線", f"無法開啟鏡頭（編號 {idx}），請確認鏡頭已插上後再重新啟動程式。")
            self._on_close()

    def _start_scan_thread(self):
        """啟動背景掃描執行緒（僅負責 decode，不碰 UI）。"""
        self._scan_thread_running = True
        t = threading.Thread(target=self._scan_worker, daemon=True)
        t.start()
        # 設定執行緒為低優先權，避免搶走 UI 執行緒的 CPU
        try:
            import ctypes
            ctypes.windll.kernel32.SetThreadPriority(
                ctypes.windll.kernel32.GetCurrentThread(), -1)  # THREAD_PRIORITY_BELOW_NORMAL
        except Exception:
            pass

    def _scan_worker(self):
        """背景執行緒：持續從 _latest_frame 取最新幀來解碼。"""
        import time as _time
        # 掃描用解析度上限：1280px 寬，辨識率與速度的最佳平衡點
        SCAN_MAX_W = 1280

        while self._scan_thread_running:
            if not self.scan_var.get():
                with self._codes_lock:
                    self._latest_codes = []
                self._new_frame_event.wait(timeout=0.2)
                self._new_frame_event.clear()
                continue

            # 等新幀（最多 150ms）
            self._new_frame_event.wait(timeout=0.15)
            self._new_frame_event.clear()

            with self._frame_lock:
                frame = self._latest_frame

            if frame is None:
                continue

            # 縮小到掃描解析度（640px 寬對 pyzbar 已足夠辨識 QR）
            h, w = frame.shape[:2]
            if w > SCAN_MAX_W:
                scale = SCAN_MAX_W / w
                scan_frame = cv2.resize(
                    frame,
                    (SCAN_MAX_W, int(h * scale)),
                    interpolation=cv2.INTER_AREA)
            else:
                scale = 1.0
                scan_frame = frame

            codes = self._decode_best(scan_frame, scale if w > SCAN_MAX_W else 1.0)
            with self._codes_lock:
                self._latest_codes = codes

            if codes:
                self.root.after(0, lambda c=list(codes): self._process_codes(c))
            # 無論有無掃到，只稍微讓出 CPU，不做額外冷卻
            # （冷卻邏輯在 _process_codes 的 scan_interval 控制）
            _time.sleep(0.03)

    # ── AF 預熱：開機後自動觸發驅動焦距學習 ─────────────
    def _af_warmup_cycle(self, step=0):
        """
        鏡頭開啟後，連續做 2 次「關閉 AF → 等待 → 開啟 AF → 等待」的循環。
        這會強制 UVC/DSHOW 驅動在軟體層完成 AF calibration，
        繞過「需要手動掃碼觸發」的 30 秒硬體限制。
        每次循環間隔 3 秒，全程約 13 秒完成。
        預熱完成後彈出提示，請使用者對準條碼進行 5 秒學習。
        """
        if not self.cap or not self.cap.isOpened():
            return
        if not self.autofocus_var.get():
            return   # 使用者已關閉 AF，不干擾
        MAX_STEPS = 4   # 2 次循環 × 2（關/開）= 4 步
        if step >= MAX_STEPS:
            # 預熱完成：確保 AF 是開啟狀態，等最後一次開 AF 穩定後再跳提示
            self.cap.set(cv2.CAP_PROP_AUTOFOCUS, 1)
            self.status_var.set("🔄 AF 預熱中，即將完成…")
            self.root.after(1000, self._af_warmup_done)
            return
        if step % 2 == 0:
            # 偶數步：關閉 AF，固定住目前焦距值
            current = int(self.cap.get(cv2.CAP_PROP_FOCUS))
            self.cap.set(cv2.CAP_PROP_AUTOFOCUS, 0)
            self.cap.set(cv2.CAP_PROP_FOCUS, current if current > 0 else 50)
        else:
            # 奇數步：重新開啟 AF，讓驅動重新對焦
            self.cap.set(cv2.CAP_PROP_AUTOFOCUS, 1)
        cycle_num = step // 2 + 1
        self.status_var.set(f"🔄 AF 預熱中（第 {cycle_num}/2 次循環）…")
        self.root.after(3000, lambda: self._af_warmup_cycle(step + 1))

    def _af_warmup_done(self):
        """最後一次 AF 開啟穩定後，才顯示提示視窗。"""
        if not self.cap or not self.cap.isOpened():
            return
        self.status_var.set("✅ AF 預熱完成，請對準條碼開始學習")
        self._show_learning_prompt()

    def _show_learning_prompt(self):
        """預熱完成後彈出提示視窗，請使用者對準條碼，倒數 5 秒後提示學習完成。"""
        # 若已有舊視窗，先關掉
        self._dismiss_learning_prompt()

        t = tk.Toplevel(self.root)
        t.overrideredirect(True)
        t.attributes("-topmost", True)
        t.attributes("-alpha", 0.95)
        t.configure(bg="#1d4ed8")   # 藍色邊框
        self._learn_win = t

        inner = tk.Frame(t, bg="#1e293b", padx=0, pady=0)
        inner.pack(fill="both", expand=True, padx=(4, 0))

        # 標題列
        top_row = tk.Frame(inner, bg="#1e293b")
        top_row.pack(fill="x", padx=(10, 8), pady=(10, 2))

        tk.Label(top_row, text="📷",
                 font=(_UI, 14), fg="#60a5fa",
                 bg="#1e293b").pack(side="left", padx=(0, 8))
        tk.Label(top_row, text="焦距學習",
                 font=(_UI, 11, "bold"), fg="#ffffff",
                 bg="#1e293b").pack(side="left", fill="x", expand=True)

        close_lbl = tk.Label(top_row, text="✕",
                             font=(_UI, 10, "bold"), fg="#94a3b8",
                             bg="#1e293b", cursor="hand2")
        close_lbl.pack(side="left")
        close_lbl.bind("<Button-1>", lambda e: self._dismiss_learning_prompt())

        # 說明文字
        tk.Label(inner,
                 text="請將鏡頭對準條碼，\n系統將進行 8 秒焦距學習。",
                 font=(_UI, 9), fg="#bfdbfe", bg="#1e293b",
                 justify="left", anchor="w").pack(anchor="w", padx=14, pady=(4, 6))

        # 倒數文字
        self._learn_count_var = tk.StringVar(value="⏳ 剩餘 5 秒…")
        tk.Label(inner, textvariable=self._learn_count_var,
                 font=(_UI, 10, "bold"), fg="#fbbf24",
                 bg="#1e293b").pack(anchor="w", padx=14, pady=(0, 10))

        # 定位到主視窗右上角
        self.root.update_idletasks()
        rx = self.root.winfo_x()
        ry = self.root.winfo_y()
        rw = self.root.winfo_width()
        t.geometry(f"260x130+{rx + rw - 270}+{ry + 60}")

        # 開始倒數
        self._learning_countdown(8)

    def _learning_countdown(self, remaining):
        """倒數計時，結束後顯示學習完成訊息。"""
        if not hasattr(self, "_learn_win") or not self._learn_win:
            return
        if remaining <= 0:
            self._learn_count_var.set("✅ 學習完成！焦距已鎖定")
            self.status_var.set("✅ 焦距學習完成")
            # 3 秒後自動關閉視窗
            self._learn_close_job = self.root.after(3000, self._dismiss_learning_prompt)
            return
        self._learn_count_var.set(f"⏳ 剩餘 {remaining} 秒…")
        self._learn_countdown_job = self.root.after(
            1000, lambda: self._learning_countdown(remaining - 1))

    def _dismiss_learning_prompt(self):
        """關閉焦距學習提示視窗，取消所有相關排程。"""
        for attr in ("_learn_close_job", "_learn_countdown_job"):
            job = getattr(self, attr, None)
            if job:
                self.root.after_cancel(job)
                setattr(self, attr, None)
        if getattr(self, "_learn_win", None):
            try:
                self._learn_win.destroy()
            except Exception:
                pass
            self._learn_win = None

    def _restart_camera(self):
        self._scan_thread_running = False
        self.running = False
        self._stop_focus_guardian()
        if self.cap:
            self.cap.release()
        self.root.after(300, self._start_camera)

    def _toggle_camera(self):
        if self.running:
            self.running = False
            self.status_var.set("○ 預覽已暫停")
        else:
            self.running = True
            self._update_frame()

    def _update_frame(self):
        """主執行緒：只負責讀取鏡頭畫面 + 顯示，不做掃描運算。"""
        if not self.running:
            return

        # ── 視窗最小化時完全暫停，節省 CPU ──────────────────
        try:
            state = self.root.state()
        except Exception:
            state = ""
        if state == "iconic":
            if not self._minimized:
                self._minimized = True
            self.root.after(500, self._update_frame)   # 最小化時每 500ms 檢查一次
            return
        if self._minimized:
            self._minimized = False                    # 還原視窗，恢復正常速率

        ret, frame = self.cap.read()
        if not ret:
            messagebox.showerror("鏡頭連線中斷", "鏡頭突然斷線，程式即將關閉。")
            self._on_close()
            return
        if ret:
            # 把最新幀給掃描執行緒（掃描永遠拿最新幀，不受顯示 FPS 影響）
            with self._frame_lock:
                self._latest_frame = frame
            self._new_frame_event.set()  # 喚醒掃描執行緒

            # 取最新掃碼結果（不等待）
            with self._codes_lock:
                codes = list(self._latest_codes)

            # ── 顯示節流：只在該更新畫面時才做 resize/轉色/渲染 ──
            now = getattr(self, '_last_display_time', 0)
            import time as _time
            # 掃描中 15 FPS（配合鏡頭），未掃描 5 FPS
            display_interval = 0.067 if self.scan_var.get() else 0.2
            if _time.monotonic() - now >= display_interval:
                self._last_display_time = _time.monotonic()

                # 數位放大 crop（僅影響顯示）
                z = self.zoom_var.get()
                if z > 1.0:
                    h, w = frame.shape[:2]
                    crop_w = int(w / z)
                    crop_h = int(h / z)
                    cx = int(self._zoom_cx * w)
                    cy = int(self._zoom_cy * h)
                    x1 = max(0, cx - crop_w // 2)
                    y1 = max(0, cy - crop_h // 2)
                    x1 = min(x1, w - crop_w)
                    y1 = min(y1, h - crop_h)
                    display_frame = frame[y1:y1+crop_h, x1:x1+crop_w].copy()
                    if codes and self.scan_var.get():
                        display_frame = self._draw_codes_on_crop(
                            display_frame, codes, x1, y1, z)
                else:
                    display_frame = frame.copy()
                    if codes and self.scan_var.get():
                        display_frame = self._draw_codes(display_frame, codes)

                # 顯示
                if PIL_OK:
                    dw = getattr(self, "_compact_dw", DISPLAY_W) if self._compact_mode else DISPLAY_W
                    dh = getattr(self, "_compact_dh", DISPLAY_H) if self._compact_mode else DISPLAY_H
                    display = cv2.resize(display_frame, (dw, dh))

                    # 👁 銳化預覽（只影響畫面顯示，不影響掃碼）
                    if getattr(self, "preview_sharpen_var", None) and self.preview_sharpen_var.get():
                        blur_p  = cv2.GaussianBlur(display, (0, 0), 2)
                        display = cv2.addWeighted(display, 2.0, blur_p, -1.0, 0)

                    display = cv2.cvtColor(display, cv2.COLOR_BGR2RGB)
                    img = ImageTk.PhotoImage(Image.fromarray(display))
                    self.video_label.config(image=img)
                    self.video_label.image = img

        # 固定 30ms 讀鏡頭（掃描幀率不變），顯示節流在上面處理
        self.root.after(30, self._update_frame)

    # ── 釘選（Always on Top）────────────────────────────────
    def _toggle_pin(self):
        self._pinned = not self._pinned
        self.root.wm_attributes("-topmost", self._pinned)
        if self._pinned:
            self.pin_btn.config(text="📌", bg=C["accent2"],
                                activebackground=C["btn_green_hover"])
        else:
            self.pin_btn.config(text="📌", bg="#475569",
                                activebackground="#334155")

    # ── 簡易模式切換 ──────────────────────────────────────
    def _startup_compact_right(self):
        """啟動時自動切到簡易模式，並把視窗貼到螢幕右上角。"""
        self._startup_move_right = True   # 旗標：_apply_compact_geo 完成後自動靠右
        if not self._compact_mode:
            self._toggle_compact()
        else:
            self.root.after(50, self._move_to_right)

    def _move_to_right(self):
        """把視窗移到上次記住的簡易模式位置與大小；若無記錄則靠右上角。"""
        self.root.update_idletasks()
        ww = self.root.winfo_width()
        wh = self.root.winfo_height()
        if self._compact_pos and len(self._compact_pos) == 4:
            x, y, ww, wh = self._compact_pos
        elif self._compact_pos:
            x, y = self._compact_pos[0], self._compact_pos[1]
        else:
            sw = self.root.winfo_screenwidth()
            x  = sw - ww - 8
            y  = 8
        self.root.geometry(f"{ww}x{wh}+{x}+{y}")

    def _start_compact_pos_tracking(self):
        """簡易模式下每 500ms 記錄一次視窗位置與大小，拖移/縮放後自動儲存。"""
        if not self._compact_mode:
            return
        try:
            geo = self.root.geometry()          # "WxH+X+Y"
            parts = geo.replace("-", "+-").split("+")
            x = int(parts[1])
            y = int(parts[2])
            wh_parts = parts[0].split("x")
            w = int(wh_parts[0])
            h = int(wh_parts[1])
            new_pos = (x, y, w, h)
            if self._compact_pos != new_pos:
                self._compact_pos = new_pos
                self._save_compact_pos()
        except Exception:
            pass
        self._compact_track_job = self.root.after(500, self._start_compact_pos_tracking)

    def _save_compact_pos(self):
        """只把 compact_pos 更新寫入設定檔（背景執行緒，不阻塞鏡頭）。"""
        pos = list(self._compact_pos)
        def _write():
            try:
                data = {}
                if os.path.exists(SETTINGS_PATH):
                    with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                        data = json.load(f)
                data["compact_pos"] = pos
                with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
            except Exception:
                pass
        threading.Thread(target=_write, daemon=True).start()

    def _on_window_resize(self, event):
        """簡易模式下，視窗拉大時同步更新影像顯示尺寸。"""
        if not getattr(self, "_compact_mode", False):
            return
        if event.widget != self.root:
            return
        new_w = max(100, event.width - 6)
        new_h = max(80, event.height
                    - self._topbar.winfo_reqheight()
                    - self._status_bar.winfo_reqheight()
                    - 40)   # 留給按鈕高度
        self._compact_dw = new_w
        self._compact_dh = new_h
        self.video_label.config(width=new_w, height=new_h)

    def _toggle_compact(self):
        self._compact_mode = not self._compact_mode
        if self._compact_mode:
            self.root.update_idletasks()
            self._full_geometry = self.root.geometry()

            # 隱藏右側、頂列控制、整個底列、標題文字
            self._right_outer.pack_forget()
            self._topbar_ctrl.pack_forget()
            self._topbar_title.pack_forget()
            self._btn_bar.pack_forget()
            # 顯示重新啟動按鈕於 topbar 左側
            self._topbar_restart_btn.pack(side="left", padx=(8, 0), pady=10)

            # topbar 讓它自動縮（只剩完整按鈕）
            self._topbar.pack_propagate(True)

            # 關掉 body / left 的 expand
            self._body.pack_forget()
            self._body.pack(fill="none", expand=False, padx=0, pady=0)
            self._left.pack_forget()
            self._left.pack(side="left", fill="none", expand=False, padx=0)

            # 視訊縮小
            self._compact_dw = int(DISPLAY_W / 1.2)
            self._compact_dh = int(DISPLAY_H / 1.2)
            self.video_label.config(width=self._compact_dw,
                                    height=self._compact_dh)

            # 精簡掃描按鈕
            scan_text = "⏹  停止掃描" if self.scan_var.get() else "▶  開始掃描"
            scan_bg   = "#dc2626"     if self.scan_var.get() else C["purple"]
            self._compact_scan_btn.config(text=scan_text, bg=scan_bg,
                activebackground="#b91c1c" if self.scan_var.get() else C["purple_hover"])
            self._compact_btn_bar.pack(fill="x", side="bottom")

            self.compact_btn.config(text="⊞  完整", bg=C["accent2"],
                                    activebackground=C["btn_green_hover"])

            # 解除 minsize，讓視窗自由縮小
            self.root.minsize(0, 0)
            self.root.resizable(True, True)

            def _apply_compact_geo():
                self.root.update_idletasks()
                cw = self._compact_dw + 6
                ch = (self._topbar.winfo_reqheight()
                      + self._body.winfo_reqheight()
                      + 80)   # 掃描按鈕 + 重新啟動按鈕 + 間距
                # 若有記憶的大小就還原，否則用計算值
                if self._compact_pos and len(self._compact_pos) == 4:
                    cw, ch = self._compact_pos[2], self._compact_pos[3]
                self.root.geometry(f"{cw}x{ch}")
                # 保持可拉動，讓使用者自由調整大小
                if getattr(self, "_startup_move_right", False):
                    self._startup_move_right = False
                # 移到記住的位置（或預設右上角），然後開始追蹤拖移
                self.root.after(50, self._move_to_right)
                self.root.after(600, self._start_compact_pos_tracking)
            self.root.after(80, _apply_compact_geo)

        else:
            # 切回完整模式：停止簡易位置追蹤
            if hasattr(self, "_compact_track_job") and self._compact_track_job:
                self.root.after_cancel(self._compact_track_job)
                self._compact_track_job = None
            self.root.resizable(True, True)
            self.root.minsize(1020, 620)

            # 隱藏精簡按鈕列
            self._compact_btn_bar.pack_forget()

            # 恢復 topbar
            self._topbar.pack_propagate(False)
            self._topbar_restart_btn.pack_forget()
            self._topbar_title.pack(side="left", padx=16, pady=10)
            # 工具列恢復（在 topbar 下方、body 上方）
            self._topbar_ctrl.pack(fill="x", side="top", before=self._body)

            # 恢復 body / left expand
            self._body.pack_forget()
            self._body.pack(fill="both", expand=True, padx=10, pady=8)
            self._left.pack_forget()
            self._left.pack(side="left", fill="both", expand=True, padx=(0, 8))

            # 恢復右側、底列
            self._right_outer.pack(side="right", fill="y")
            self._btn_bar.pack(side="bottom", fill="x")
            self._pause_btn.pack(side="left", pady=8)

            # 視訊恢復
            self.video_label.config(width=DISPLAY_W, height=DISPLAY_H)
            self._compact_dw = DISPLAY_W
            self._compact_dh = DISPLAY_H

            self.compact_btn.config(text="⊟  簡易", bg="#475569",
                                    activebackground="#334155")

            def _apply_full_geo():
                geo = getattr(self, "_full_geometry",
                              f"{DISPLAY_W + 350}x{DISPLAY_H + 140}")
                self.root.geometry(geo)
            self.root.after(80, _apply_full_geo)

    def _restart_app(self):
        """原視窗內重置：停掉所有資源，清空 UI，重新執行 __init__ 流程。"""
        # 1. 停止所有背景工作
        self._scan_thread_running = False
        self.running = False
        self._stop_cursor_tracking()
        self._stop_ae_loop()
        self._stop_focus_guardian()
        if hasattr(self, "_compact_track_job") and self._compact_track_job:
            self.root.after_cancel(self._compact_track_job)
            self._compact_track_job = None
        if hasattr(self, "_idle_lock_job") and self._idle_lock_job:
            self.root.after_cancel(self._idle_lock_job)
            self._idle_lock_job = None
        self._dismiss_learning_prompt()
        _com_thread_stop()
        if self.cap:
            self.cap.release()
            self.cap = None

        # 2. 清空視窗內所有 widget
        for w in self.root.winfo_children():
            w.destroy()

        # 3. 重新初始化（等 200ms 讓資源完全釋放）
        self.root.after(200, lambda: QRExcelApp.__init__(self, self.root))

    def _on_close(self):
        self._scan_thread_running = False
        self.running = False
        self._stop_cursor_tracking()
        self._stop_ae_loop()
        self._stop_focus_guardian()
        if hasattr(self, "_compact_track_job") and self._compact_track_job:
            self.root.after_cancel(self._compact_track_job)
        self._dismiss_learning_prompt()
        if self.cap:
            self.cap.release()
        self._save_settings()
        # ── 停止常駐 COM 執行緒，讓 worker 自行 CoUninitialize（STA 規定同執行緒釋放）──
        _com_thread_stop()
        self.root.destroy()


# ══════════════════════════════════════════════════════════
if __name__ == "__main__":
    # 自動安裝 Pillow（預覽必要）
    if not PIL_OK:
        try:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", "Pillow"],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            from PIL import Image, ImageTk
            PIL_OK = True
        except Exception:
            pass

    root = tk.Tk()
    app  = QRExcelApp(root)
    root.mainloop()