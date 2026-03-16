import os
import tkinter as tk
from tkinter import ttk, filedialog
from plugin_manager import BasePlugin
import logger

MAX_LINES = 500


def _default_log_path():
    here = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(os.path.dirname(here), "telopon_debug.log")


class LogViewerPlugin(BasePlugin):
    PLUGIN_ID   = "log_viewer"
    PLUGIN_NAME = "デバッグログビューア"
    PLUGIN_TYPE = "TOOL"

    def __init__(self):
        super().__init__()
        self.plugin_queue = None
        self._last_pos    = 0
        self._after_id    = None
        self._auto_scroll = True

    def get_default_settings(self):
        return {
            "enabled":  False,
            "log_path": _default_log_path(),
        }

    def start(self, prompt_config, plugin_queue):
        self.plugin_queue = plugin_queue

    def stop(self):
        self.plugin_queue = None

    # ==========================================
    # UI
    # ==========================================
    def open_settings_ui(self, parent_window):
        if hasattr(self, "panel") and self.panel is not None and self.panel.winfo_exists():
            self.panel.lift()
            return

        self.panel = tk.Toplevel(parent_window)
        self.panel.title("デバッグログビューア")
        self.panel.geometry("860x540")
        self.panel.attributes("-topmost", True)

        # 有効/無効チェック（最上部）
        settings = self.get_settings()
        self._var_enabled = tk.BooleanVar(value=settings.get("enabled", False))
        ttk.Checkbutton(
            self.panel, text="デバッグログビューアを有効にする",
            variable=self._var_enabled
        ).pack(anchor="w", padx=6, pady=(6, 0))

        # ツールバー
        bar = tk.Frame(self.panel, bg="#2d2d2d")
        bar.pack(fill=tk.X)

        log_path  = settings.get("log_path", "") or _default_log_path()

        self._var_path = tk.StringVar(value=log_path)

        tk.Label(bar, text="📄", bg="#2d2d2d", fg="white").pack(side="left", padx=(6, 2), pady=4)
        ent_path = tk.Entry(bar, textvariable=self._var_path, width=60, bg="#3c3c3c", fg="#cccccc",
                            insertbackground="white", relief="flat")
        ent_path.pack(side="left", pady=4)

        tk.Button(bar, text="参照", bg="#555", fg="white", relief="flat",
                  command=self._browse_log).pack(side="left", padx=4)
        tk.Button(bar, text="再読込", bg="#555", fg="white", relief="flat",
                  command=self._reload).pack(side="left")

        self._var_auto = tk.BooleanVar(value=True)
        tk.Checkbutton(bar, text="自動スクロール", variable=self._var_auto,
                       bg="#2d2d2d", fg="white", selectcolor="#2d2d2d",
                       activebackground="#2d2d2d", activeforeground="white",
                       command=lambda: setattr(self, "_auto_scroll", self._var_auto.get())
                       ).pack(side="right", padx=6)
        tk.Button(bar, text="クリア", bg="#555", fg="white", relief="flat",
                  command=self._clear).pack(side="right", padx=4)

        # テキストエリア
        wrap_f = tk.Frame(self.panel)
        wrap_f.pack(fill=tk.BOTH, expand=True)

        self.txt = tk.Text(
            wrap_f,
            font=("Courier New", 9),
            bg="#1e1e1e", fg="#d4d4d4",
            state="disabled",
            wrap="none",
            cursor="arrow"
        )
        self.txt.tag_config("ERROR",   foreground="#f44747")
        self.txt.tag_config("WARNING", foreground="#dcdcaa")
        self.txt.tag_config("DEBUG",   foreground="#808080")
        self.txt.tag_config("INFO",    foreground="#9cdcfe")

        vsb = ttk.Scrollbar(wrap_f, orient="vertical",   command=self.txt.yview)
        hsb = ttk.Scrollbar(wrap_f, orient="horizontal", command=self.txt.xview)
        self.txt.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.txt.pack(fill=tk.BOTH, expand=True)

        # ステータスバー
        self.lbl_stat = tk.Label(self.panel, text="", fg="gray", bg="#2d2d2d", anchor="w")
        self.lbl_stat.pack(fill=tk.X, padx=6)

        self._last_pos = 0
        self._load_all()
        self._poll()

        self.panel.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        if self._after_id:
            try:
                self.panel.after_cancel(self._after_id)
            except Exception:
                pass
        self._after_id = None
        # パスと有効フラグを設定に保存
        s = self.get_settings()
        s["enabled"]  = self._var_enabled.get()
        s["log_path"] = self._var_path.get().strip()
        self.save_settings(s)
        self.panel.destroy()

    def _browse_log(self):
        path = filedialog.askopenfilename(
            title="ログファイルを選択",
            filetypes=[("Log files", "*.log *.txt"), ("All files", "*.*")]
        )
        if path:
            self._var_path.set(path)
            self._reload()

    def _reload(self):
        self._last_pos = 0
        self._load_all()

    # ==========================================
    # ポーリング（1秒ごと）
    # ==========================================
    def _poll(self):
        if not hasattr(self, "panel") or not self.panel.winfo_exists():
            return
        log_path = self._var_path.get().strip()
        try:
            size = os.path.getsize(log_path)
            if size < self._last_pos:
                self._load_all()
            elif size > self._last_pos:
                self._load_new(self._last_pos, log_path)
                self._last_pos = size
        except FileNotFoundError:
            self._set_status(f"ファイルが見つかりません: {log_path}")
        except Exception as e:
            self._set_status(f"読み込みエラー: {e}")

        self._after_id = self.panel.after(1000, self._poll)

    def _load_all(self):
        log_path = self._var_path.get().strip()
        try:
            with open(log_path, "r", encoding="utf-8", errors="replace") as f:
                lines = f.readlines()
            self._last_pos = os.path.getsize(log_path)
            self.txt.config(state="normal")
            self.txt.delete("1.0", tk.END)
            for line in lines[-MAX_LINES:]:
                self._insert_line(line)
            self.txt.config(state="disabled")
            if self._auto_scroll:
                self.txt.see(tk.END)
            self._set_status(f"{len(lines)} 行読み込み済み  |  {log_path}")
        except Exception as e:
            self._set_status(f"読み込みエラー: {e}")

    def _load_new(self, from_pos, log_path):
        try:
            with open(log_path, "r", encoding="utf-8", errors="replace") as f:
                f.seek(from_pos)
                lines = f.readlines()
            if not lines:
                return
            self.txt.config(state="normal")
            for line in lines:
                self._insert_line(line)
            end_line = int(self.txt.index("end-1c").split(".")[0])
            if end_line > MAX_LINES:
                self.txt.delete("1.0", f"{end_line - MAX_LINES}.0")
            self.txt.config(state="disabled")
            if self._auto_scroll:
                self.txt.see(tk.END)
            self._set_status(f"+{len(lines)} 行追加")
        except Exception as e:
            self._set_status(f"読み込みエラー: {e}")

    def _insert_line(self, line):
        upper = line.upper()
        if   "[ERROR]"   in upper: tag = "ERROR"
        elif "[WARN"     in upper: tag = "WARNING"
        elif "[DEBUG]"   in upper: tag = "DEBUG"
        else:                      tag = "INFO"
        self.txt.insert(tk.END, line, tag)

    def _clear(self):
        self.txt.config(state="normal")
        self.txt.delete("1.0", tk.END)
        self.txt.config(state="disabled")

    def _set_status(self, msg):
        if hasattr(self, "lbl_stat") and self.lbl_stat.winfo_exists():
            self.lbl_stat.config(text=f"  {msg}")
