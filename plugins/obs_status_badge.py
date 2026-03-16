import os
import json
import threading
import time
import tkinter as tk
from tkinter import ttk
from plugin_manager import BasePlugin
import logger

# OBS GDI+ テキストソースの色（ABGR整数: A=bits31-24, B=23-16, G=15-8, R=7-0）
_COLOR_GREEN  = 0xFF00FF00   # 緑  : R=0,   G=255, B=0
_COLOR_YELLOW = 0xFF00FFFF   # 黄  : R=255, G=255, B=0
_COLOR_RED    = 0xFF0000FF   # 赤  : R=255, G=0,   B=0

_ST_CONNECTED    = ("● 接続中",  _COLOR_GREEN)
_ST_THINKING     = ("○ 思考中",  _COLOR_YELLOW)
_ST_DISCONNECTED = ("● 切断",    _COLOR_RED)

# ログに出たら「思考中」とみなすキーワード（大文字・小文字無視）
_THINKING_KEYWORDS = [
    "生成中", "generating", "thinking",
    "ai送信", "キューに追加", "送信中", "processing",
    "llm", "gemini", "response start",
]
_THINKING_TIMEOUT = 6   # 秒以内にキーワードが出れば思考中

# ログに出たら「切断」とみなすキーワード
_DISCONNECT_KEYWORDS = [
    "データ受信中にAPIから切断されました",
    "deadline expired",
    "1011",
    "websocket connection closed",
    "connection closed",
    "disconnected",
]

# ログに出たら「再接続完了」とみなすキーワード（自動リトライ後の復帰検知）
_RECONNECT_KEYWORDS = [
    "connected to gemini live api",
    "システム再接続完了",
]


def _root():
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def _log_path():
    return os.path.join(_root(), "telopon_debug.log")

def _load_obs_conn():
    path = os.path.join(_root(), "plugins", "obs_capture.json")
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


class ObsStatusBadgePlugin(BasePlugin):
    PLUGIN_ID   = "obs_status_badge"
    PLUGIN_NAME = "OBS接続ステータス表示"
    PLUGIN_TYPE = "TOOL"

    def __init__(self):
        super().__init__()
        self.plugin_queue    = None
        self.is_running      = False
        self.is_connected    = True   # v1.22b: ライブ前からバッジをアクティブ表示
        self._ai_connected   = False
        self._thread         = None
        self._last_log_pos   = 0
        self._last_thinking  = 0.0
        self._obs_err_count  = 0   # 連続エラー数（ログ抑制用）

    def get_default_settings(self):
        return {
            "enabled": True, "badge_enabled": True, "source_name": "",
            "text_connected":    _ST_CONNECTED[0],
            "text_thinking":     _ST_THINKING[0],
            "text_disconnected": _ST_DISCONNECTED[0],
        }

    # ==========================================
    # ライフサイクル
    # ==========================================
    def start(self, prompt_config, plugin_queue):
        if not self.get_settings().get("badge_enabled", True):
            return
        self.plugin_queue  = plugin_queue
        self._ai_connected = True
        self.is_running    = True
        self._last_log_pos = self._get_log_size()
        self._last_thinking = 0.0

        if self._thread is None or not self._thread.is_alive():
            self._thread = threading.Thread(target=self._loop, daemon=True)
            self._thread.start()

        logger.info(f"[{self.PLUGIN_NAME}] OBSステータス監視を開始しました。")

    def stop(self):
        self._ai_connected = False
        self.is_running    = False
        self.plugin_queue  = None
        # 切断ステータスを即送信
        txt = self.get_settings().get("text_disconnected", _ST_DISCONNECTED[0])
        self._update_obs(txt, _ST_DISCONNECTED[1])
        logger.info(f"[{self.PLUGIN_NAME}] OBSステータス監視を停止しました。")

    # ==========================================
    # 設定UI
    # ==========================================
    def open_settings_ui(self, parent_window):
        if hasattr(self, "panel") and self.panel is not None and self.panel.winfo_exists():
            self.panel.lift()
            return

        self.panel = tk.Toplevel(parent_window)
        self.panel.title(self.PLUGIN_NAME)
        self.panel.geometry("430x340")
        self.panel.attributes("-topmost", True)

        settings = self.get_settings()
        main_f = ttk.Frame(self.panel, padding=15)
        main_f.pack(fill=tk.BOTH, expand=True)

        # 有効/無効チェック（最上部）
        self._var_badge_enabled = tk.BooleanVar(value=settings.get("badge_enabled", True))
        tk.Checkbutton(
            main_f, text="OBS接続ステータス表示を有効にする",
            variable=self._var_badge_enabled
        ).pack(anchor="w", pady=(0, 6))

        ttk.Label(
            main_f,
            text="OBS接続設定は「OBS画面AI実況」プラグインと共有されます。",
            foreground="gray", wraplength=390
        ).pack(anchor="w", pady=(0, 10))

        # テキストソース名
        row = ttk.Frame(main_f)
        row.pack(fill=tk.X, pady=3)
        ttk.Label(row, text="テキストソース名:", width=16).pack(side="left")
        self.ent_source = ttk.Entry(row, font=("", 11))
        self.ent_source.pack(side="left", fill=tk.X, expand=True)
        self.ent_source.insert(0, settings.get("source_name", ""))

        ttk.Separator(main_f, orient="horizontal").pack(fill=tk.X, pady=(10, 6))
        ttk.Label(main_f, text="各ステータスの表示文字列：", foreground="gray").pack(anchor="w", pady=(0, 4))

        # ステータス文字列入力行
        def _status_row(label, color, key, default):
            r = ttk.Frame(main_f)
            r.pack(fill=tk.X, pady=2)
            ttk.Label(r, text=label, foreground=color, width=8).pack(side="left")
            var = tk.StringVar(value=settings.get(key, default))
            ttk.Entry(r, textvariable=var, font=("", 10)).pack(side="left", fill=tk.X, expand=True)
            return var

        self._var_text_connected    = _status_row("接続中 ●", "green",  "text_connected",    _ST_CONNECTED[0])
        self._var_text_thinking     = _status_row("思考中 ○", "#b8860b", "text_thinking",     _ST_THINKING[0])
        self._var_text_disconnected = _status_row("切断   ●", "red",    "text_disconnected", _ST_DISCONNECTED[0])

        ttk.Label(
            main_f,
            text="※ OBS側でテキストソース（GDI+）を使うと色が反映されます。",
            foreground="gray", justify="left"
        ).pack(anchor="w", pady=(8, 0))

        self.lbl_status = ttk.Label(main_f, text="", foreground="gray")
        self.lbl_status.pack(anchor="w", pady=(4, 0))

        tk.Button(
            main_f, text="保存して閉じる",
            bg="#007bff", fg="white",
            command=self._save_and_close
        ).pack(fill=tk.X, pady=(12, 0))

    def _save_and_close(self):
        s = self.get_settings()
        s["enabled"]            = True
        s["badge_enabled"]      = self._var_badge_enabled.get()
        s["source_name"]        = self.ent_source.get().strip()
        s["text_connected"]     = self._var_text_connected.get().strip()    or _ST_CONNECTED[0]
        s["text_thinking"]      = self._var_text_thinking.get().strip()     or _ST_THINKING[0]
        s["text_disconnected"]  = self._var_text_disconnected.get().strip() or _ST_DISCONNECTED[0]
        self.save_settings(s)
        self.panel.destroy()

    # ==========================================
    # バックグラウンドループ
    # ==========================================
    def _loop(self):
        while self.is_running:
            time.sleep(1.0)

            new_lines = self._read_new_log_lines()
            lowered   = " ".join(new_lines).lower()

            s = self.get_settings()
            txt_connected    = s.get("text_connected",    _ST_CONNECTED[0])
            txt_thinking     = s.get("text_thinking",     _ST_THINKING[0])
            txt_disconnected = s.get("text_disconnected", _ST_DISCONNECTED[0])

            # 切断エラーをログから検知（stop() が呼ばれなくても対応）
            if self._ai_connected:
                if any(kw.lower() in lowered for kw in _DISCONNECT_KEYWORDS):
                    self._ai_connected = False
                    self._update_obs(txt_disconnected, _ST_DISCONNECTED[1])
                    logger.info(f"[{self.PLUGIN_NAME}] ログで切断を検知 → ステータスを切断に更新")
                    continue

            # 切断中：再接続ログを検知したら復帰、それ以外は待機
            if not self._ai_connected:
                if any(kw.lower() in lowered for kw in _RECONNECT_KEYWORDS):
                    self._ai_connected = True
                    self._last_thinking = 0.0
                    logger.info(f"[{self.PLUGIN_NAME}] ログで再接続を検知 → ステータスを接続中に更新")
                else:
                    continue

            # 思考中キーワードを検出
            if any(kw in lowered for kw in _THINKING_KEYWORDS):
                self._last_thinking = time.time()

            if time.time() - self._last_thinking < _THINKING_TIMEOUT:
                self._update_obs(txt_thinking, _ST_THINKING[1])
            else:
                self._update_obs(txt_connected, _ST_CONNECTED[1])

    # ==========================================
    # OBS更新
    # ==========================================
    def _update_obs(self, text, color):
        source = self.get_settings().get("source_name", "").strip()
        if not source:
            return

        conn     = _load_obs_conn()
        host     = conn.get("host", "127.0.0.1")
        port     = int(conn.get("port", 4455))
        password = conn.get("password", "")

        try:
            import obsws_python as obs
            cl = obs.ReqClient(host=host, port=port, password=password, timeout=3)
            cl.set_input_settings(source, {"text": text, "color": color}, True)
            self._obs_err_count = 0  # 成功したらリセット
        except Exception as e:
            self._obs_err_count += 1
            # 初回と10回ごとだけログ出力（ログ爆発を防ぐ）
            if self._obs_err_count == 1 or self._obs_err_count % 10 == 0:
                logger.debug(f"[{self.PLUGIN_NAME}] OBS更新失敗 ({self._obs_err_count}回): {e}")

    # ==========================================
    # ログ読み取り
    # ==========================================
    def _get_log_size(self):
        try:
            return os.path.getsize(_log_path())
        except Exception:
            return 0

    def _read_new_log_lines(self):
        try:
            lp   = _log_path()
            size = os.path.getsize(lp)
            if size <= self._last_log_pos:
                return []
            with open(lp, "r", encoding="utf-8", errors="replace") as f:
                f.seek(self._last_log_pos)
                lines = f.readlines()
            self._last_log_pos = size
            return lines
        except Exception:
            return []
