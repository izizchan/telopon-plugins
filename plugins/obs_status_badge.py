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
        self.plugin_queue     = None
        self.is_running       = False
        self.is_connected     = True   # v1.22b: ライブ前からバッジをアクティブ表示
        self._ai_connected    = False
        self._thread          = None
        self._last_log_pos    = 0
        self._last_thinking   = 0.0
        self._obs_err_count   = 0   # 連続エラー数（ログ抑制用）
        self._font_size_fixed = False  # 初回フォントサイズ補正済みフラグ
        self._last_obs_state  = (None, None)  # 最後にOBSへ送信した (text, color)
        self._settings_open   = False          # 設定画面が開いている間はTrue

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
    def _check_source_exists(self, source_name):
        """OBSに接続してテキストソースが存在するか確認する。OBS未起動時はTrueを返す。"""
        conn = _load_obs_conn()
        try:
            import obsws_python as obs
            cl = obs.ReqClient(
                host=conn.get("host", "127.0.0.1"),
                port=int(conn.get("port", 4455)),
                password=conn.get("password", ""),
                timeout=3,
            )
            inputs = cl.get_input_list().inputs
            cl.disconnect()
            return any(inp.get("inputName") == source_name for inp in inputs)
        except Exception:
            return True  # OBS未起動などの場合は存在扱い（誤無効化を防ぐ）

    def _disable_badge(self):
        """badge_enabled を False に保存し、UI チェックも外す。"""
        s = self.get_settings()
        s["badge_enabled"] = False
        self.save_settings(s)
        try:
            if hasattr(self, "_var_badge_enabled"):
                self._var_badge_enabled.set(False)
        except Exception:
            pass

    def start(self, prompt_config, plugin_queue):
        if not self.get_settings().get("badge_enabled", True):
            return
        self.plugin_queue  = plugin_queue

        # テキストソースの存在確認
        source = self.get_settings().get("source_name", "").strip()
        if source and not self._check_source_exists(source):
            logger.warning(f"[{self.PLUGIN_NAME}] テキストソース '{source}' が見つかりません。プラグインを無効化します。")
            self._disable_badge()
            return

        self._ai_connected    = True
        self.is_running       = True
        self._last_log_pos    = self._get_log_size()
        self._last_thinking   = 0.0
        self._font_size_fixed = False
        self._last_obs_state  = (None, None)  # 再起動時は必ず初回送信を通す

        if self._thread is None or not self._thread.is_alive():
            self._thread = threading.Thread(target=self._loop, daemon=True)
            self._thread.start()

        logger.info(f"[{self.PLUGIN_NAME}] OBSステータス監視を開始しました。")

    def stop(self):
        self._ai_connected = False
        self.is_running    = False
        self.plugin_queue  = None
        # badge_enabled が有効な場合のみ切断ステータスを送信
        if self.get_settings().get("badge_enabled", True):
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

        self._settings_open = True

        self.panel = tk.Toplevel(parent_window)
        self.panel.title(self.PLUGIN_NAME)
        self.panel.geometry("480x340")
        self.panel.attributes("-topmost", True)
        self.panel.protocol("WM_DELETE_WINDOW", self._on_settings_close)

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
            foreground="gray", wraplength=430
        ).pack(anchor="w", pady=(0, 10))

        # テキストソース名
        row = ttk.Frame(main_f)
        row.pack(fill=tk.X, pady=3)
        ttk.Label(row, text="テキストソース名:", width=16).pack(side="left")
        self.ent_source = ttk.Entry(row, font=("", 11))
        self.ent_source.pack(side="left", fill=tk.X, expand=True)
        self.ent_source.insert(0, settings.get("source_name", ""))

        ttk.Separator(main_f, orient="horizontal").pack(fill=tk.X, pady=(10, 6))
        ttk.Label(
            main_f,
            text="各ステータスの表示文字列（テストボタンでOBSに即時送信）：",
            foreground="gray"
        ).pack(anchor="w", pady=(0, 4))

        # ステータス文字列入力行（テストボタン付き）
        def _status_row(label, fg_color, key, default, obs_color):
            r = ttk.Frame(main_f)
            r.pack(fill=tk.X, pady=2)
            ttk.Label(r, text=label, foreground=fg_color, width=8).pack(side="left")
            var = tk.StringVar(value=settings.get(key, default))
            ttk.Entry(r, textvariable=var, font=("", 10)).pack(side="left", fill=tk.X, expand=True, padx=(0, 4))

            def _test(v=var, d=default, c=obs_color):
                txt = v.get().strip() or d
                self._update_obs(txt, c, force=True)
                self._set_lbl_status(f"テスト送信: {txt}", "green")

            ttk.Button(r, text="テスト", width=6, command=_test).pack(side="left")
            return var

        self._var_text_connected    = _status_row("接続中 ●", "green",   "text_connected",    _ST_CONNECTED[0],    _COLOR_GREEN)
        self._var_text_thinking     = _status_row("思考中 ○", "#b8860b", "text_thinking",     _ST_THINKING[0],     _COLOR_YELLOW)
        self._var_text_disconnected = _status_row("切断   ●", "red",     "text_disconnected", _ST_DISCONNECTED[0], _COLOR_RED)

        ttk.Label(
            main_f,
            text="※ OBS側でテキストソース（GDI+）を使うと色が反映されます。",
            foreground="gray", justify="left"
        ).pack(anchor="w", pady=(8, 0))

        self.lbl_status = tk.Label(main_f, text="", foreground="gray",
                                   font=("", 9), anchor="w", padx=4, pady=1)
        self.lbl_status.pack(anchor="w", pady=(4, 0), fill=tk.X)

        tk.Button(
            main_f, text="保存して閉じる",
            bg="#007bff", fg="white",
            command=self._save_and_close
        ).pack(fill=tk.X, pady=(12, 0))

    _MSG_STYLE = {
        "gray":   {"foreground": "gray",    "background": ""},
        "green":  {"foreground": "#00aa44", "background": ""},
        "orange": {"foreground": "#FFD700", "background": "#3d2800"},
        "red":    {"foreground": "#FF8080", "background": "#3a0000"},
    }

    def _set_lbl_status(self, text, color="gray"):
        style = self._MSG_STYLE.get(color, self._MSG_STYLE["gray"])
        try:
            if self.lbl_status and self.lbl_status.winfo_exists():
                bg = style["background"] or self.lbl_status.master.cget("background")
                self.lbl_status.config(text=text, foreground=style["foreground"], background=bg)
        except Exception:
            pass

    def _on_settings_close(self):
        """設定画面をXボタンで閉じたとき。"""
        self._settings_open = False
        self._last_obs_state = (None, None)  # 次のループで状態を再送信させる
        try:
            self.panel.destroy()
        except Exception:
            pass

    def _save_and_close(self):
        s = self.get_settings()
        s["enabled"]            = True
        s["source_name"]        = self.ent_source.get().strip()
        s["text_connected"]     = self._var_text_connected.get().strip()    or _ST_CONNECTED[0]
        s["text_thinking"]      = self._var_text_thinking.get().strip()     or _ST_THINKING[0]
        s["text_disconnected"]  = self._var_text_disconnected.get().strip() or _ST_DISCONNECTED[0]

        # テキストソースが存在しない場合はチェックを無効化
        badge_enabled = self._var_badge_enabled.get()
        if badge_enabled and s["source_name"] and not self._check_source_exists(s["source_name"]):
            logger.warning(f"[{self.PLUGIN_NAME}] テキストソース '{s['source_name']}' が見つかりません。チェックを無効化します。")
            badge_enabled = False
            self._var_badge_enabled.set(False)
            self._set_lbl_status(f"⚠ ソース '{s['source_name']}' がOBSに見つかりません", "orange")
        s["badge_enabled"] = badge_enabled

        self.save_settings(s)

        # 設定画面フラグをリセット（ループ再開前に必ず解除）
        self._settings_open  = False
        self._last_obs_state = (None, None)  # 設定変更後は必ず次の更新を通す

        # badge_enabled が無効化された場合は実行中のループを停止する
        if not badge_enabled and self.is_running:
            self.is_running = False
            logger.info(f"[{self.PLUGIN_NAME}] badge_enabledが無効化されたため監視ループを停止しました。")

        # ライブ接続中かつ監視ループが止まっている場合は再開する
        if badge_enabled and self.plugin_queue and not self.is_running:
            self._ai_connected    = True
            self.is_running       = True
            self._last_log_pos    = self._get_log_size()
            self._last_thinking   = 0.0
            self._font_size_fixed = False
            if self._thread is None or not self._thread.is_alive():
                self._thread = threading.Thread(target=self._loop, daemon=True)
                self._thread.start()
            logger.info(f"[{self.PLUGIN_NAME}] 設定保存後にOBSステータス監視を再開しました。")

        self.panel.destroy()

    # ==========================================
    # バックグラウンドループ
    # ==========================================
    def _loop(self):
        while self.is_running:
            time.sleep(1.0)

            # 設定画面が開いている間は自動送信を停止（操作競合によるクラッシュ防止）
            if self._settings_open:
                continue

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
    def _update_obs(self, text, color, force=False):
        source = self.get_settings().get("source_name", "").strip()
        if not source:
            return

        # 変化がない場合はスキップ（force=True のテスト送信は除く）
        if not force and (text, color) == self._last_obs_state:
            return

        conn     = _load_obs_conn()
        host     = conn.get("host", "127.0.0.1")
        port     = int(conn.get("port", 4455))
        password = conn.get("password", "")

        try:
            import obsws_python as obs
            cl = obs.ReqClient(host=host, port=port, password=password, timeout=3)

            # 初回のみフォントサイズを確認し、200以上なら32に補正する
            if not self._font_size_fixed:
                try:
                    res = cl.get_input_settings(source)
                    font = res.input_settings.get("font", {})
                    size = font.get("size", 0)
                    if size >= 200:
                        font["size"] = 32
                        cl.set_input_settings(source, {"font": font}, True)
                        logger.info(f"[{self.PLUGIN_NAME}] フォントサイズを {size} → 25 に補正しました。")
                except Exception as fe:
                    logger.debug(f"[{self.PLUGIN_NAME}] フォントサイズ確認失敗: {fe}")
                self._font_size_fixed = True

            cl.set_input_settings(source, {"text": text, "color": color}, True)
            self._last_obs_state = (text, color)  # 送信成功時のみ記録
            self._obs_err_count  = 0
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
