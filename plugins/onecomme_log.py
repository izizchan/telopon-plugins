"""
わんコメ（OneComme）ログ監視プラグイン

監視対象: %APPDATA%\onecomme\comments\YYYY-MM-DD.log
各行は JSON: { "data": { "id", "displayName", "name", "comment", "timestamp", ... } }
処理: 新規IDのコメントのみ抽出 → バッチでAIに送信
      日付が変わったら自動で新しいログファイルに切り替え
"""

import os
import json
import time
import datetime
import threading
import tkinter as tk
from tkinter import ttk, filedialog
from plugin_manager import BasePlugin
import logger

_DEFAULT_LOG_DIR = os.path.join(os.environ.get("APPDATA", ""), "onecomme", "comments")


class OnecommeLogPlugin(BasePlugin):
    PLUGIN_ID   = "onecomme_log"
    PLUGIN_NAME = "わんコメログ読み込み"
    PLUGIN_TYPE = "BACKGROUND"

    def __init__(self):
        super().__init__()
        self.is_running   = False
        self.is_connected = self.get_settings().get("enabled", False)
        self.thread       = None
        # 保存済みパスがデフォルト値の絶対展開と一致すればリセット
        s = self.get_settings()
        if s.get("log_dir") == _DEFAULT_LOG_DIR:
            s["log_dir"] = ""
            self.save_settings(s)

    def get_default_settings(self):
        return {
            "enabled":      False,
            "log_dir":      "",
            "cooldown_sec": 5.0,
        }

    # ==========================================
    # 設定UI
    # ==========================================
    def open_settings_ui(self, parent_window):
        if hasattr(self, "settings_win") and self.settings_win is not None and self.settings_win.winfo_exists():
            self.settings_win.lift()
            return

        self.settings_win = tk.Toplevel(parent_window)
        self.settings_win.title(f"{self.PLUGIN_NAME} 設定")
        self.settings_win.geometry("520x240")
        self.settings_win.transient(parent_window)
        self.settings_win.grab_set()

        settings = self.get_settings()
        base = ttk.Frame(self.settings_win, padding=20)
        base.pack(fill="both", expand=True)

        # 有効チェック
        var_enabled = tk.BooleanVar(value=settings.get("enabled", False))
        ttk.Checkbutton(base, text="わんコメログ連携を有効にする", variable=var_enabled).pack(anchor="w", pady=(0, 12))

        # ログフォルダ
        ttk.Label(base, text="わんコメ コメントログフォルダ:").pack(anchor="w")
        row_dir = ttk.Frame(base)
        row_dir.pack(fill="x", pady=(4, 10))
        var_dir = tk.StringVar(value=settings.get("log_dir", "") or _DEFAULT_LOG_DIR)
        ent_dir = ttk.Entry(row_dir, textvariable=var_dir)
        ent_dir.pack(side="left", fill="x", expand=True)

        def _browse():
            path = filedialog.askdirectory(title="わんコメ コメントフォルダを選択")
            if path:
                var_dir.set(path)

        ttk.Button(row_dir, text="参照...", command=_browse).pack(side="right", padx=(5, 0))

        # クールダウン
        row_cd = ttk.Frame(base)
        row_cd.pack(fill="x")
        ttk.Label(row_cd, text="AIへの送信間隔 (秒):").pack(side="left")
        var_cd = tk.StringVar(value=str(settings.get("cooldown_sec", 5.0)))
        ttk.Entry(row_cd, textvariable=var_cd, width=6).pack(side="left", padx=5)
        ttk.Label(row_cd, text="※短いとAIがパニックになります。5秒推奨。", foreground="gray").pack(side="left")

        # ボタン
        btn_f = ttk.Frame(base)
        btn_f.pack(fill="x", pady=(18, 0))

        def _save():
            try:
                cd = float(var_cd.get())
            except ValueError:
                cd = 5.0
            self.save_settings({
                "enabled":      var_enabled.get(),
                "log_dir":      "" if var_dir.get().strip() == _DEFAULT_LOG_DIR else var_dir.get().strip(),
                "cooldown_sec": cd,
            })
            self.is_connected = var_enabled.get()
            self.settings_win.destroy()

        ttk.Button(btn_f, text="キャンセル", command=self.settings_win.destroy).pack(side="right")
        ttk.Button(btn_f, text="保存", command=_save).pack(side="right", padx=(0, 8))

    # ==========================================
    # ライフサイクル
    # ==========================================
    def start(self, prompt_config, plugin_queue):
        settings = self.get_settings()
        if not settings.get("enabled"):
            return
        if self.is_running:
            return

        self.is_running   = True
        self.is_connected = True
        self.thread = threading.Thread(
            target=self._watch_loop,
            args=(settings, prompt_config, plugin_queue),
            daemon=True
        )
        self.thread.start()
        logger.info(f"[{self.PLUGIN_NAME}] わんコメログ監視を開始しました。")

    def stop(self):
        self.is_running   = False
        self.is_connected = False
        logger.info(f"[{self.PLUGIN_NAME}] わんコメログ監視を停止しました。")

    # ==========================================
    # バックグラウンドループ
    # ==========================================
    def _watch_loop(self, settings, prompt_config, plugin_queue):
        log_dir     = settings.get("log_dir", "") or _DEFAULT_LOG_DIR
        cooldown    = float(settings.get("cooldown_sec", 5.0))

        processed_ids = set()       # 処理済みID（重複防止）
        pending       = []          # 送信待ちコメント
        last_send     = 0.0
        last_size     = -1

        today, src_path = self._today_path(log_dir)

        # 起動時: 既存行のIDを登録（過去ログはAIに送らない）
        if os.path.exists(src_path):
            for obj in self._iter_json_lines(src_path):
                if obj.get("data", {}).get("id"):
                    processed_ids.add(obj["data"]["id"])
            last_size = os.path.getsize(src_path)
            logger.debug(f"[{self.PLUGIN_NAME}] 起動時スキップ: {len(processed_ids)}件")

        while self.is_running:
            time.sleep(1.0)

            # 日付が変わったら新しいログファイルに切り替え
            new_today, new_path = self._today_path(log_dir)
            if new_today != today:
                today, src_path = new_today, new_path
                processed_ids.clear()
                last_size = -1
                logger.info(f"[{self.PLUGIN_NAME}] 日付切替: {src_path}")

            if not os.path.exists(src_path):
                continue

            try:
                cur_size = os.path.getsize(src_path)
            except OSError:
                continue

            if cur_size == last_size:
                pass  # 変化なし
            else:
                last_size = cur_size
                time.sleep(0.2)  # 書き込み競合回避

                for obj in self._iter_json_lines(src_path):
                    data = obj.get("data", {})
                    cid  = data.get("id")
                    if not cid:
                        continue
                    if cid in processed_ids:
                        continue
                    processed_ids.add(cid)

                    display = (data.get("displayName") or data.get("name") or "名無し").strip()
                    comment = (data.get("comment") or "").strip()
                    if not comment:
                        continue

                    if len(comment) > 100:
                        comment = comment[:100] + "…"

                    line = f"[COMMENT] {display}: {comment}"
                    pending.append(line)
                    logger.debug(f"[{self.PLUGIN_NAME}] 新規: {line}")

            # クールダウン経過後にまとめて送信
            now = time.time()
            if pending and (now - last_send) >= cooldown:
                cmt_msg  = prompt_config.get("CMT_MSG", "※新着コメントです。")
                ai_name  = prompt_config.get("ai_name", "AI")
                streamer = prompt_config.get("streamer_name", "")
                cmt_msg  = cmt_msg.replace("{ai_name}", ai_name).replace("{streamer}", streamer).strip()

                prefix = f"（{len(pending)}件）\n" if len(pending) > 1 else ""
                body   = "\n".join(pending)
                text   = f"{prefix}{body}\n\n{cmt_msg}" if cmt_msg else f"{prefix}{body}"

                self.send_text(plugin_queue, text)
                logger.info(f"[{self.PLUGIN_NAME}] AIに送信: {len(pending)}件")
                pending.clear()
                last_send = now

    # ==========================================
    # ヘルパー
    # ==========================================
    @staticmethod
    def _today_path(log_dir):
        today = datetime.date.today().strftime("%Y-%m-%d")
        return today, os.path.join(log_dir, f"{today}.log")

    @staticmethod
    def _iter_json_lines(path):
        """ファイル全行をJSONとして解析して yield する。パース失敗行はスキップ。"""
        try:
            with open(path, "r", encoding="utf-8", errors="replace") as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        yield json.loads(line)
                    except json.JSONDecodeError:
                        pass
        except OSError:
            pass
