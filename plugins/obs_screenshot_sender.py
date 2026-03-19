import os
import json
import base64
import io
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox

import logger

try:
    import obsws_python as obs
    from PIL import Image, ImageTk
except ImportError:
    messagebox.showerror(
        "ライブラリ不足",
        "コマンドプロンプトで以下を実行してください。\n\npip install obsws-python pillow"
    )

from plugin_manager import BasePlugin

_DEFAULT_PROMPT = "【配信画面の更新】今の画面の状況です！これを見て実況やツッコミを入れてください！"
_PICKER_EMPTY = "-- 未選択(空) --"


def _load_obs_conn():
    here = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(here, "obs_capture.json")
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


class ObsScreenshotSenderPlugin(BasePlugin):
    PLUGIN_ID   = "obs_screenshot_sender"
    PLUGIN_NAME = "OBS画面AI送信"
    PLUGIN_TYPE = "TOOL"

    def __init__(self):
        super().__init__()
        self.plugin_queue    = None
        self.is_running      = False
        self.is_connected    = True   # v1.22b: ライブ前からバッジをアクティブ表示
        self.ent_sources     = []
        self.txt_prompts     = []
        self.ent_scenes      = []
        self.lbl_previews    = []
        self.lbl_slot_msgs   = []
        self.preview_photos  = [None, None, None, None]
        self.lbl_status      = None
        self._auto_thread    = None
        self._auto_start_time = 0.0
        self._last_hashes        = [None, None, None, None]
        self._event_client       = None
        self._current_obs_scene  = None
        self._capture_buttons    = []

    def get_default_settings(self):
        return {
            "plugin_enabled": True,
            "source1": "", "prompt1": _DEFAULT_PROMPT, "scene1": "",
            "source2": "", "prompt2": _DEFAULT_PROMPT, "scene2": "",
            "source3": "", "prompt3": _DEFAULT_PROMPT, "scene3": "",
            "source4": "", "prompt4": _DEFAULT_PROMPT, "scene4": "",
            "auto_send1": False, "auto_send2": False,
            "auto_send3": False, "auto_send4": False,
            "interval_sec1": 120, "interval_sec2": 120,
            "interval_sec3": 120, "interval_sec4": 120,
            "scene_send1": False, "scene_send2": False,
            "scene_send3": False, "scene_send4": False,
            "auto_stop_min": 180,
            "skip_duplicate": True,
        }

    # ==========================================
    # ライフサイクル
    # ==========================================
    def start(self, prompt_config, plugin_queue):
        self.plugin_queue = plugin_queue
        self.is_running   = True
        self._last_hashes = [None, None, None, None]
        logger.debug(f"[{self.PLUGIN_NAME}] キューを接続しました。")

        s = self.get_settings()
        if any(s.get(f"auto_send{i}", False) for i in range(1, 5)):
            self._start_auto_loop()

        self._start_event_listener()

    def stop(self):
        self.is_running   = False
        self.plugin_queue = None
        self._stop_event_listener()
        logger.debug(f"[{self.PLUGIN_NAME}] キューを切断しました。")

    # ==========================================
    # 定期送信ループ（スロットごとに独立タイマー）
    # ==========================================
    def _start_auto_loop(self):
        self._auto_start_time = time.time()
        if self._auto_thread is None or not self._auto_thread.is_alive():
            self._auto_thread = threading.Thread(target=self._auto_loop, daemon=True)
            self._auto_thread.start()
            logger.info(f"[{self.PLUGIN_NAME}] 定期送信ループを開始しました。")

    def _auto_loop(self):
        last_sends = [time.time() for _ in range(4)]

        while self.is_running:
            s = self.get_settings()

            # プラグイン無効時はスキップ
            if not s.get("plugin_enabled", True):
                time.sleep(1.0)
                continue

            now = time.time()

            # 自動停止チェック（いずれかのスロットが定期ON中の場合のみ）
            stop_min = int(s.get("auto_stop_min", 0))
            any_active = any(s.get(f"auto_send{i}", False) for i in range(1, 5))
            if stop_min > 0 and any_active:
                if (now - self._auto_start_time) >= stop_min * 60:
                    logger.info(f"[{self.PLUGIN_NAME}] {stop_min}分経過のため全スロットの定期送信・シーン限定送信を自動停止しました。")
                    for i in range(1, 5):
                        s[f"auto_send{i}"]  = False
                        s[f"scene_send{i}"] = False
                    self.save_settings(s)
                    self._update_auto_send_ui_all(False)
                    # 自動停止後もループは継続（手動で再ONできるよう）
                    time.sleep(1.0)
                    continue

            for idx in range(4):
                if not self.is_running:
                    break
                if not s.get(f"auto_send{idx + 1}", False):
                    continue

                interval = max(10, int(s.get(f"interval_sec{idx + 1}", 120)))
                if now - last_sends[idx] < interval:
                    continue

                # シーン名が設定されている場合は一致しないとスキップ
                active_scene = s.get(f"scene{idx + 1}", "").strip()
                if active_scene and self._current_obs_scene != active_scene:
                    continue

                source = s.get(f"source{idx + 1}", "").strip()
                if source:
                    self._capture_and_send(source, idx + 1, idx,
                                           skip_dup=s.get("skip_duplicate", True))
                last_sends[idx] = time.time()

            time.sleep(1.0)

    def _update_auto_send_ui_all(self, value: bool):
        """全スロットの定期送信・シーン限定送信チェックをUIに反映する。"""
        try:
            if hasattr(self, "_var_auto_sends"):
                for var in self._var_auto_sends:
                    var.set(value)
            if hasattr(self, "_var_scene_sends"):
                for var in self._var_scene_sends:
                    var.set(value)
        except Exception:
            pass

    # ==========================================
    # OBSイベント待ち受け（シーン切り替え・WSコマンド）
    # ==========================================
    def _start_event_listener(self):
        try:
            conn = _load_obs_conn()
            host     = conn.get("host", "127.0.0.1")
            port     = int(conn.get("port", 4455))
            password = conn.get("password", "")
            # 現在シーンを初期取得
            cl = obs.ReqClient(host=host, port=port, password=password, timeout=3)
            self._current_obs_scene = cl.get_current_program_scene().current_program_scene_name
            cl.disconnect()
            self._event_client = obs.EventClient(host=host, port=port, password=password)
            self._event_client.callback.register(self._on_scene_changed)
            self._event_client.callback.register(self._on_custom_event)
            logger.info(f"[{self.PLUGIN_NAME}] OBSイベント待ち受け開始（現在: {self._current_obs_scene}）")
        except Exception as e:
            logger.debug(f"[{self.PLUGIN_NAME}] OBSイベント接続失敗（OBS未起動？）: {e}")
            self._event_client = None

    def _stop_event_listener(self):
        if self._event_client:
            try:
                self._event_client.disconnect()
            except Exception:
                pass
            self._event_client = None
            logger.debug(f"[{self.PLUGIN_NAME}] OBSイベント待ち受けを停止しました。")

    def _on_scene_changed(self, data):
        try:
            self._current_obs_scene = data.scene_name
            logger.debug(f"[{self.PLUGIN_NAME}] シーン変更: {self._current_obs_scene}")
            try:
                if hasattr(self, "panel") and self.panel and self.panel.winfo_exists():
                    self.panel.after(0, self._refresh_capture_buttons)
            except Exception:
                pass
        except AttributeError:
            pass

    def _on_custom_event(self, data):
        """OBS BroadcastCustomEvent で送られた AI-SS-Sender コマンドを処理する。"""
        try:
            ev = data.event_data
            if not isinstance(ev, dict):
                return
            if ev.get("command") != "AI-SS-Sender":
                return
            action = ev.get("action", "")
            logger.debug(f"[{self.PLUGIN_NAME}] WSコマンド受信: action={action} / {ev}")
            self._dispatch_ws_command(action, ev)
        except AttributeError:
            pass
        except Exception as e:
            logger.debug(f"[{self.PLUGIN_NAME}] WSコマンド処理エラー: {e}")

    def _dispatch_ws_command(self, action, ev):
        """
        受け付けるコマンド一覧（OBS側から BroadcastCustomEvent で送信）:
          send        : {"command":"AI-SS-Sender","action":"send","slot":1}
                         → ソース1～4を即時キャプチャしてAIに送る
          set_source  : {"command":"AI-SS-Sender","action":"set_source","slot":1,"name":"ソース名"}
                         → スロット1～4のソース名を変更して設定を保存する
          set_interval: {"command":"AI-SS-Sender","action":"set_interval","slot":1,"seconds":60}
                         → 指定スロットの定期送信間隔（秒）を変更する（slot省略時は全スロット）
          auto        : {"command":"AI-SS-Sender","action":"auto","slot":1,"enabled":true}
                         → 指定スロットの定期送信ON/OFFを切り替える（slot省略時は全スロット）
          plugin      : {"command":"AI-SS-Sender","action":"plugin","enabled":true}
                         → プラグイン全体の有効/無効を切り替える
        """
        if action == "send":
            slot = int(ev.get("slot", 1))
            if 1 <= slot <= 4:
                s = self.get_settings()
                source = s.get(f"source{slot}", "").strip()
                if source:
                    threading.Thread(
                        target=self._capture_and_send,
                        args=(source, slot, slot - 1),
                        kwargs={"skip_dup": False},
                        daemon=True
                    ).start()
                    logger.info(f"[{self.PLUGIN_NAME}] WSコマンド send: ソース{slot}（{source}）を送信します。")
                else:
                    logger.warning(f"[{self.PLUGIN_NAME}] WSコマンド send: ソース{slot}の名前が未設定です。")
            else:
                logger.warning(f"[{self.PLUGIN_NAME}] WSコマンド send: slot={slot} は範囲外です（1～4）。")

        elif action == "set_source":
            slot = int(ev.get("slot", 1))
            name = str(ev.get("name", "")).strip()
            if 1 <= slot <= 4:
                s = self.get_settings()
                s[f"source{slot}"] = name
                self.save_settings(s)
                try:
                    if self.ent_sources and len(self.ent_sources) >= slot:
                        ent = self.ent_sources[slot - 1]
                        if ent.winfo_exists():
                            ent.delete(0, tk.END)
                            ent.insert(0, name)
                except Exception:
                    pass
                logger.info(f"[{self.PLUGIN_NAME}] WSコマンド set_source: ソース{slot}を「{name}」に変更しました。")
            else:
                logger.warning(f"[{self.PLUGIN_NAME}] WSコマンド set_source: slot={slot} は範囲外です（1～4）。")

        elif action == "set_interval":
            seconds = max(10, int(ev.get("seconds", 120)))
            slot = ev.get("slot", None)
            s = self.get_settings()
            targets = [int(slot)] if slot and 1 <= int(slot) <= 4 else range(1, 5)
            for t in targets:
                s[f"interval_sec{t}"] = seconds
            self.save_settings(s)
            try:
                if hasattr(self, "_var_intervals"):
                    for t in targets:
                        self._var_intervals[t - 1].set(str(seconds))
            except Exception:
                pass
            logger.info(f"[{self.PLUGIN_NAME}] WSコマンド set_interval: スロット{list(targets)}の間隔を{seconds}秒に変更しました。")

        elif action == "auto":
            enabled = bool(ev.get("enabled", False))
            slot = ev.get("slot", None)
            s = self.get_settings()
            targets = [int(slot)] if slot and 1 <= int(slot) <= 4 else range(1, 5)
            for t in targets:
                s[f"auto_send{t}"] = enabled
            self.save_settings(s)
            if enabled and self.is_running:
                self._start_auto_loop()
            try:
                if hasattr(self, "_var_auto_sends"):
                    for t in targets:
                        self._var_auto_sends[t - 1].set(enabled)
            except Exception:
                pass
            logger.info(f"[{self.PLUGIN_NAME}] WSコマンド auto: スロット{list(targets)}の定期送信を{'ON' if enabled else 'OFF'}にしました。")

        elif action == "plugin":
            enabled = bool(ev.get("enabled", True))
            s = self.get_settings()
            s["plugin_enabled"] = enabled
            self.save_settings(s)
            try:
                if hasattr(self, "_var_plugin_enabled"):
                    self._var_plugin_enabled.set(enabled)
            except Exception:
                pass
            logger.info(f"[{self.PLUGIN_NAME}] WSコマンド plugin: プラグインを{'有効' if enabled else '無効'}にしました。")

        else:
            logger.warning(f"[{self.PLUGIN_NAME}] WSコマンド: 不明なaction「{action}」を受信しました。")

    # ==========================================
    # OBS ピッカー（ドロップダウン）
    # ==========================================
    def _make_obs_client(self):
        conn = _load_obs_conn()
        host     = conn.get("host", "127.0.0.1")
        port     = int(conn.get("port", 4455))
        password = conn.get("password", "")
        return obs.ReqClient(host=host, port=port, password=password, timeout=3)

    def _close_active_picker(self):
        """開いているピッカーポップアップを閉じ、パネルの topmost を復元する。"""
        popup = getattr(self, "_active_picker_popup", None)
        if popup:
            try:
                if popup.winfo_exists():
                    popup.destroy()
            except Exception:
                pass
            self._active_picker_popup = None
        try:
            if self.panel.winfo_exists():
                self.panel.attributes("-topmost", True)
        except Exception:
            pass

    def _show_picker_popup(self, anchor_widget, entry_widget):
        """anchorウィジェットの直下にリスト選択ポップアップを表示する。
        パネルの topmost を一時解除してポップアップを最前面に出す。
        """
        self.panel.attributes("-topmost", False)

        popup = tk.Toplevel(self.panel)
        popup.overrideredirect(True)
        popup.attributes("-topmost", True)
        self._active_picker_popup = popup

        x = anchor_widget.winfo_rootx()
        y = anchor_widget.winfo_rooty() + anchor_widget.winfo_height()
        popup.geometry(f"+{x}+{y}")

        frame = tk.Frame(popup, bd=1, relief="solid", bg="#cccccc")
        frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(frame, orient=tk.VERTICAL)
        listbox = tk.Listbox(
            frame, yscrollcommand=scrollbar.set,
            selectmode=tk.SINGLE, font=("", 10),
            height=6, width=36,
            activestyle="dotbox"
        )
        scrollbar.config(command=listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        confirmed = [entry_widget.get()]  # Escape で戻るベースライン（リストで可変参照）

        def _get_selected_value():
            sel = listbox.curselection()
            if not sel:
                return None
            v = listbox.get(sel[0])
            return "" if v == _PICKER_EMPTY else v

        def on_apply(event=None):
            """現在の選択をエントリに即時反映する（プレビュー）。confirmed は更新しない。"""
            v = _get_selected_value()
            if v is not None:
                entry_widget.delete(0, tk.END)
                entry_widget.insert(0, v)

        def on_confirm(event=None):
            """マウスクリック・Space: 反映してベースラインも更新。ポップアップは閉じない。"""
            v = _get_selected_value()
            if v is not None:
                entry_widget.delete(0, tk.END)
                entry_widget.insert(0, v)
                confirmed[0] = v
            return "break"

        def on_confirm_close(event=None):
            """Enter: 確定してポップアップを閉じる。"""
            on_confirm()
            self._close_active_picker()

        def on_escape(event=None):
            """Escape: ベースラインに戻してポップアップを閉じる。"""
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, confirmed[0])
            self._close_active_picker()

        def on_any_click(event):
            """grab_set により全クリックがここに届く。ポップアップ外なら閉じる。"""
            rx, ry = event.x_root, event.y_root
            px, py = popup.winfo_rootx(), popup.winfo_rooty()
            pw, ph = popup.winfo_width(), popup.winfo_height()
            if not (px <= rx <= px + pw and py <= ry <= py + ph):
                self._close_active_picker()

        listbox.bind("<<ListboxSelect>>", on_apply)       # マウス選択・矢印キー両方で発火
        listbox.bind("<KeyRelease-Up>", on_apply)          # <<ListboxSelect>> が発火しない環境の保険
        listbox.bind("<KeyRelease-Down>", on_apply)
        listbox.bind("<ButtonRelease-1>", on_confirm)      # マウスクリック確定 → baseline 更新
        listbox.bind("<space>", on_confirm)                # Space 確定 → baseline 更新
        listbox.bind("<Return>", on_confirm_close)
        popup.bind("<Escape>", on_escape)
        popup.bind("<Button-1>", on_any_click)
        popup.grab_set()
        listbox.focus_set()

        return listbox, popup

    def _open_scene_picker(self, anchor_btn, entry_widget, source_idx):
        """OBSからシーン一覧を取得してピッカーを表示する。
        既にポップアップが開いていれば閉じてリターン（トグル）。
        """
        if getattr(self, "_active_picker_popup", None) and self._active_picker_popup.winfo_exists():
            self._close_active_picker()
            return

        listbox, popup = self._show_picker_popup(anchor_btn, entry_widget)
        listbox.insert(tk.END, "読込中...")

        def fetch():
            try:
                cl = self._make_obs_client()
                res = cl.get_scene_list()
                cl.disconnect()
                scenes = res.scenes
                return [s.scene_name if hasattr(s, "scene_name") else s.get("sceneName", "")
                        for s in reversed(scenes)]
            except Exception as e:
                return None, str(e)

        def on_done(result):
            if not popup.winfo_exists():
                return
            listbox.delete(0, tk.END)
            if isinstance(result, tuple):  # エラー
                _, msg = result
                self._slot_msg(source_idx, f"❌ OBS接続エラー: {msg}", "red")
                self._close_active_picker()
                return
            listbox.insert(tk.END, _PICKER_EMPTY)
            for name in result:
                listbox.insert(tk.END, name)
            listbox.config(height=min(12, max(3, len(result) + 1)))

        def run():
            result = fetch()
            if anchor_btn.winfo_exists():
                anchor_btn.after(0, lambda: on_done(result))

        threading.Thread(target=run, daemon=True).start()

    def _open_source_picker(self, anchor_btn, entry_widget, scene_entry_widget, source_idx):
        """OBSからソース一覧を取得してピッカーを表示する。
        scene_entry_widget が空ならアクティブシーン、入力があればそのシーンのアイテムを返す。
        既にポップアップが開いていれば閉じてリターン（トグル）。
        """
        if getattr(self, "_active_picker_popup", None) and self._active_picker_popup.winfo_exists():
            self._close_active_picker()
            return

        listbox, popup = self._show_picker_popup(anchor_btn, entry_widget)
        listbox.insert(tk.END, "読込中...")

        def fetch():
            try:
                cl = self._make_obs_client()
                scene_name = scene_entry_widget.get().strip() if scene_entry_widget else ""
                if not scene_name:
                    scene_name = cl.get_current_program_scene().current_program_scene_name
                res = cl.get_scene_item_list(scene_name)
                cl.disconnect()
                items = res.scene_items
                names = []
                for item in items:
                    if hasattr(item, "source_name"):
                        names.append(item.source_name)
                    elif isinstance(item, dict):
                        names.append(item.get("sourceName", ""))
                return list(dict.fromkeys(reversed(names)))  # 重複除去・順序保持
            except Exception as e:
                return None, str(e)

        def on_done(result):
            if not popup.winfo_exists():
                return
            listbox.delete(0, tk.END)
            if isinstance(result, tuple):  # エラー
                _, msg = result
                self._slot_msg(source_idx, f"❌ OBS接続エラー: {msg}", "red")
                self._close_active_picker()
                return
            listbox.insert(tk.END, _PICKER_EMPTY)
            for name in result:
                listbox.insert(tk.END, name)
            listbox.config(height=min(12, max(3, len(result) + 1)))

        def run():
            result = fetch()
            if anchor_btn.winfo_exists():
                anchor_btn.after(0, lambda: on_done(result))

        threading.Thread(target=run, daemon=True).start()

    # ==========================================
    # UI
    # ==========================================
    def open_settings_ui(self, parent_window):
        if hasattr(self, "panel") and self.panel is not None and self.panel.winfo_exists():
            self.panel.lift()
            return

        self.panel = tk.Toplevel(parent_window)
        self.panel.title(self.PLUGIN_NAME)
        self.panel.geometry("810x850")
        self.panel.resizable(True, True)
        self.panel.attributes("-topmost", True)

        settings = self.get_settings()

        # ── ヘッダー ──
        header = ttk.Frame(self.panel, padding=(12, 8, 12, 4))
        header.pack(fill=tk.X)

        self._var_plugin_enabled = tk.BooleanVar(value=settings.get("plugin_enabled", True))
        tk.Checkbutton(
            header, text="プラグイン有効",
            variable=self._var_plugin_enabled,
            font=("", 10, "bold")
        ).pack(side="left")

        ttk.Label(
            header,
            text="※OBS接続設定は「OBS画面AI実況」と共有",
            foreground="gray"
        ).pack(side="right")

        ttk.Separator(self.panel, orient="horizontal").pack(fill=tk.X, padx=8)

        # ── グローバル設定（自動停止・重複スキップ） ──
        opt_f = ttk.Frame(self.panel, padding=(12, 6, 12, 4))
        opt_f.pack(fill=tk.X)

        self._var_auto_stop = tk.BooleanVar(value=settings.get("auto_stop_min", 0) > 0)
        tk.Checkbutton(
            opt_f, text="自動停止:",
            variable=self._var_auto_stop
        ).grid(row=0, column=0, sticky="w")
        self._var_stop_min = tk.StringVar(value=str(max(1, settings.get("auto_stop_min", 180))))
        ttk.Spinbox(opt_f, from_=1, to=600, width=5,
                    textvariable=self._var_stop_min).grid(row=0, column=1, sticky="w")
        ttk.Label(opt_f, text="分後に定期送信を全停止").grid(row=0, column=2, padx=(2, 16), sticky="w")

        self._var_skip_dup = tk.BooleanVar(value=settings.get("skip_duplicate", True))
        tk.Checkbutton(
            opt_f, text="前回と同じ画像は送信しない",
            variable=self._var_skip_dup
        ).grid(row=0, column=3, sticky="w")

        # ── ソースカード（2列グリッド） ──
        grid_f = ttk.Frame(self.panel, padding=(8, 0, 8, 4))
        grid_f.pack(fill=tk.BOTH, expand=True)
        grid_f.columnconfigure(0, weight=1, uniform="col")
        grid_f.columnconfigure(1, weight=1, uniform="col")

        self.ent_sources      = []
        self.txt_prompts      = []
        self.ent_scenes       = []
        self.lbl_previews     = []
        self.lbl_slot_msgs    = []
        self.preview_photos   = [None, None, None, None]
        self._var_auto_sends  = []
        self._var_intervals   = []
        self._var_scene_sends = []
        self._capture_buttons = []

        positions = [(0, 0), (0, 1), (1, 0), (1, 1)]

        for i in range(1, 5):
            row_g, col_g = positions[i - 1]
            lf = ttk.LabelFrame(grid_f, text=f" ソース {i} ", padding=8)
            lf.grid(row=row_g, column=col_g, sticky="nsew", padx=5, pady=5)
            lf.columnconfigure(4, weight=1)  # ソース名入力欄が伸びる
            grid_f.rowconfigure(row_g, weight=1)
            idx = i - 1

            # ソース名行: [定期☐] [間隔:] [spinbox] [秒] [OBSソース名:] [entry]
            var_as = tk.BooleanVar(value=settings.get(f"auto_send{i}", False))
            self._var_auto_sends.append(var_as)
            tk.Checkbutton(lf, text="定期送信", variable=var_as).grid(
                row=0, column=0, sticky="w", padx=(0, 2))

            var_iv = tk.StringVar(value=str(settings.get(f"interval_sec{i}", 120)))
            self._var_intervals.append(var_iv)
            ttk.Spinbox(lf, from_=10, to=3600, width=5,
                        textvariable=var_iv).grid(row=0, column=1, sticky="w", padx=(0, 0))
            ttk.Label(lf, text="秒").grid(row=0, column=2, sticky="w", padx=(0, 8))

            ttk.Label(lf, text="OBSソース名:").grid(row=0, column=3, sticky="w", padx=(0, 4))
            ent = ttk.Entry(lf, font=("", 11))
            ent.grid(row=0, column=4, sticky="ew", pady=(0, 3))
            ent.insert(0, settings.get(f"source{i}", ""))
            self.ent_sources.append(ent)
            btn_src_drop = tk.Button(lf, text="▼", font=("", 8), padx=3, pady=0, width=2)
            btn_src_drop.grid(row=0, column=5, sticky="ns", padx=(2, 0), pady=(0, 3))
            btn_src_drop.config(command=lambda b=btn_src_drop, e=ent, ix=idx:
                self._open_source_picker(b, e, self.ent_scenes[ix] if ix < len(self.ent_scenes) else None, ix))

            # シーン行: [自動ON/OFF☐] [送信有効シーン名: / シーン検知で定期ON/OFF:] [entry(colspan=3)]
            var_ss = tk.BooleanVar(value=settings.get(f"scene_send{i}", False))
            self._var_scene_sends.append(var_ss)
            lbl_scene = ttk.Label(lf, text="対象シーン名:")

            def _make_scene_label_updater(var=var_ss):
                def _update():
                    self._refresh_capture_buttons()
                return _update

            tk.Checkbutton(lf, text="シーン限定定期有効化", variable=var_ss,
                           command=_make_scene_label_updater()).grid(
                row=1, column=0, columnspan=3, sticky="w", padx=(0, 2))
            lbl_scene.grid(row=1, column=3, sticky="e", padx=(0, 4))
            var_scene_str = tk.StringVar(value=settings.get(f"scene{i}", ""))
            var_scene_str.trace_add("write", lambda *a: self._refresh_capture_buttons())
            ent_scene = ttk.Entry(lf, font=("", 10), textvariable=var_scene_str)
            ent_scene.grid(row=1, column=4, sticky="ew", pady=(0, 4))
            self.ent_scenes.append(ent_scene)
            btn_scene_drop = tk.Button(lf, text="▼", font=("", 8), padx=3, pady=0, width=2)
            btn_scene_drop.grid(row=1, column=5, sticky="ns", padx=(2, 0), pady=(0, 4))
            btn_scene_drop.config(command=lambda b=btn_scene_drop, e=ent_scene, ix=idx:
                self._open_scene_picker(b, e, ix))

            # キャプチャボタン
            btn_cap = tk.Button(
                lf,
                text="キャプチャ & AIに送る",
                bg="#007bff", fg="white",
                font=("", 11, "bold"),
                pady=1,
                command=lambda e=ent, n=i, ix=idx: self._capture_and_send(
                    e.get().strip(), n, ix, skip_dup=False)
            )
            btn_cap.grid(row=2, column=0, columnspan=6, sticky="ew", pady=(1, 3))
            self._capture_buttons.append(btn_cap)

            # 指示文（ラベルと入力を同一サブフレーム内でpackして密着させる）
            txt_row = ttk.Frame(lf)
            txt_row.grid(row=3, column=0, columnspan=6, sticky="ew", pady=(2, 4))
            txt_row.columnconfigure(1, weight=1)
            ttk.Label(txt_row, text="指示文:").grid(row=0, column=0, sticky="nw", padx=(0, 4))
            txt = tk.Text(txt_row, height=2, font=("", 9))
            txt.grid(row=0, column=1, sticky="ew")
            txt.insert(tk.END, settings.get(f"prompt{i}", _DEFAULT_PROMPT))
            self.txt_prompts.append(txt)

            # プレビュー
            lbl_prev = ttk.Label(
                lf,
                text="[ キャプチャ待ち ]\n取込み上限 1280×720 / 送信上限 640×640（縦横比保持）\n例）X:Y比が16:9の場合、1280×720取込み → 640×360で送信\n例）X:Y比が1:1の場合、720×720取込み → 640×640で送信",
                background="#dddddd", anchor="center", justify="center"
            )
            lbl_prev.grid(row=5, column=0, columnspan=6, sticky="ew", ipady=0)
            self.lbl_previews.append(lbl_prev)

            # スロットごとのメッセージ表示
            lbl_msg = tk.Label(lf, text="", foreground="gray", anchor="w",
                               font=("", 9), padx=4, pady=1)
            lbl_msg.grid(row=6, column=0, columnspan=6, sticky="ew", pady=(2, 0))
            self.lbl_slot_msgs.append(lbl_msg)

        # ── フッター ──
        footer = ttk.Frame(self.panel, padding=(8, 4, 8, 8))
        footer.pack(fill=tk.X)
        footer.columnconfigure(0, weight=1)

        tk.Button(
            footer,
            text="設定を保存",
            bg="#6c757d", fg="white",
            command=self._save_settings
        ).grid(row=0, column=0, sticky="ew")

        self.lbl_status = ttk.Label(footer, text="待機中...", foreground="gray")
        self.lbl_status.grid(row=1, column=0, pady=(4, 0))

        self._refresh_capture_buttons()

    def _refresh_capture_buttons(self):
        """チェックOFF + シーン名入力あり の場合、現在シーンと不一致ならボタンをdisabled にする。"""
        try:
            for idx in range(len(self._capture_buttons)):
                btn = self._capture_buttons[idx]
                if not btn.winfo_exists():
                    continue
                scene_send = self._var_scene_sends[idx].get()
                scene_name = self.ent_scenes[idx].get().strip()
                if not scene_send and scene_name:
                    # チェックOFF + シーン名入力あり: シーン不一致時はdisabled
                    if self._current_obs_scene != scene_name:
                        btn.config(state="disabled", bg="#aaaaaa", fg="#666666")
                    else:
                        btn.config(state="normal", bg="#007bff", fg="white")
                else:
                    btn.config(state="normal", bg="#007bff", fg="white")
        except Exception:
            pass

    # ==========================================
    # 設定保存
    # ==========================================
    def _save_settings(self):
        s = self.get_settings()
        s["plugin_enabled"] = self._var_plugin_enabled.get()
        s["skip_duplicate"] = self._var_skip_dup.get()
        s["auto_stop_min"]  = int(self._var_stop_min.get()) if self._var_auto_stop.get() else 0

        for i, (ent, txt, ent_scene, var_as, var_iv, var_ss) in enumerate(
                zip(self.ent_sources, self.txt_prompts, self.ent_scenes,
                    self._var_auto_sends, self._var_intervals, self._var_scene_sends), 1):
            s[f"source{i}"]       = ent.get().strip()
            s[f"prompt{i}"]       = txt.get("1.0", tk.END).strip()
            s[f"scene{i}"]        = ent_scene.get().strip()
            s[f"auto_send{i}"]    = var_as.get()
            try:
                s[f"interval_sec{i}"] = max(10, int(var_iv.get()))
            except ValueError:
                pass
            s[f"scene_send{i}"]   = var_ss.get()

        # 定期ONスロットがあればループ起動
        if self.is_running and any(s.get(f"auto_send{i}", False) for i in range(1, 5)):
            self._start_auto_loop()

        self.save_settings(s)

        # シーン設定変更があればEventClientを再起動
        if self.is_running:
            self._stop_event_listener()
            self._start_event_listener()

        if self.lbl_status and self.lbl_status.winfo_exists():
            self.lbl_status.config(text="✅ 保存しました", foreground="green")

    # ==========================================
    # キャプチャ・送信
    # ==========================================
    def _image_hash(self, img_bytes):
        img = Image.open(io.BytesIO(img_bytes)).convert("L").resize((16, 16))
        return hash(img.tobytes())

    def _slot_msg(self, source_idx, text, color="gray"):
        """スロットのメッセージラベルを更新する（UIスレッド安全）。
        color: "gray" / "green" / "orange"（警告） / "red"（エラー）
        """
        _STYLE = {
            "gray":   {"foreground": "gray",    "background": ""},
            "green":  {"foreground": "#00aa44", "background": ""},
            "orange": {"foreground": "#FFD700", "background": "#3d2800"},
            "red":    {"foreground": "#FF8080", "background": "#3a0000"},
        }
        style = _STYLE.get(color, _STYLE["gray"])
        try:
            if source_idx >= len(self.lbl_slot_msgs):
                return  # パネル未表示時はスキップ
            lbl = self.lbl_slot_msgs[source_idx]
            if lbl.winfo_exists():
                bg = style["background"] or lbl.master.cget("background")
                lbl.config(text=text, foreground=style["foreground"], background=bg)
        except Exception:
            pass

    def _capture_and_send(self, source_name, source_num, source_idx, skip_dup=False):
        # プラグイン無効時はスキップ
        if not self.get_settings().get("plugin_enabled", True):
            return False

        if not source_name:
            if not skip_dup:
                self._slot_msg(source_idx, "⚠ OBSソース名を入力してください", "orange")
            return False

        try:
            prompt_text = self.txt_prompts[source_idx].get("1.0", tk.END).strip()
        except Exception:
            prompt_text = self.get_settings().get(f"prompt{source_num}", _DEFAULT_PROMPT)

        conn     = _load_obs_conn()
        host     = conn.get("host", "127.0.0.1")
        port     = int(conn.get("port", 4455))
        password = conn.get("password", "")

        # ── OBS接続 ──
        try:
            cl = obs.ReqClient(host=host, port=port, password=password, timeout=3)
        except Exception as e:
            if not skip_dup:
                self._slot_msg(source_idx, f"❌ OBS接続エラー: {e}", "red")
            logger.error(f"[{self.PLUGIN_NAME}] OBS接続失敗: {e}")
            return False

        # ── スクリーンショット取得 ──
        try:
            res = cl.get_source_screenshot(source_name, "jpeg", 1280, 720, 80)
        except Exception as e:
            if not skip_dup:
                err_str = str(e).lower()
                if ("not found" in err_str or "no source" in err_str
                        or "resourcenotfound" in err_str
                        or getattr(getattr(e, "req_status", None), "code", None) == 600):
                    self._slot_msg(source_idx, f"⚠ ソース '{source_name}' がOBSに見つかりません", "orange")
                else:
                    self._slot_msg(source_idx, f"❌ OBSエラー: {e}", "red")
            logger.error(f"[{self.PLUGIN_NAME}] スクリーンショット取得失敗 (ソース'{source_name}'): {e}")
            return False

        img_data = res.image_data
        if img_data.startswith("data:image/jpeg;base64,"):
            img_data = img_data.replace("data:image/jpeg;base64,", "")
        img_bytes = base64.b64decode(img_data)

        if skip_dup:
            h = self._image_hash(img_bytes)
            if h == self._last_hashes[source_idx]:
                logger.debug(f"[{self.PLUGIN_NAME}] ソース{source_num} は前回と同じ画像のためスキップ。")
                return False
            self._last_hashes[source_idx] = h

        self._set_preview(source_idx, img_bytes)

        # ── AIへ送信 ──
        if not self.plugin_queue:
            if not skip_dup:
                self._slot_msg(source_idx, "⚠ AIライブ未接続（ライブ接続を開始してください）", "orange")
            logger.debug(f"[{self.PLUGIN_NAME}] plugin_queue なし: キャプチャのみ実行（AI未送信）")
            return False

        try:
            img = Image.open(io.BytesIO(img_bytes))
            if img.mode != "RGB":
                img = img.convert("RGB")
            img.thumbnail((640, 640))
            out = io.BytesIO()
            img.save(out, format="JPEG", quality=80)

            self.send_image(self.plugin_queue, out.getvalue(), "image/jpeg")
            self.send_text(self.plugin_queue, prompt_text)
        except Exception as e:
            self._slot_msg(source_idx, f"❌ AI送信エラー: {e}", "red")
            logger.error(f"[{self.PLUGIN_NAME}] AI送信失敗 (ソース'{source_name}'): {e}")
            return False

        self._slot_msg(source_idx, f"✅ 送信しました", "green")
        if self.lbl_status and self.lbl_status.winfo_exists():
            self.lbl_status.config(
                text=f"✅ ソース{source_num}（{source_name}）を送信しました！",
                foreground="green"
            )
        logger.info(f"[{self.PLUGIN_NAME}] ソース'{source_name}' をAIに送信しました。")
        return True

    def _set_preview(self, idx, img_bytes):
        try:
            if idx >= len(self.lbl_previews):
                return  # パネル未表示時はスキップ
            lbl = self.lbl_previews[idx]
            if not lbl.winfo_exists():
                return
            img = Image.open(io.BytesIO(img_bytes))
            if img.mode != "RGB":
                img = img.convert("RGB")
            img.thumbnail((360, 160))
            photo = ImageTk.PhotoImage(img)
            self.preview_photos[idx] = photo
            lbl.config(image=photo, text="")
        except Exception as e:
            logger.warning(f"[{self.PLUGIN_NAME}] プレビュー表示失敗: {e}")
