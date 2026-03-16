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
            "enabled":       True,
            "source1": "", "prompt1": _DEFAULT_PROMPT, "scene1": "",
            "source2": "", "prompt2": _DEFAULT_PROMPT, "scene2": "",
            "source3": "", "prompt3": _DEFAULT_PROMPT, "scene3": "",
            "source4": "", "prompt4": _DEFAULT_PROMPT, "scene4": "",
            "auto_send1": False, "auto_send2": False,
            "auto_send3": False, "auto_send4": False,
            "scene_send1": False, "scene_send2": False,
            "scene_send3": False, "scene_send4": False,
            "auto_send":     False,
            "interval_sec":  120,
            "auto_stop_min": 60,
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

        if self.get_settings().get("auto_send", False):
            self._start_auto_loop()

        self._start_event_listener()

    def stop(self):
        self.is_running   = False
        self.plugin_queue = None
        self._stop_event_listener()
        logger.debug(f"[{self.PLUGIN_NAME}] キューを切断しました。")

    # ==========================================
    # 定期送信ループ
    # ==========================================
    def _start_auto_loop(self):
        self._auto_start_time = time.time()
        if self._auto_thread is None or not self._auto_thread.is_alive():
            self._auto_thread = threading.Thread(target=self._auto_loop, daemon=True)
            self._auto_thread.start()
            logger.info(f"[{self.PLUGIN_NAME}] 定期送信を開始しました。")

    def _auto_loop(self):
        s        = self.get_settings()
        interval = max(10, int(s.get("interval_sec", 120)))
        stop_min = int(s.get("auto_stop_min", 0))
        last_send = time.time()

        while self.is_running and self.get_settings().get("auto_send", False):
            now = time.time()

            if stop_min > 0 and (now - self._auto_start_time) >= stop_min * 60:
                logger.info(f"[{self.PLUGIN_NAME}] {stop_min}分経過のため定期送信を自動停止しました。")
                s = self.get_settings()
                s["auto_send"] = False
                self.save_settings(s)
                self._update_auto_send_ui(False)
                break

            if now - last_send >= interval:
                s = self.get_settings()
                for idx in range(4):
                    if not self.is_running:
                        break
                    if not s.get(f"auto_send{idx + 1}", False):
                        continue
                    # シーン名が設定されている場合は一致しないとスキップ（チェック状態に関わらず）
                    active_scene = s.get(f"scene{idx + 1}", "").strip()
                    if active_scene and self._current_obs_scene != active_scene:
                        continue
                    source = s.get(f"source{idx + 1}", "").strip()
                    if source:
                        self._capture_and_send(source, idx + 1, idx,
                                               skip_dup=s.get("skip_duplicate", True))
                last_send = time.time()

            time.sleep(1.0)

    def _update_auto_send_ui(self, value: bool):
        try:
            if hasattr(self, "_var_auto_send") and self._var_auto_send.get() != value:
                self._var_auto_send.set(value)
                self._refresh_auto_send_state()
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
          send       : {"command":"AI-SS-Sender","action":"send","slot":1}
                        → ソース1～4を即時キャプチャしてAIに送る
          set_source : {"command":"AI-SS-Sender","action":"set_source","slot":1,"name":"ソース名"}
                        → スロット1～4のソース名を変更して設定を保存する
          set_interval: {"command":"AI-SS-Sender","action":"set_interval","seconds":60}
                        → 定期送信の間隔（秒）を変更して設定を保存する
          auto       : {"command":"AI-SS-Sender","action":"auto","enabled":true}
                        → 定期送信のON/OFFを切り替える
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
            s = self.get_settings()
            s["interval_sec"] = seconds
            self.save_settings(s)
            try:
                if hasattr(self, "_var_interval"):
                    self._var_interval.set(str(seconds))
            except Exception:
                pass
            logger.info(f"[{self.PLUGIN_NAME}] WSコマンド set_interval: 定期間隔を{seconds}秒に変更しました。")

        elif action == "auto":
            enabled = bool(ev.get("enabled", False))
            s = self.get_settings()
            s["auto_send"] = enabled
            self.save_settings(s)
            if enabled and self.is_running:
                self._start_auto_loop()
            self._update_auto_send_ui(enabled)
            logger.info(f"[{self.PLUGIN_NAME}] WSコマンド auto: 定期送信を{'ON' if enabled else 'OFF'}にしました。")

        else:
            logger.warning(f"[{self.PLUGIN_NAME}] WSコマンド: 不明なaction「{action}」を受信しました。")

    # ==========================================
    # UI
    # ==========================================
    def open_settings_ui(self, parent_window):
        if hasattr(self, "panel") and self.panel is not None and self.panel.winfo_exists():
            self.panel.lift()
            return

        self.panel = tk.Toplevel(parent_window)
        self.panel.title(self.PLUGIN_NAME)
        self.panel.geometry("860x820")
        self.panel.resizable(True, True)
        self.panel.attributes("-topmost", True)

        settings = self.get_settings()

        # ── ヘッダー ──
        header = ttk.Frame(self.panel, padding=(12, 8, 12, 4))
        header.pack(fill=tk.X)
        ttk.Label(
            header,
            text="※OBS接続設定は「OBS画面AI実況」と共有",
            foreground="gray"
        ).pack(side="right")

        ttk.Separator(self.panel, orient="horizontal").pack(fill=tk.X, padx=8)

        # ── 定期送信設定 ──
        auto_f = ttk.Frame(self.panel, padding=(12, 6, 12, 4))
        auto_f.pack(fill=tk.X)

        self._var_auto_send = tk.BooleanVar(value=settings.get("auto_send", False))
        tk.Checkbutton(
            auto_f, text="定期送信",
            variable=self._var_auto_send,
            command=self._on_auto_send_toggle
        ).grid(row=0, column=0, sticky="w")

        ttk.Label(auto_f, text="間隔:").grid(row=0, column=1, padx=(12, 2), sticky="w")
        self._var_interval = tk.StringVar(value=str(settings.get("interval_sec", 120)))
        ttk.Spinbox(auto_f, from_=10, to=3600, width=6,
                    textvariable=self._var_interval).grid(row=0, column=2, sticky="w")
        ttk.Label(auto_f, text="秒").grid(row=0, column=3, padx=(2, 16), sticky="w")

        self._var_auto_stop = tk.BooleanVar(value=settings.get("auto_stop_min", 0) > 0)
        tk.Checkbutton(
            auto_f, text="自動停止:",
            variable=self._var_auto_stop
        ).grid(row=0, column=4, sticky="w")
        self._var_stop_min = tk.StringVar(value=str(max(1, settings.get("auto_stop_min", 60))))
        ttk.Spinbox(auto_f, from_=1, to=600, width=5,
                    textvariable=self._var_stop_min).grid(row=0, column=5, sticky="w")
        ttk.Label(auto_f, text="分後").grid(row=0, column=6, padx=(2, 16), sticky="w")

        self._var_skip_dup = tk.BooleanVar(value=settings.get("skip_duplicate", True))
        tk.Checkbutton(
            auto_f, text="前回と同じ画像は送信しない",
            variable=self._var_skip_dup
        ).grid(row=0, column=7, sticky="w")

        self._lbl_auto_status = ttk.Label(auto_f, text="", foreground="gray")
        self._lbl_auto_status.grid(row=0, column=8, padx=(12, 0), sticky="w")
        self._refresh_auto_send_state()

        ttk.Separator(self.panel, orient="horizontal").pack(fill=tk.X, padx=8)

        # ── ソースカード（2列グリッド） ──
        grid_f = ttk.Frame(self.panel, padding=(8, 8, 8, 4))
        grid_f.pack(fill=tk.BOTH, expand=True)
        grid_f.columnconfigure(0, weight=1, uniform="col")
        grid_f.columnconfigure(1, weight=1, uniform="col")

        self.ent_sources      = []
        self.txt_prompts      = []
        self.ent_scenes       = []
        self.lbl_previews     = []
        self.preview_photos   = [None, None, None, None]
        self._var_auto_sends  = []
        self._var_scene_sends = []
        self._capture_buttons = []

        positions = [(0, 0), (0, 1), (1, 0), (1, 1)]

        for i in range(1, 5):
            row_g, col_g = positions[i - 1]
            lf = ttk.LabelFrame(grid_f, text=f" ソース {i} ", padding=8)
            lf.grid(row=row_g, column=col_g, sticky="nsew", padx=5, pady=5)
            lf.columnconfigure(2, weight=1)
            grid_f.rowconfigure(row_g, weight=1)
            idx = i - 1

            # ソース名行: [定期☐] [OBSソース名:] [entry]
            var_as = tk.BooleanVar(value=settings.get(f"auto_send{i}", False))
            self._var_auto_sends.append(var_as)
            tk.Checkbutton(lf, text="定期", variable=var_as).grid(
                row=0, column=0, sticky="w", padx=(0, 2))
            ttk.Label(lf, text="OBSソース名:").grid(row=0, column=1, sticky="w", padx=(0, 4))
            ent = ttk.Entry(lf, font=("", 11))
            ent.grid(row=0, column=2, sticky="ew", pady=(0, 3))
            ent.insert(0, settings.get(f"source{i}", ""))
            self.ent_sources.append(ent)

            # シーン行: [自動ON/OFF☐] [送信有効シーン: / シーン検知で定期ON/OFF:] [entry]
            var_ss = tk.BooleanVar(value=settings.get(f"scene_send{i}", False))
            self._var_scene_sends.append(var_ss)
            lbl_scene = ttk.Label(
                lf,
                text="定期送信対象シーン名:" if var_ss.get() else "送信許可シーン名:"
            )

            def _make_scene_label_updater(lbl=lbl_scene, var=var_ss):
                def _update():
                    lbl.config(text="シーン検知で定期ON/OFF:" if var.get() else "送信有効シーン:")
                    self._refresh_capture_buttons()
                return _update

            tk.Checkbutton(lf, text="自動ON/OFF", variable=var_ss,
                           command=_make_scene_label_updater()).grid(
                row=1, column=0, sticky="w", padx=(0, 2))
            lbl_scene.grid(row=1, column=1, sticky="w", padx=(0, 4))
            var_scene_str = tk.StringVar(value=settings.get(f"scene{i}", ""))
            var_scene_str.trace_add("write", lambda *a: self._refresh_capture_buttons())
            ent_scene = ttk.Entry(lf, font=("", 10), textvariable=var_scene_str)
            ent_scene.grid(row=1, column=2, sticky="ew", pady=(0, 4))
            self.ent_scenes.append(ent_scene)

            # キャプチャボタン
            btn_cap = tk.Button(
                lf,
                text="📸  キャプチャ & AIに送る",
                bg="#007bff", fg="white",
                font=("", 11, "bold"),
                pady=6,
                command=lambda e=ent, n=i, ix=idx: self._capture_and_send(
                    e.get().strip(), n, ix, skip_dup=False)
            )
            btn_cap.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(2, 6))
            self._capture_buttons.append(btn_cap)

            # 指示テキスト
            ttk.Label(lf, text="指示テキスト:").grid(row=3, column=0, columnspan=3, sticky="w")
            txt = tk.Text(lf, height=2, font=("", 9))
            txt.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(2, 6))
            txt.insert(tk.END, settings.get(f"prompt{i}", _DEFAULT_PROMPT))
            self.txt_prompts.append(txt)

            # プレビュー
            lbl_prev = ttk.Label(lf, text="[ キャプチャ待ち ]", background="#dddddd", anchor="center")
            lbl_prev.grid(row=5, column=0, columnspan=3, sticky="ew", ipady=25)
            self.lbl_previews.append(lbl_prev)

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

    def _on_auto_send_toggle(self):
        if self._var_auto_send.get() and self.is_running:
            self._start_auto_loop()
        self._refresh_auto_send_state()

    def _refresh_auto_send_state(self):
        try:
            if not hasattr(self, "_lbl_auto_status"):
                return
            if self.is_running and self._var_auto_send.get():
                elapsed = int(time.time() - self._auto_start_time) // 60
                self._lbl_auto_status.config(text=f"▶ 送信中 ({elapsed}分経過)", foreground="green")
            elif self._var_auto_send.get():
                self._lbl_auto_status.config(text="（配信開始後に有効）", foreground="gray")
            else:
                self._lbl_auto_status.config(text="", foreground="gray")
        except Exception:
            pass

    # ==========================================
    # 設定保存
    # ==========================================
    def _save_settings(self):
        s = self.get_settings()
        s["enabled"]        = True
        s["auto_send"]      = self._var_auto_send.get()
        s["skip_duplicate"] = self._var_skip_dup.get()
        try:
            s["interval_sec"] = max(10, int(self._var_interval.get()))
        except ValueError:
            pass
        s["auto_stop_min"] = int(self._var_stop_min.get()) if self._var_auto_stop.get() else 0

        for i, (ent, txt, ent_scene, var_as, var_ss) in enumerate(
                zip(self.ent_sources, self.txt_prompts, self.ent_scenes,
                    self._var_auto_sends, self._var_scene_sends), 1):
            s[f"source{i}"]     = ent.get().strip()
            s[f"prompt{i}"]     = txt.get("1.0", tk.END).strip()
            s[f"scene{i}"]      = ent_scene.get().strip()
            s[f"auto_send{i}"]  = var_as.get()
            s[f"scene_send{i}"] = var_ss.get()

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

    def _capture_and_send(self, source_name, source_num, source_idx, skip_dup=False):
        if not self.plugin_queue:
            if not skip_dup:
                messagebox.showwarning("警告", "配信を開始してから押してください。")
            return False
        if not source_name:
            if not skip_dup:
                messagebox.showwarning("警告", f"ソース {source_num} の名前を入力してください。")
            return False

        try:
            prompt_text = self.txt_prompts[source_idx].get("1.0", tk.END).strip()
        except Exception:
            prompt_text = self.get_settings().get(f"prompt{source_num}", _DEFAULT_PROMPT)

        conn     = _load_obs_conn()
        host     = conn.get("host", "127.0.0.1")
        port     = int(conn.get("port", 4455))
        password = conn.get("password", "")

        try:
            cl  = obs.ReqClient(host=host, port=port, password=password, timeout=3)
            res = cl.get_source_screenshot(source_name, "jpeg", 1280, 720, 80)

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

            img = Image.open(io.BytesIO(img_bytes))
            if img.mode != "RGB":
                img = img.convert("RGB")
            img.thumbnail((640, 640))
            out = io.BytesIO()
            img.save(out, format="JPEG", quality=80)

            self.send_image(self.plugin_queue, out.getvalue(), "image/jpeg")
            self.send_text(self.plugin_queue, prompt_text)

            if self.lbl_status and self.lbl_status.winfo_exists():
                self.lbl_status.config(
                    text=f"✅ ソース{source_num}（{source_name}）を送信しました！",
                    foreground="green"
                )
            logger.info(f"[{self.PLUGIN_NAME}] ソース'{source_name}' をAIに送信しました。")
            return True

        except Exception as e:
            if self.lbl_status and self.lbl_status.winfo_exists():
                self.lbl_status.config(text=f"❌ エラー: {e}", foreground="red")
            logger.error(f"[{self.PLUGIN_NAME}] OBSキャプチャ失敗 (ソース'{source_name}'): {e}")
            return False

    def _set_preview(self, idx, img_bytes):
        try:
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
