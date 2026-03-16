"""
テロップ読み上げプラグイン
http://localhost:8000/data.json をポーリングしてテロップ変化を検知し、
選択したバックエンド（SAPI / VOICEVOX / COEIROINK v2）で読み上げる。
"""
import io
import json
import os
import queue
import re
import subprocess
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox
import urllib.request
import urllib.parse

import logger
from plugin_manager import BasePlugin

_DATA_URL  = "http://localhost:8000/data.json"
_POLL_SEC  = 0.5


def _load_obs_conn() -> dict:
    here = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(here, "obs_capture.json")
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _try_import_obs():
    try:
        import obsws_python as obs
        return obs
    except ImportError:
        return None

# ── SAPI：PowerShell経由でCOMを操作（pywin32不要）──────────────────────
def _ps_run(command, timeout=8):
    """PowerShellコマンドを実行して標準出力を行リストで返す。"""
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", command],
            capture_output=True, text=True, timeout=timeout,
            creationflags=0x08000000  # CREATE_NO_WINDOW
        )
        return [l.strip() for l in r.stdout.strip().splitlines() if l.strip()]
    except Exception:
        return []

def _get_sapi_voices():
    return _ps_run(
        "$v=$null; try{$v=New-Object -ComObject SAPI.SpVoice}catch{exit}; "
        "$vs=$v.GetVoices(); 0..($vs.Count-1)|%{$vs.Item($_).GetDescription()}"
    )

def _get_sapi_devices():
    return _ps_run(
        "$v=$null; try{$v=New-Object -ComObject SAPI.SpVoice}catch{exit}; "
        "$os=$v.GetAudioOutputs(); 0..($os.Count-1)|%{$os.Item($_).GetDescription()}"
    )

def _try_import_sounddevice():
    try:
        import sounddevice as sd
        import numpy as np
        return sd, np
    except Exception:
        return None, None


def _resample_wav(wav_bytes: bytes, target_hz: int) -> bytes:
    """WAVを target_hz にリサンプリングして返す（標準ライブラリのみ）。"""
    import wave as wm, struct, array
    with wm.open(io.BytesIO(wav_bytes)) as wf:
        src_hz = wf.getframerate()
        ch     = wf.getnchannels()
        sw     = wf.getsampwidth()
        n      = wf.getnframes()
        raw    = wf.readframes(n)
    if src_hz == target_hz:
        return wav_bytes
    n_src = n * ch
    samples = struct.unpack(f"<{n_src}h", raw)
    n_dst = int(n * target_hz / src_hz) * ch
    ratio = src_hz / target_hz
    out = array.array("h")
    for i in range(n_dst):
        fi  = i * ratio
        lo  = int(fi)
        hi  = min(lo + 1, n_src - 1)
        s   = int(samples[lo] * (1.0 - (fi - lo)) + samples[hi] * (fi - lo))
        out.append(max(-32768, min(32767, s)))
    buf = io.BytesIO()
    with wm.open(buf, "wb") as wf:
        wf.setnchannels(ch)
        wf.setsampwidth(sw)
        wf.setframerate(target_hz)
        wf.writeframes(out.tobytes())
    return buf.getvalue()


class TelopReaderPlugin(BasePlugin):
    PLUGIN_ID   = "telop_reader"
    PLUGIN_NAME = "テロップ読み上げ"
    PLUGIN_TYPE = "TOOL"

    def __init__(self):
        super().__init__()
        self.is_running          = False
        self.is_connected        = True   # v1.22b: ライブ前からバッジをアクティブ表示
        self._tts_queue          = queue.Queue(maxsize=5)
        self._poll_thread        = None
        self._tts_thread         = None
        self._last_explain_key   = None   # (update_time, main)
        self._last_normal_keys   = set()  # set of (topic, main)
        self._current_obs_scene  = None   # OBS現在シーン（Noneは未取得）
        self._scene_client       = None   # obsws_python EventClient

    def get_default_settings(self):
        return {
            "enabled":          True,
            "tts_enabled":      False,
            "backend":          "sapi",       # "sapi" / "voicevox" / "coeiroink"
            # SAPI
            "sapi_voice":       "",
            "sapi_device":      "",
            "sapi_rate":        0,
            "sapi_volume":      100,
            # VOICEVOX
            "vv_url":           "http://localhost:50021",
            "vv_speaker":       0,
            "vv_device":        "",
            "vv_speed":         1.0,
            "vv_volume":        0.6,
            # COEIROINK v2
            "ci_url":           "http://localhost:50032",
            "ci_speaker_uuid":  "",
            "ci_style_id":      0,
            "ci_device":        "",
            "ci_speed":         1.0,
            "ci_volume":        0.6,
            # フィルター
            "active_scene":     "",    # 空=すべてのシーンで読む
            "skip_system_msg":  False, # [topic]始まりのシステムメッセージをスキップ
            # 共通
            "read_main":        True,
            "read_topic":       False,
            "read_explain":     True,
            "read_normal":      True,
        }

    # ==========================================
    # ライフサイクル
    # ==========================================
    def start(self, prompt_config, plugin_queue):
        self.is_running = True
        s = self.get_settings()
        if s.get("tts_enabled", False):
            self._start_threads()
        self._start_scene_listener(s)
        logger.debug(f"[{self.PLUGIN_NAME}] 起動しました。")

    def stop(self):
        self.is_running = False
        self._stop_scene_listener()
        logger.debug(f"[{self.PLUGIN_NAME}] 停止しました。")

    # ==========================================
    # OBSシーン追跡
    # ==========================================
    def _start_scene_listener(self, s=None):
        if s is None:
            s = self.get_settings()
        if not s.get("active_scene", "").strip():
            return  # シーン指定なし → 追跡不要
        obs = _try_import_obs()
        if obs is None:
            logger.warning(f"[{self.PLUGIN_NAME}] obsws-python が未インストールのためシーンフィルターは無効です。")
            return
        try:
            conn = _load_obs_conn()
            # 現在シーンを初期取得
            cl = obs.ReqClient(
                host=conn.get("host", "127.0.0.1"),
                port=int(conn.get("port", 4455)),
                password=conn.get("password", ""),
                timeout=3
            )
            self._current_obs_scene = cl.get_current_program_scene().current_program_scene_name
            cl.disconnect()
            # イベントで追跡
            self._scene_client = obs.EventClient(
                host=conn.get("host", "127.0.0.1"),
                port=int(conn.get("port", 4455)),
                password=conn.get("password", "")
            )
            self._scene_client.callback.register(self._on_scene_changed)
            logger.info(f"[{self.PLUGIN_NAME}] OBSシーン追跡開始（現在: {self._current_obs_scene}）")
        except Exception as e:
            logger.debug(f"[{self.PLUGIN_NAME}] OBS接続失敗（シーンフィルター無効）: {e}")
            self._scene_client = None

    def _stop_scene_listener(self):
        if self._scene_client:
            try:
                self._scene_client.disconnect()
            except Exception:
                pass
            self._scene_client = None

    def _on_scene_changed(self, data):
        try:
            self._current_obs_scene = data.scene_name
            logger.debug(f"[{self.PLUGIN_NAME}] シーン変更: {self._current_obs_scene}")
        except AttributeError:
            pass

    def _is_active_scene(self, s) -> bool:
        """active_scene 設定に合致するシーンかどうかを返す（空=常にTrue）。"""
        scene = s.get("active_scene", "").strip()
        if not scene:
            return True
        if self._current_obs_scene is None:
            return True  # 未取得（OBS未接続）はスルー
        return self._current_obs_scene == scene

    def _is_system_telop(self, telop) -> bool:
        """topic が [ で始まるシステムメッセージかどうかを返す。"""
        topic = telop.get("topic", "")
        return bool(topic) and topic.startswith("[")

    def _start_threads(self):
        self._last_explain_key = None
        self._last_normal_keys = set()
        if self._poll_thread is None or not self._poll_thread.is_alive():
            self._poll_thread = threading.Thread(target=self._poll_loop, daemon=True)
            self._poll_thread.start()
        if self._tts_thread is None or not self._tts_thread.is_alive():
            self._tts_thread = threading.Thread(target=self._tts_worker, daemon=True)
            self._tts_thread.start()

    # ==========================================
    # ポーリングループ
    # ==========================================
    def _poll_loop(self):
        while self.is_running:
            s = self.get_settings()
            if not s.get("tts_enabled", False):
                time.sleep(1.0)
                continue
            try:
                with urllib.request.urlopen(_DATA_URL, timeout=2) as resp:
                    data = json.loads(resp.read())
                self._check_explain(data, s)
                self._check_normal(data, s)
            except Exception:
                pass
            time.sleep(_POLL_SEC)

    def _check_explain(self, data, s):
        if not s.get("read_explain", True):
            return
        if not self._is_active_scene(s):
            return
        exp = data.get("explain", {})
        if not exp.get("visible", False):
            return
        if s.get("skip_system_msg", False) and self._is_system_telop(exp):
            return
        key = (exp.get("update_time", 0), exp.get("main", ""))
        if key == self._last_explain_key:
            return
        self._last_explain_key = key
        text = self._build_text(exp, s)
        if text:
            self._enqueue(text)

    def _check_normal(self, data, s):
        if not s.get("read_normal", True):
            return
        if not self._is_active_scene(s):
            return
        telops = data.get("normal", {}).get("active_telops", [])
        for t in telops:
            if s.get("skip_system_msg", False) and self._is_system_telop(t):
                continue
            key = (t.get("topic", ""), t.get("main", ""))
            if key in self._last_normal_keys:
                continue
            self._last_normal_keys.add(key)
            text = self._build_text(t, s)
            if text:
                self._enqueue(text)
        # 古いキーを整理（active_telopが消えたら忘れる）
        current = {(t.get("topic", ""), t.get("main", "")) for t in telops}
        self._last_normal_keys &= current

    _RE_XML_TAG = re.compile(r"<[^>]+>")

    def _strip_xml(self, text: str) -> str:
        """XMLタグを除去する。"""
        return self._RE_XML_TAG.sub("", text).strip()

    def _build_text(self, telop, s):
        parts = []
        if s.get("read_topic", False):
            topic = self._strip_xml(telop.get("topic", ""))
            if topic and not topic.startswith("["):  # システムコードを除外
                parts.append(topic)
        if s.get("read_main", True):
            main = self._strip_xml(telop.get("main", ""))
            if main:
                parts.append(main)
        return "\n".join(parts)

    def _enqueue(self, text):
        try:
            self._tts_queue.put_nowait(text)
        except queue.Full:
            # キューが満杯なら古いものを捨てて追加
            try:
                self._tts_queue.get_nowait()
                self._tts_queue.put_nowait(text)
            except Exception:
                pass

    # ==========================================
    # TTSワーカー
    # ==========================================
    def _tts_worker(self):
        while self.is_running:
            try:
                text = self._tts_queue.get(timeout=1.0)
            except queue.Empty:
                continue
            s = self.get_settings()
            if not s.get("tts_enabled", False):
                continue
            try:
                backend = s.get("backend", "sapi")
                if backend == "sapi":
                    self._speak_sapi(text, s)
                elif backend == "voicevox":
                    self._speak_voicevox(text, s)
                elif backend == "coeiroink":
                    self._speak_coeiroink(text, s)
            except Exception as e:
                logger.warning(f"[{self.PLUGIN_NAME}] 読み上げエラー: {e}")
                try:
                    if hasattr(self, "panel") and self.panel and self.panel.winfo_exists():
                        err = str(e)
                        self.panel.after(0, lambda msg=err: self._lbl_status.config(
                            text=f"❌ {msg}", foreground="red"))
                except Exception:
                    pass

    # ── SAPI ──────────────────────────────────
    def _speak_sapi(self, text, s):
        voice_name  = s.get("sapi_voice", "")
        device_name = s.get("sapi_device", "")
        rate        = int(s.get("sapi_rate", 0))
        volume      = int(s.get("sapi_volume", 100))

        lines = ["$v = New-Object -ComObject SAPI.SpVoice"]
        if device_name:
            dn = device_name.replace("'", "''")
            lines.append(
                f"$os=$v.GetAudioOutputs(); $t=0..($os.Count-1)|"
                f"Where-Object{{$os.Item($_).GetDescription() -eq '{dn}'}}|"
                f"Select-Object -First 1; if($t -ne $null){{$v.AudioOutput=$os.Item($t)}}"
            )
        if voice_name:
            vn = voice_name.replace("'", "''")
            lines.append(
                f"$vs=$v.GetVoices(); $tv=0..($vs.Count-1)|"
                f"Where-Object{{$vs.Item($_).GetDescription() -eq '{vn}'}}|"
                f"Select-Object -First 1; if($tv -ne $null){{$v.Voice=$vs.Item($tv)}}"
            )
        lines.append(f"$v.Rate={rate}; $v.Volume={volume}")
        escaped = text.replace("'", "''")
        lines.append(f"$v.Speak('{escaped}')")

        subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", "; ".join(lines)],
            timeout=30, creationflags=0x08000000
        )

    # ── VOICEVOX ──────────────────────────────
    def _speak_voicevox(self, text, s):
        url     = s.get("vv_url", "http://localhost:50021").rstrip("/")
        speaker = int(s.get("vv_speaker", 0))
        speed   = float(s.get("vv_speed", 1.0))

        # audio_query
        req = urllib.request.Request(
            f"{url}/audio_query?text={urllib.parse.quote(text)}&speaker={speaker}",
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            query = json.loads(resp.read())

        query["speedScale"]  = speed
        query["volumeScale"] = float(s.get("vv_volume", 0.6))

        # synthesis
        req2 = urllib.request.Request(
            f"{url}/synthesis?speaker={speaker}",
            data=json.dumps(query).encode(),
            headers={"Content-Type": "application/json"},
            method="POST"
        )
        with urllib.request.urlopen(req2, timeout=30) as resp:
            wav_bytes = resp.read()

        self._play_wav(wav_bytes, s.get("vv_device", ""))

    def _play_wav(self, wav_bytes, device_name=""):
        if not wav_bytes or len(wav_bytes) < 44:
            raise RuntimeError(f"WAVデータが不正です（{len(wav_bytes) if wav_bytes else 0} bytes）")

        # sounddevice が使えれば優先（24000Hz対応・デバイス指定可）
        sd, np = _try_import_sounddevice()
        if sd is not None:
            try:
                import wave as wm
                with wm.open(io.BytesIO(wav_bytes)) as wf:
                    data = wf.readframes(wf.getnframes())
                    rate = wf.getframerate()
                    ch   = wf.getnchannels()
                    sw   = wf.getsampwidth()
                dtype = {1: np.int8, 2: np.int16, 4: np.int32}.get(sw, np.int16)
                audio = np.frombuffer(data, dtype=dtype)
                if ch == 2:
                    audio = audio.reshape(-1, 2)
                device_id = None
                if device_name:
                    for i, d in enumerate(sd.query_devices()):
                        if device_name in d["name"] and d["max_output_channels"] > 0:
                            device_id = i
                            break
                sd.play(audio, samplerate=rate, device=device_id)
                sd.wait()
                return
            except Exception as e:
                logger.warning(f"[{self.PLUGIN_NAME}] sounddevice 再生失敗（フォールバック）: {e}")

        # フォールバック: 22050Hz にリサンプリング → PowerShell SoundPlayer
        import tempfile
        resampled = _resample_wav(wav_bytes, 22050)
        with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as f:
            f.write(resampled)
            tmp = f.name
        try:
            path_esc = tmp.replace("'", "''")
            r = subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-Command",
                 f"(New-Object System.Media.SoundPlayer '{path_esc}').PlaySync()"],
                timeout=60, creationflags=0x08000000
            )
            if r.returncode != 0:
                raise RuntimeError(f"SoundPlayer 失敗 (exitcode={r.returncode})")
        finally:
            os.unlink(tmp)

    # ── COEIROINK v2 ──────────────────────────
    def _speak_coeiroink(self, text, s):
        url          = s.get("ci_url", "http://localhost:50032").rstrip("/")
        speaker_uuid = s.get("ci_speaker_uuid", "").strip()
        style_id     = int(s.get("ci_style_id", 0))
        speed        = float(s.get("ci_speed", 1.0))

        if not speaker_uuid:
            raise RuntimeError("COEIROINKのスピーカーが選択されていません")

        body = json.dumps({
            "speakerUuid":        speaker_uuid,
            "styleId":            style_id,
            "text":               text,
            "speedScale":         speed,
            "volumeScale":        float(s.get("ci_volume", 0.6)),
            "pitchScale":         0.0,
            "intonationScale":    1.0,
            "prePhonemeLength":   0.1,
            "postPhonemeLength":  0.1,
            "outputSamplingRate": 44100,
            "prosodyDetail":      [],
        }).encode()

        req = urllib.request.Request(
            f"{url}/v1/synthesis",
            data=body,
            headers={"Content-Type": "application/json"},
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            wav_bytes = resp.read()

        self._play_wav(wav_bytes, s.get("ci_device", ""))

    # ==========================================
    # UI
    # ==========================================
    def open_settings_ui(self, parent_window):
        if hasattr(self, "panel") and self.panel is not None and self.panel.winfo_exists():
            self.panel.lift()
            return

        self.panel = tk.Toplevel(parent_window)
        self.panel.title(self.PLUGIN_NAME)
        self.panel.geometry("500x640")
        self.panel.resizable(False, False)
        self.panel.attributes("-topmost", True)

        s = self.get_settings()
        main_f = ttk.Frame(self.panel, padding=14)
        main_f.pack(fill=tk.BOTH, expand=True)

        # 有効チェック
        self._var_enabled = tk.BooleanVar(value=s.get("tts_enabled", False))
        tk.Checkbutton(
            main_f, text="読み上げを有効にする",
            variable=self._var_enabled, font=("", 10, "bold")
        ).pack(anchor="w", pady=(0, 8))

        ttk.Separator(main_f).pack(fill=tk.X, pady=(0, 8))

        # バックエンド選択
        ttk.Label(main_f, text="バックエンド:", font=("", 9, "bold")).pack(anchor="w")
        self._var_backend = tk.StringVar(value=s.get("backend", "sapi"))
        bf = ttk.Frame(main_f)
        bf.pack(anchor="w", pady=(2, 8))
        for val, label in [("sapi", "Windows SAPI"), ("voicevox", "VOICEVOX"), ("coeiroink", "COEIROINK v2")]:
            tk.Radiobutton(bf, text=label, variable=self._var_backend,
                           value=val, command=self._on_backend_change).pack(side="left", padx=(0, 12))

        # バックエンド別パネル（Notebook）
        self._notebook = ttk.Notebook(main_f)
        self._notebook.pack(fill=tk.X, pady=(0, 8))

        self._build_sapi_tab(s)
        self._build_voicevox_tab(s)
        self._build_coeiroink_tab(s)

        self._on_backend_change()

        # 音声・デバイス一覧を自動取得（バックグラウンド）
        self._fetch_sapi_voices()
        self._fetch_sapi_devices()
        self._fetch_vv_devices()
        self._fetch_vv_speakers()    # スピーカー一覧 + 保存済み選択を復元
        self._fetch_ci_speakers()    # COEIROINKスピーカー一覧 + 保存済み選択を復元

        ttk.Separator(main_f).pack(fill=tk.X, pady=(0, 8))

        # 読み上げ対象
        ttk.Label(main_f, text="読み上げ対象:", font=("", 9, "bold")).pack(anchor="w")
        opt_f = ttk.Frame(main_f)
        opt_f.pack(anchor="w", pady=(2, 8))
        self._var_read_explain = tk.BooleanVar(value=s.get("read_explain", True))
        self._var_read_normal  = tk.BooleanVar(value=s.get("read_normal", True))
        self._var_read_topic   = tk.BooleanVar(value=s.get("read_topic", False))
        self._var_read_main    = tk.BooleanVar(value=s.get("read_main", True))
        tk.Checkbutton(opt_f, text="explainテロップ", variable=self._var_read_explain).pack(side="left", padx=(0, 12))
        tk.Checkbutton(opt_f, text="通常テロップ",    variable=self._var_read_normal).pack(side="left", padx=(0, 12))
        tk.Checkbutton(opt_f, text="TOPICも読む",     variable=self._var_read_topic).pack(side="left", padx=(0, 12))

        ttk.Separator(main_f).pack(fill=tk.X, pady=(0, 8))

        # フィルター
        ttk.Label(main_f, text="フィルター:", font=("", 9, "bold")).pack(anchor="w")

        scene_f = ttk.Frame(main_f)
        scene_f.pack(fill=tk.X, pady=(4, 4))
        scene_f.columnconfigure(1, weight=1)
        ttk.Label(scene_f, text="有効シーン:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self._var_active_scene = tk.StringVar(value=s.get("active_scene", ""))
        ttk.Entry(scene_f, textvariable=self._var_active_scene).grid(row=0, column=1, sticky="ew")
        ttk.Label(scene_f, text="（空=すべて）", foreground="gray").grid(row=0, column=2, padx=(6, 0))

        self._var_skip_system = tk.BooleanVar(value=s.get("skip_system_msg", False))
        tk.Checkbutton(
            main_f, text="システムメッセージをスキップ（接続中・切断など）",
            variable=self._var_skip_system
        ).pack(anchor="w", pady=(0, 8))

        ttk.Separator(main_f).pack(fill=tk.X, pady=(0, 8))

        # 保存ボタン
        tk.Button(
            main_f, text="保存して適用",
            bg="#007bff", fg="white", font=("", 10, "bold"),
            command=self._save_settings
        ).pack(fill=tk.X)

        self._lbl_status = ttk.Label(main_f, text="", foreground="gray")
        self._lbl_status.pack(pady=(6, 0))

    # ── SAPIタブ ──────────────────────────────
    def _build_sapi_tab(self, s):
        f = ttk.Frame(self._notebook, padding=10)
        self._notebook.add(f, text="SAPI")
        f.columnconfigure(1, weight=1)

        # 音声
        ttk.Label(f, text="音声:").grid(row=0, column=0, sticky="w", pady=3)
        self._var_sapi_voice = tk.StringVar(value=s.get("sapi_voice", ""))
        self._cb_sapi_voice = ttk.Combobox(f, textvariable=self._var_sapi_voice, state="readonly")
        self._cb_sapi_voice.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        tk.Button(f, text="一覧取得", command=self._fetch_sapi_voices).grid(row=0, column=2, padx=(4, 0))

        # デバイス
        ttk.Label(f, text="デバイス:").grid(row=1, column=0, sticky="w", pady=3)
        self._var_sapi_device = tk.StringVar(value=s.get("sapi_device", ""))
        self._cb_sapi_device = ttk.Combobox(f, textvariable=self._var_sapi_device, state="readonly")
        self._cb_sapi_device.grid(row=1, column=1, sticky="ew", padx=(8, 0))
        tk.Button(f, text="一覧取得", command=self._fetch_sapi_devices).grid(row=1, column=2, padx=(4, 0))

        # 速度・音量
        ttk.Label(f, text="速度 (-10〜10):").grid(row=2, column=0, sticky="w", pady=3)
        self._var_sapi_rate = tk.StringVar(value=str(s.get("sapi_rate", 0)))
        ttk.Spinbox(f, from_=-10, to=10, width=6,
                    textvariable=self._var_sapi_rate).grid(row=2, column=1, sticky="w", padx=(8, 0))

        ttk.Label(f, text="音量 (0〜100):").grid(row=3, column=0, sticky="w", pady=3)
        self._var_sapi_volume = tk.StringVar(value=str(s.get("sapi_volume", 100)))
        ttk.Spinbox(f, from_=0, to=100, width=6,
                    textvariable=self._var_sapi_volume).grid(row=3, column=1, sticky="w", padx=(8, 0))

        tk.Button(f, text="テスト", command=lambda: self._test_speak("sapi")).grid(
            row=4, column=1, sticky="e", pady=(8, 0))

    # ── VOICEVOXタブ ──────────────────────────
    def _build_voicevox_tab(self, s):
        f = ttk.Frame(self._notebook, padding=10)
        self._notebook.add(f, text="VOICEVOX")
        f.columnconfigure(1, weight=1)

        ttk.Label(f, text="URL:").grid(row=0, column=0, sticky="w", pady=3)
        self._var_vv_url = tk.StringVar(value=s.get("vv_url", "http://localhost:50021"))
        ttk.Entry(f, textvariable=self._var_vv_url).grid(row=0, column=1, sticky="ew", padx=(8, 0))

        ttk.Label(f, text="スピーカー:").grid(row=1, column=0, sticky="w", pady=3)
        self._var_vv_speaker = tk.StringVar(value=str(s.get("vv_speaker", 0)))
        self._cb_vv_speaker = ttk.Combobox(f, textvariable=self._var_vv_speaker, width=30)
        self._cb_vv_speaker.grid(row=1, column=1, sticky="ew", padx=(8, 0))
        tk.Button(f, text="一覧取得", command=self._fetch_vv_speakers).grid(
            row=1, column=2, padx=(4, 0))

        ttk.Label(f, text="速度:").grid(row=2, column=0, sticky="w", pady=3)
        self._var_vv_speed = tk.StringVar(value=str(s.get("vv_speed", 1.0)))
        ttk.Spinbox(f, from_=0.5, to=2.0, increment=0.1, width=6,
                    textvariable=self._var_vv_speed).grid(row=2, column=1, sticky="w", padx=(8, 0))

        ttk.Label(f, text="デバイス:").grid(row=3, column=0, sticky="w", pady=3)
        self._var_vv_device = tk.StringVar(value=s.get("vv_device", ""))
        self._cb_vv_device = ttk.Combobox(
            f, textvariable=self._var_vv_device,
            values=["（デフォルト）"], state="readonly")
        self._cb_vv_device.grid(row=3, column=1, sticky="ew", padx=(8, 0))
        tk.Button(f, text="一覧取得", command=self._fetch_vv_devices).grid(
            row=3, column=2, padx=(4, 0))

        ttk.Label(f, text="音量:").grid(row=4, column=0, sticky="w", pady=3)
        self._var_vv_volume = tk.IntVar(value=int(s.get("vv_volume", 0.6) * 100))
        vv_vol_f = ttk.Frame(f)
        vv_vol_f.grid(row=4, column=1, sticky="ew", padx=(8, 0))
        self._lbl_vv_volume = ttk.Label(vv_vol_f, text=f"{self._var_vv_volume.get()}%", width=5)
        self._lbl_vv_volume.pack(side="right")
        tk.Scale(vv_vol_f, variable=self._var_vv_volume, from_=0, to=200, resolution=5,
                 orient="horizontal", showvalue=False,
                 command=lambda v: self._lbl_vv_volume.config(text=f"{int(float(v))}%")
                 ).pack(side="left", fill="x", expand=True)

        tk.Button(f, text="テスト", command=lambda: self._test_speak("voicevox")).grid(
            row=5, column=1, sticky="e", pady=(8, 0))

    # ── COEIROINK v2 タブ ────────────────────
    def _build_coeiroink_tab(self, s):
        f = ttk.Frame(self._notebook, padding=10)
        self._notebook.add(f, text="COEIROINK v2")
        f.columnconfigure(1, weight=1)
        self._ci_item_to_params = {}  # display → (uuid, styleId)

        ttk.Label(f, text="URL:").grid(row=0, column=0, sticky="w", pady=3)
        self._var_ci_url = tk.StringVar(value=s.get("ci_url", "http://localhost:50032"))
        ttk.Entry(f, textvariable=self._var_ci_url).grid(row=0, column=1, sticky="ew", padx=(8, 0))

        ttk.Label(f, text="スピーカー:").grid(row=1, column=0, sticky="w", pady=3)
        self._var_ci_speaker_display = tk.StringVar()
        self._cb_ci_speaker = ttk.Combobox(f, textvariable=self._var_ci_speaker_display, width=30)
        self._cb_ci_speaker.grid(row=1, column=1, sticky="ew", padx=(8, 0))
        tk.Button(f, text="一覧取得", command=self._fetch_ci_speakers).grid(
            row=1, column=2, padx=(4, 0))

        ttk.Label(f, text="速度:").grid(row=2, column=0, sticky="w", pady=3)
        self._var_ci_speed = tk.StringVar(value=str(s.get("ci_speed", 1.0)))
        ttk.Spinbox(f, from_=0.5, to=2.0, increment=0.1, width=6,
                    textvariable=self._var_ci_speed).grid(row=2, column=1, sticky="w", padx=(8, 0))

        ttk.Label(f, text="デバイス:").grid(row=3, column=0, sticky="w", pady=3)
        self._var_ci_device = tk.StringVar(value=s.get("ci_device", ""))
        self._cb_ci_device = ttk.Combobox(
            f, textvariable=self._var_ci_device,
            values=["（デフォルト）"], state="readonly")
        self._cb_ci_device.grid(row=3, column=1, sticky="ew", padx=(8, 0))
        tk.Button(f, text="一覧取得", command=self._fetch_ci_devices).grid(
            row=3, column=2, padx=(4, 0))

        ttk.Label(f, text="音量:").grid(row=4, column=0, sticky="w", pady=3)
        self._var_ci_volume = tk.IntVar(value=int(s.get("ci_volume", 0.6) * 100))
        ci_vol_f = ttk.Frame(f)
        ci_vol_f.grid(row=4, column=1, sticky="ew", padx=(8, 0))
        self._lbl_ci_volume = ttk.Label(ci_vol_f, text=f"{self._var_ci_volume.get()}%", width=5)
        self._lbl_ci_volume.pack(side="right")
        tk.Scale(ci_vol_f, variable=self._var_ci_volume, from_=0, to=200, resolution=5,
                 orient="horizontal", showvalue=False,
                 command=lambda v: self._lbl_ci_volume.config(text=f"{int(float(v))}%")
                 ).pack(side="left", fill="x", expand=True)

        tk.Button(f, text="テスト", command=lambda: self._test_speak("coeiroink")).grid(
            row=5, column=1, sticky="e", pady=(8, 0))

    def _on_backend_change(self):
        tab_map = {"sapi": 0, "voicevox": 1, "coeiroink": 2}
        idx = tab_map.get(self._var_backend.get(), 0)
        self._notebook.select(idx)

    # ── 設定取得ヘルパー ─────────────────────
    def _fetch_sapi_voices(self):
        threading.Thread(target=self._fetch_sapi_voices_bg, daemon=True).start()

    def _fetch_sapi_voices_bg(self):
        voices = _get_sapi_voices()
        self.panel.after(0, lambda: self._cb_sapi_voice.config(values=voices))
        msg = f"✅ {len(voices)}件" if voices else "❌ 取得失敗"
        clr = "green" if voices else "red"
        self.panel.after(0, lambda: self._lbl_status.config(text=msg, foreground=clr))

    def _fetch_sapi_devices(self):
        threading.Thread(target=self._fetch_sapi_devices_bg, daemon=True).start()

    def _fetch_sapi_devices_bg(self):
        devices = _get_sapi_devices()
        self.panel.after(0, lambda: self._cb_sapi_device.config(values=["（デフォルト）"] + devices))
        msg = f"✅ {len(devices)}件" if devices else "❌ 取得失敗"
        clr = "green" if devices else "red"
        self.panel.after(0, lambda: self._lbl_status.config(text=msg, foreground=clr))

    def _get_playback_devices(self) -> list:
        """
        再生デバイス一覧を返す。
        sounddevice が使える場合はその名前を返す（_play_wav のマッチング対象と一致する）。
        使えない場合は SAPI 経由で取得（デバイス指定はできないが一覧は表示できる）。
        """
        sd, _ = _try_import_sounddevice()
        if sd is not None:
            try:
                return [d["name"] for d in sd.query_devices()
                        if d["max_output_channels"] > 0]
            except Exception:
                pass
        return _get_sapi_devices()

    def _fetch_vv_devices(self):
        threading.Thread(target=self._fetch_vv_devices_bg, daemon=True).start()

    def _fetch_vv_devices_bg(self):
        devices = self._get_playback_devices()
        values = ["（デフォルト）"] + devices
        self.panel.after(0, lambda: self._cb_vv_device.config(values=values))
        msg = f"✅ {len(devices)}件" if devices else "❌ 取得失敗"
        clr = "green" if devices else "red"
        self.panel.after(0, lambda: self._lbl_status.config(text=msg, foreground=clr))

    def _fetch_vv_speakers(self):
        threading.Thread(target=self._fetch_vv_speakers_bg, daemon=True).start()

    def _fetch_vv_speakers_bg(self):
        try:
            url = self._var_vv_url.get().rstrip("/")
            with urllib.request.urlopen(f"{url}/speakers", timeout=5) as resp:
                speakers = json.loads(resp.read())
            items = []
            for sp in speakers:
                for style in sp.get("styles", []):
                    items.append(f"{style['id']} : {sp['name']} ({style['name']})")

            def _update():
                self._cb_vv_speaker["values"] = items
                # 保存済み speaker ID に対応するエントリを選択
                saved_id = self._var_vv_speaker.get().split(":")[0].strip()
                for item in items:
                    if item.split(":")[0].strip() == saved_id:
                        self._var_vv_speaker.set(item)
                        break
                self._lbl_status.config(text=f"✅ {len(items)}件取得", foreground="green")

            self.panel.after(0, _update)
        except Exception as e:
            self.panel.after(0, lambda: self._lbl_status.config(
                text=f"❌ 取得失敗: {e}", foreground="red"))

    def _fetch_ci_speakers(self):
        threading.Thread(target=self._fetch_ci_speakers_bg, daemon=True).start()

    def _fetch_ci_speakers_bg(self):
        try:
            url = self._var_ci_url.get().rstrip("/")
            with urllib.request.urlopen(f"{url}/v1/speakers", timeout=5) as resp:
                speakers = json.loads(resp.read())
            items = []
            item_map = {}
            for sp in speakers:
                sp_uuid = sp.get("speakerUuid", "")
                sp_name = sp.get("speakerName", "")
                for style in sp.get("styles", []):
                    display = f"{sp_name} ({style.get('styleName', '')})"
                    items.append(display)
                    item_map[display] = (sp_uuid, int(style.get("styleId", 0)))

            def _update():
                self._ci_item_to_params = item_map
                self._cb_ci_speaker["values"] = items
                # 保存済み UUID + styleId に対応するエントリを選択
                saved_uuid  = self.get_settings().get("ci_speaker_uuid", "")
                saved_style = int(self.get_settings().get("ci_style_id", 0))
                for display, (uuid, sid) in item_map.items():
                    if uuid == saved_uuid and sid == saved_style:
                        self._var_ci_speaker_display.set(display)
                        break
                self._lbl_status.config(text=f"✅ {len(items)}件取得", foreground="green")

            self.panel.after(0, _update)
        except Exception as e:
            self.panel.after(0, lambda: self._lbl_status.config(
                text=f"❌ 取得失敗: {e}", foreground="red"))

    def _fetch_ci_devices(self):
        threading.Thread(target=self._fetch_ci_devices_bg, daemon=True).start()

    def _fetch_ci_devices_bg(self):
        devices = self._get_playback_devices()
        values = ["（デフォルト）"] + devices
        self.panel.after(0, lambda: self._cb_ci_device.config(values=values))
        msg = f"✅ {len(devices)}件" if devices else "❌ 取得失敗"
        clr = "green" if devices else "red"
        self.panel.after(0, lambda: self._lbl_status.config(text=msg, foreground=clr))

    def _test_speak(self, backend):
        s = self._collect_settings()
        s["backend"] = backend
        s["tts_enabled"] = True
        threading.Thread(
            target=self._speak_test_bg,
            args=("テロップ読み上げのテストです。", s),
            daemon=True
        ).start()

    def _speak_test_bg(self, text, s):
        try:
            if s["backend"] == "sapi":
                self._speak_sapi(text, s)
            elif s["backend"] == "voicevox":
                self._speak_voicevox(text, s)
            elif s["backend"] == "coeiroink":
                self._speak_coeiroink(text, s)
        except Exception as e:
            self.panel.after(0, lambda: self._lbl_status.config(
                text=f"❌ {e}", foreground="red"))

    def _collect_settings(self):
        s = self.get_settings()
        s["tts_enabled"]  = self._var_enabled.get()
        s["backend"]      = self._var_backend.get()
        s["sapi_voice"]   = self._var_sapi_voice.get()
        d = self._var_sapi_device.get()
        s["sapi_device"]  = "" if d == "（デフォルト）" else d
        s["sapi_rate"]    = int(self._var_sapi_rate.get())
        s["sapi_volume"]  = int(self._var_sapi_volume.get())
        s["vv_url"]       = self._var_vv_url.get().strip()
        # スピーカーは "ID : 名前" 形式 or 数値
        vv_sp = self._var_vv_speaker.get().split(":")[0].strip()
        try:
            s["vv_speaker"] = int(vv_sp)
        except ValueError:
            pass
        try:
            s["vv_speed"] = float(self._var_vv_speed.get())
        except ValueError:
            pass
        d2 = self._var_vv_device.get()
        s["vv_device"]    = "" if d2 == "（デフォルト）" else d2
        s["vv_volume"]    = self._var_vv_volume.get() / 100.0
        s["read_explain"]    = self._var_read_explain.get()
        s["read_normal"]     = self._var_read_normal.get()
        s["read_topic"]      = self._var_read_topic.get()
        # COEIROINK v2
        s["ci_url"]   = self._var_ci_url.get().strip()
        try:
            s["ci_speed"] = float(self._var_ci_speed.get())
        except ValueError:
            pass
        d3 = self._var_ci_device.get()
        s["ci_device"]  = "" if d3 == "（デフォルト）" else d3
        s["ci_volume"]  = self._var_ci_volume.get() / 100.0
        display = self._var_ci_speaker_display.get()
        if display and display in self._ci_item_to_params:
            uuid, style_id = self._ci_item_to_params[display]
            s["ci_speaker_uuid"] = uuid
            s["ci_style_id"]     = style_id
        s["active_scene"]    = self._var_active_scene.get().strip()
        s["skip_system_msg"] = self._var_skip_system.get()
        return s

    def _save_settings(self):
        s = self._collect_settings()
        self.save_settings(s)

        if self.is_running:
            # 読み上げ有効ならスレッドを起動
            if s["tts_enabled"]:
                self._start_threads()
            # シーンリスナーを設定に合わせて再起動
            self._stop_scene_listener()
            self._start_scene_listener(s)

        self._lbl_status.config(text="✅ 保存しました", foreground="green")
