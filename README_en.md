# TeloPon Plugin Collection

Unofficial plugins for [TeloPon](https://github.com/miyumiyu/TeloPon), an AI-powered streaming assistant for OBS.

> **日本語ドキュメント** → [README.md](README.md)

---

## Plugins

| Plugin | Type | Description |
|---|---|---|
| [OBS Screenshot Sender](#obs-screenshot-sender) | TOOL | Capture OBS sources and send to AI |
| [OBS Connection Status Badge](#obs-connection-status-badge) | TOOL | Show AI connection state in OBS text source |
| [Telop Reader (TTS)](#telop-reader-tts) | TOOL | Read telops aloud via speech synthesis |
| [OneComme Log Reader](#onecomme-log-reader) | BACKGROUND | Forward OneComme chat comments to AI |
| [Debug Log Viewer](#debug-log-viewer) | TOOL | Real-time TeloPon log viewer |

---

## Requirements

- **TeloPon v1.21b** or later
- Windows 10 / 11
- OBS Studio (for OBS-integrated features)
  - WebSocket Server must be enabled: OBS → Tools → WebSocket Server Settings
- VOICEVOX Engine (for VOICEVOX TTS backend)
- OneComme (for chat log plugin)

> Required libraries (such as `obsws-python` and `Pillow`) are bundled with TeloPon — no additional installation is needed.

---

## Installation

1. Download or `git clone` this repository
2. Copy the `.py` and `.json` files from `plugins/` into TeloPon's `plugins/` folder
   - **Copy `.json` files only on first install** to avoid overwriting your settings
3. Restart TeloPon

```
TeloPon-1.21b/
└── plugins/
    ├── obs_screenshot_sender.py   ← copy
    ├── obs_screenshot_sender.json ← first install only
    ├── obs_status_badge.py
    ├── obs_status_badge.json
    ├── telop_reader.py
    ├── telop_reader.json
    ├── onecomme_log.py
    ├── onecomme_log.json
    ├── log_viewer.py
    └── log_viewer.json
```

---

## Plugin Details

---

### OBS Screenshot Sender

**File:** `obs_screenshot_sender.py`

Captures OBS video sources and sends the image + a text prompt to the AI.
Supports up to 4 independent source slots with auto-send, scene filtering, and OBS WebSocket command control.

#### Features

- **Slots 1–4**: Configure source name, prompt text, and active scene per slot independently
- **Manual send**: "Capture & Send to AI" button for each slot
- **Auto-send**: Periodic automatic capture (configurable interval, auto-stop, duplicate skip)
- **Scene integration**
  - *Active scene*: Disable send button when not in the specified scene
  - *Auto ON/OFF*: Start auto-send when entering a scene, stop when leaving
- **OBS WebSocket commands**: Control from OBS scripts, hotkeys, or external tools

#### OBS WebSocket Commands

Send a `BroadcastCustomEvent` from OBS with the following JSON:

```json
{"command": "AI-SS-Sender", "action": "send", "slot": 1}
```

| action | Parameters | Behavior |
|---|---|---|
| `send` | `slot`: 1–4 | Immediately capture and send the specified slot |
| `set_source` | `slot`: 1–4, `name`: source name | Update and save the source name |
| `set_interval` | `seconds`: interval (min 10) | Update and save the auto-send interval |
| `auto` | `enabled`: true/false | Toggle auto-send on or off |

#### Settings File

`obs_screenshot_sender.json` (example):

```json
{
  "enabled": true,
  "source1": "Game Capture",
  "prompt1": "Please describe what's happening on screen!",
  "scene1": "",
  "auto_send1": true,
  "auto_send": true,
  "interval_sec": 120,
  "auto_stop_min": 60,
  "skip_duplicate": true
}
```

#### Dependencies

- `obsws-python` (bundled with TeloPon)
- `Pillow` (bundled with TeloPon)

---

### OBS Connection Status Badge

**File:** `obs_status_badge.py`

Displays the AI connection state in real time on an OBS GDI+ text source.
Monitors TeloPon's debug log to determine the current state.

#### States

| State | Default text | Color |
|---|---|---|
| Connected (idle) | `● Connected` | Green |
| Thinking (generating) | `○ Thinking` | Yellow |
| Disconnected | `● Disconnected` | Red |

- Display strings are customizable via the settings UI
- "Thinking" state is triggered by generation-start log entries and lasts 6 seconds

#### Settings

| Setting | Description |
|---|---|
| Text source name | Name of the OBS GDI+ text source (e.g. `AI_Status`) |
| Connected text | Default: `● 接続中` |
| Thinking text | Default: `○ 思考中` |
| Disconnected text | Default: `● 切断` |

#### Dependencies

- `obsws-python` (bundled with TeloPon)

---

### Telop Reader (TTS)

**File:** `telop_reader.py`

Polls `http://localhost:8000/data.json` (TeloPon OBS browser source) for telop changes
and reads them aloud using text-to-speech.

#### TTS Backends

| Backend | Notes |
|---|---|
| Windows SAPI | No pywin32 required — uses PowerShell COM interop |
| VOICEVOX | Requires VOICEVOX Engine running separately |
| COEIROINK v2 | Requires COEIROINK Engine running separately |

#### Features

- Select voice and audio output device from the UI (SAPI)
- Select speaker and output device from the UI (VOICEVOX / COEIROINK)
- Speed control for all backends
- Choose what to read: explain telops / normal telops / TOPIC field
- Active scene filter: only read aloud in the specified OBS scene
- Skip system messages: ignore "connecting", "disconnected" etc.

#### Device Selection Note

Output device selection depends on the libraries bundled with TeloPon.
If the required library is not included, audio plays on the system default output device.

---

### OneComme Log Reader

**File:** `onecomme_log.py`

Watches [OneComme](https://onecomme.com/) comment logs and forwards new messages to the AI in batches.

#### How it works

- Monitors `%APPDATA%\onecomme\comments\YYYY-MM-DD.log`
- Extracts only new comment IDs to avoid duplicate submissions
- Automatically switches to the new log file at midnight

#### Settings

| Setting | Description |
|---|---|
| Log folder | Path to OneComme logs (default: OneComme standard path) |
| Cooldown (sec) | Minimum interval between batch sends |

---

### Debug Log Viewer

**File:** `log_viewer.py`

Displays TeloPon's debug log (`telopon_debug.log`) in real time.
Useful for verifying plugin behavior and troubleshooting.

---

## Shared OBS WebSocket Configuration

Both `obs_screenshot_sender` and `obs_status_badge` share the OBS connection settings
stored in **`plugins/obs_capture.json`** — the same file used by TeloPon's built-in
"OBS Screen AI Commentary" plugin. No separate configuration needed.

```json
{
  "host": "127.0.0.1",
  "port": 4455,
  "password": "your_password"
}
```

---

## Multilingual Support

These plugins currently support **Japanese only**.

TeloPon's standard plugin convention uses an **inline `_L` dictionary + `_t()` function** pattern
for i18n — all translations live in the same `.py` file with no separate folders or files needed.
This is the correct approach for single-file plugins (folder-based i18n is used only for
multi-paragraph AI prompt files).

Contributions adding English, Russian, or Korean translations are welcome.

---

## License

MIT License

---

## Disclaimer

- These are unofficial third-party plugins for TeloPon
- Compatibility may break when TeloPon itself is updated
- Use at your own risk
