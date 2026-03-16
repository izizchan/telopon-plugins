# TeloPon プラグイン集

[TeloPon（公式）](https://github.com/miyumiyu/TeloPon) 向けの非公式プラグインです。

> **English documentation** → [README_en.md](README_en.md)

---

## 収録プラグイン

| プラグイン | 種別 | 説明 |
|---|---|---|
| [OBS画面AI送信](#obs画面ai送信) | TOOL | OBSソースをキャプチャしてAIに送信 |
| [OBS接続ステータス表示](#obs接続ステータス表示) | TOOL | AIの接続状態をOBSテキストソースに反映 |
| [テロップ読み上げ](#テロップ読み上げ) | TOOL | テロップを音声合成で読み上げ |
| [わんコメログ読み込み](#わんコメログ読み込み) | BACKGROUND | わんコメのコメントログをAIに送信 |
| [デバッグログビューア](#デバッグログビューア) | TOOL | TeloPonのログをリアルタイム表示 |

---

## 動作環境

- **TeloPon v1.22b** 以降
- Windows 10 / 11
- OBS Studio（OBSと連携する機能を使う場合）
  - WebSocket サーバー有効化が必要（OBS → ツール → WebSocket サーバー設定）
- VOICEVOX Engine（テロップ読み上げで VOICEVOX を使う場合）
- わんコメ OneComme（わんコメログ読み込みを使う場合）

> 必要なライブラリ（`obsws-python`・`Pillow` 等）は TeloPon 本体に同梱されているため、追加インストールは不要です。

---

## インストール

1. このリポジトリをダウンロード（または `git clone`）
2. `plugins/` フォルダ内の `.py` ファイルを TeloPon の `plugins/` フォルダにコピー
3. TeloPon を再起動

```
TeloPon-1.22b/
└── plugins/
    ├── obs_screenshot_sender.py   ← コピー
    ├── obs_status_badge.py
    ├── telop_reader.py
    ├── onecomme_log.py
    └── log_viewer.py
```

---

## プラグイン詳細

---

### OBS画面AI送信

**ファイル:** `obs_screenshot_sender.py`

OBS の映像ソースをキャプチャして画像とテキストをAIに送信します。
4つのソースを独立して管理でき、定期送信・シーン連動・OBS WebSocket コマンドに対応しています。

#### 機能

- **スロット 1〜4**：OBSソース名・送信テキスト・有効シーンをそれぞれ個別に設定
- **即時送信**：「キャプチャ & AIに送る」ボタンで手動送信
- **定期送信**：一定間隔で自動送信（間隔・自動停止時間・重複スキップを設定可）
- **シーン連動**
  - 「送信許可シーン名」：指定シーン以外ではボタンを無効化
  - 「自動ON/OFF」：指定シーンに入ったとき定期送信を ON、離れたとき OFF
- **OBS WebSocket コマンド**：OBS スクリプトやホットキーから操作可能

#### OBS WebSocket コマンド

OBS の `BroadcastCustomEvent` で以下の JSON を送信します。

```json
{"command": "AI-SS-Sender", "action": "send", "slot": 1}
```

| action | パラメータ | 動作 |
|---|---|---|
| `send` | `slot`: 1〜4 | 指定スロットを即時キャプチャ&送信 |
| `set_source` | `slot`: 1〜4, `name`: ソース名 | ソース名を変更して保存 |
| `set_interval` | `seconds`: 秒数（最小10） | 定期送信間隔を変更して保存 |
| `auto` | `enabled`: true/false | 定期送信をON/OFFに切り替え |

#### 必要ライブラリ

- `obsws-python`（TeloPon 同梱）
- `Pillow`（TeloPon 同梱）

---

### OBS接続ステータス表示

**ファイル:** `obs_status_badge.py`

AIの接続状態を OBS の GDI+テキストソースにリアルタイムで表示します。
TeloPon のデバッグログを監視してステータスを判定します。

#### 動作

| 状態 | デフォルト表示 | 色 |
|---|---|---|
| 接続中（待機） | `● 接続中` | 緑 |
| 思考中（生成中） | `○ 思考中` | 黄 |
| 切断 | `● 切断` | 赤 |

- 表示文字列は設定UIから変更可能
- 思考中の判定は生成開始ログを検知してから 6 秒間

#### 設定

| 項目 | 説明 |
|---|---|
| テキストソース名 | OBS の GDI+テキストソースの名前（例: `AI_Status`） |
| 接続中の文字列 | デフォルト: `● 接続中` |
| 思考中の文字列 | デフォルト: `○ 思考中` |
| 切断の文字列 | デフォルト: `● 切断` |

#### 必要ライブラリ

- `obsws-python`（TeloPon 同梱）

---

### テロップ読み上げ

**ファイル:** `telop_reader.py`

`http://localhost:8000/data.json`（TeloPon OBS ブラウザソース）をポーリングし、
テロップの変化を検知して音声合成で読み上げます。

#### 対応バックエンド

| バックエンド | 説明 |
|---|---|
| Windows SAPI | pywin32 不要。PowerShell経由でCOM操作 |
| VOICEVOX | VOICEVOX Engine が別途必要 |
| COEIROINK v2 | COEIROINK Engine が別途必要 |

#### 機能

- 音声・出力デバイスを UI から選択（SAPI）
- スピーカー・出力デバイスを UI から選択（VOICEVOX / COEIROINK）
- 再生速度・音量調整（0〜200%）
- 読み上げ対象を選択（explainテロップ / 通常テロップ / TOPIC）
- 有効シーン指定（指定シーン以外では読み上げしない）
- システムメッセージスキップ（「接続中」「切断」等を読まない）

#### デバイス指定について

出力デバイスの選択は TeloPon に同梱されているライブラリの状況によって動作が変わります。
同梱されていない場合はシステムの標準デバイスで再生されます。

---

### わんコメログ読み込み

**ファイル:** `onecomme_log.py`

[わんコメ（OneComme）](https://onecomme.com/) のコメントログを監視し、
新着コメントをまとめてAIに送信します。

#### 事前設定

わんコメ側でログ書き出しを有効にしてください。

> わんコメの設定 → その他 → コメントログを残す → **「ログをファイルとしても書き出し」をチェック**

#### 動作

- `%APPDATA%\onecomme\comments\YYYY-MM-DD.log` を監視
- 新しいコメントIDのみ抽出してバッチ送信
- 日付変更時に自動でファイルを切り替え

#### 設定

| 項目 | 説明 |
|---|---|
| ログフォルダ | ログ保存先（デフォルト: OneCommeの標準パス） |
| 送信クールダウン（秒） | 同一バッチの最低送信間隔 |

---

### デバッグログビューア

**ファイル:** `log_viewer.py`

TeloPon のデバッグログ（`telopon_debug.log`）をリアルタイムで表示するツールです。
動作確認・トラブルシュートに使います。

---

## OBS WebSocket 接続設定の共有

`obs_screenshot_sender`・`obs_status_badge` は **`plugins/obs_capture.json`** の接続設定を共有します。
「OBS画面AI実況」プラグイン（TeloPon 標準）と同じファイルを使うため、別途設定不要です。

```json
{
  "host": "127.0.0.1",
  "port": 4455,
  "password": "your_password"
}
```

---

## 多言語対応について

現在のプラグインは **日本語のみ**対応です。
TeloPon の標準プラグインが採用している **ファイル内 `_L` 辞書 + `_t()` 関数** パターンで多言語化できます（英語・ロシア語・韓国語）。
コントリビューション歓迎です。

---

## ライセンス

MIT License

---

## 免責事項

- 本プラグインは TeloPon の非公式プラグインです
- TeloPon 本体の仕様変更により動作しなくなる場合があります
- 使用は自己責任でお願いします
