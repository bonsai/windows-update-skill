---
name: windows-update-skill
description: >
  Windows Update の確認・適用・ポリシー管理・トラブルシュートを行うスキル。
  WUA COM (Microsoft.Update.Session) を使い、更新検索 / インストール / 履歴 /
  ポリシー / 一時停止 / キャッシュ掃除 / コンポーネント・リセット /
  自動インストールタスク登録まで PowerShell 1 本で扱う。
  WUA COM が使えない環境では Update Orchestrator レポートへ自動フォールバックする。
  出力はデフォルトで個人情報をマスキング (アカウント名 /  computername / GUID / メール / IP)。
  キーワード: "windows update", "windowsupdate", "windows アップデート", "更新プログラム",
  "windows update", "KB", "kb適用", "アップデート", "update", "windows update 確認",
  "windows update 実行", "windows update 適用", "windows update エラー", "更新エラー",
  "0x80240022", "0x8024402c", "windows update できない", "更新できない",
  "windows update 止まる", "windows update リセット", "windows update 修復",
  "再起動 要求", "reboot pending", "windows update 一時停止", "windows update 遅い",
  "更新の一時停止", "windows update ポリシー", "wsus", "更新ポリシー", "ターゲットリリース",
  "windows update キャッシュ", "softwaredistribution", "delivery optimization",
  "windows update 履歴", "language pack", "windows update タスク", "自動更新 スケジュール"
  CLI: `~/.opencode/skills/windows-update-skill/wupm.ps1`
  MCP: `~/.opencode/mcp-servers/windows-update-mcp/windows-update-mcp.ps1`
---

# Windows Update 管理スキル

Windows Update を PowerShell 7 の CLI (`wupm.ps1`) から操作するためのスキル。
外部モジュール (PSWindowsUpdate) に依存せず、組み込みの WUA COM のみで完結する。

## 基本コマンド

```powershell
$p = "$HOME\.opencode\skills\windows-update-skill\wupm.ps1"

& $p -Action status            # 総合ステータス (OS / 再起動フラグ / サービス / ポリシー)
& $p -Action search            # 未適用更新の一覧
& $p -Action install -DryRun   # 適用対象を確認 (何もインストールしない)
& $p -Action install -NoRestart# サイレント適用 (自動再起動なし)
& $p -Action history -Count 30 # 更新履歴
& $p -Action report            # Update Orchestrator レポート (検索不可時の代替)
& $p -Action selftest          # 環境診断 + 個人情報マスキングの単体テスト
```

WSL からは PowerShell を直接叩く:

```bash
wsl -e pwsh -NoProfile -Command "& '$USERPROFILE/.opencode/skills/windows-update-skill/wupm.ps1' -Action status"
```

## アクション一覧

| Action | 説明 | 管理者 |
|--------|------|--------|
| `status` | OS / ビルド / 再起動保留フラグ / 関連サービス / ポリシー / 最近イベント。`-Deep` で未適用件数も | 不要 |
| `search` | 未適用更新一覧。`-Type Software\|Driver\|All`, `-Title <regex>`, `-MaxResults` | 不要 |
| `install` | ダウンロード+インストール。`-DryRun`, `-Title`, `-NoRestart`, `-Orchestrate` | **必要** |
| `history` | 更新履歴 (COM QueryHistory)。`-Count`, `-Search <regex>` | 不要 |
| `hotfix` | インストール済みホットフィクス (CIM)。`InstalledBy` は個人情報のため出さない | 不要 |
| `policy` | UpdatePolicy\Settings / PolicyState / Group Policy / ServiceManager を表示 | 不要 |
| `set-target` | ターゲットリリースを固定 (`-TargetRelease 23H2` → 23000) | **必要** |
| `clear-target` | ターゲットリリースを解除し最新へ追従 | **必要** |
| `pause` | 更新を一時停止 (`-PauseDays 7`, `-DeferType Feature\|Quality`) | **必要** |
| `resume` | 一時停止を解除 | **必要** |
| `cache` | キャッシュ容量を表示。`-Clean` で削除 | `-Clean` は**必要** |
| `log` | WU イベントログ。`-Days`, `-Level Info\|Warning\|Error\|All`, `-Count` | 不要 |
| `report` | `ReportingEvents.log` を解析 (タブ区切り構造化ログ) | 不要 |
| `reset` | WU コンポーネントをリセット。`-ReRegister`, `-RunDism`, `-ClearRenameOps` | **必要** |
| `schedule` | 自動インストールタスク。`-TaskState install\|remove\|status\|run`, `-At 03:00` | install/remove は**必要** |
| `selftest` | 環境診断 + マスキング単体テスト | 不要 |

### 共通フラグ

| フラグ | 説明 |
|--------|------|
| `-Json` | 機械可読な JSON 出力 |
| `-LogFile <path>` | 出力内容をファイルに追記 (タスク実行時のログ用) |
| `-NoRedact` | 個人情報マスキングを無効化 (ローカル調査のみ) |
| `-Deep` | `status` で未適用更新の検索まで実施 |

## ワークフロー

### 1. updates 確認 → 適用

```powershell
& $p -Action selftest                    # まず環境が健全か
& $p -Action status                      # 再起動保留を確認
& $p -Action search                      # 未適用一覧
& $p -Action install -DryRun             # 適用対象を確認 (何もインストールしない)
& $p -Action install -NoRestart          # 適用 (再起動は手動)
& $p -Action history -Count 10           # 結果確認
```

### 2. 更新エラー (`0x8024402C` など) からの復旧

```powershell
& $p -Action report -Count 20            # Orchestrator が何を見たか
& $p -Action cache -Clean                # キャッシュを飛ばす
& $p -Action reset -ReRegister           # コンポーネント再登録
& $p -Action reset -RunDism              # イメージ修復 (重い)
```

`reset` は SoftwareDistribution を `SoftwareDistribution.old-<timestamp>` へ退避する。
戻す場合はリネームし直す。

### 3. 大きな更新を保留したい

```powershell
& $p -Action pause -PauseDays 14 -DeferType Feature
& $p -Action set-target -TargetRelease 25H2
& $p -Action policy                      # 反映確認
& $p -Action resume                      # 解除
```

### 4. 毎晩の自動適用

```powershell
& $p -Action schedule -TaskState install -At 03:00   # SYSTEM 権限・再起動なし
& $p -Action schedule -TaskState status
& $p -Action schedule -TaskState run                 # 今すぐ1回
Get-Content "$env:ProgramData\wupm\auto-update.log" -Tail 50
```

## 個人情報のマスキング (デフォルト ON)

出力を Notes / wiki / LLM に貼っても個人を特定できないように既定でマスキングする。

対象:

- Windows アカウント名 (`$env:USERNAME`, `[Environment]::UserName`)
- computername / ドメイン名
- GUID (CorrelationId / EventId / UpdateId) → 先頭 4 文字のみ残す
- メールアドレス → `***@***`
- IPv4 アドレス → `x.x.x.x`
- `InstalledBy` (ホットフィクスの適用者) → カラム自体を出さない
- `ClientApplicationID` → 外部アプリの識別子なので `yes`/`""` に置換

`-NoRedact` で無効化できるが、外部に貼る用途では使わないこと。
検証は `-Action selftest` の "Redaction unit test" で行う。

## この環境での既知の制約 (実測)

| 事象 | 原因 | 対処 |
|------|------|------|
| `0x8024402C` WU_E_MAX_CAPACITY_REACHED | 一時的 (検索容量が埋まっている時) | 時間を置いて再実行。恒久時は `report` / `-Orchestrate` にフォールバック済 |
| `0x80240032` WU_E_INSTALL_NOT_ALLOWED | criteria に `IsBrowseOnly=0` を併用すると発生 | criteria から `IsBrowseOnly` を外す (コードで回避済) |
| `IUpdateColl.RemoveAt` が `0x80240037` | 検索結果コレクションは読み取り専用 | `New-Object -ComObject Microsoft.Update.UpdateColl` を作り `Add()` する (実装済) |
| `AcceptEula()` が False を返す | 戻り値は「既に受け入れ済みか」= 失敗ではない | 戻り値を無視し副作用のみ利用 (実装済) |
| `IUpdate.Type` が文字列でなく整数 | COM 経由で `0=Other 1=Software 2=Driver 3=Firmware` | 数値でマップ (実装済) |
| `Microsoft.Update.ServiceManager` に `GetIsPauseUpdatesEffective` が無い | COM 遅延バインディングで pause 系メンバーを公開しない (Get-Member で確認済み) | 一時停止はレジストリ (`UX\Settings`) 経由に実装変更 (実装済) |
| `ReportingEvents.log` の日時が parse 失敗 | 形式が `02:52:34:152+0900` (オフセット直前のコロが不正) | 正規表現でコロ除去後にパース (実装済) |

使える criteria は実質次の 3 種のみ:

```
IsInstalled=0
IsInstalled=0 and IsHidden=0
IsInstalled=0 and Type='Software'      # Type='Driver' も可
```

`IsBrowseOnly` は併用不可。

## MCP サーバ

`~/.opencode/mcp-servers/windows-update-mcp/windows-update-mcp.ps1` に 13 個の MCP ツールを公開している。
`~/.config/opencode/opencode.jsonc` の `mcp.windows-update-mcp` に登録済み。

| ツール | 種別 | 備考 |
|--------|------|------|
| `windows_update_status` | read | `deep` で未適用件数まで |
| `windows_update_search` | read | `type` / `title` / `max_results` |
| `windows_update_history` | read | `count` / `search` |
| `windows_update_hotfix` | read | InstalledBy は出力しない |
| `windows_update_report` | read | Orchestrator レポート |
| `windows_update_policy` | read | ポリシー + 一時停止実効値 |
| `windows_update_logs` | read | `days` / `level` / `count` |
| `windows_update_cache` | read+write | `clean` は `confirm` 必須 |
| `windows_update_install` | write | `dry_run` 既定 true、`confirm` 必須 |
| `windows_update_pause` | write | `resume` フラグで解除。`dry_run` 既定 true |
| `windows_update_reset` | write | 破壊的。`dry_run` 既定 true、`confirm` 必須 |
| `windows_update_schedule` | read+write | `task_state` install/remove/status/run |
| `windows_update_selftest` | read | 環境診断 + マスキングテスト |

安全設計: 破壊的ツールは `dry_run` 既定 true、`dry_run=false` には `confirm=true` が必須。
引数は `ProcessStartInfo.ArgumentList` で渡し、スペースを含むパスでもクォート崩れしない。

## 作業ログ (Recap)

### 完了・検証済み

- `wupm.ps1` 実装 (16 アクション)。構文チェック済み
- 実行確認済み: `status` / `status -Deep` / `search` / `search -Type Software` /
  `install -DryRun` (`-Title ASUSTeK` で 2 件に絞ることを確認) / `history` / `hotfix` /
  `report` / `cache` / `log` / `policy` / `schedule -TaskState status` / `selftest`
- `pause` → `resume` の実往復テスト済み (レジストリ書き込み → 確認 → 完全復元、残存なし)
- `-Json` 出力確認済み
- 個人情報マスキング実装 + 単体テスト 6 件すべて PASS
- ロケール表記バグ修正 (`.` 区切りに統一、`Format-Num` / `Format-DateTime` ヘルパー追加)
- `ReportingEvents.log` パーサ実装 (タブ区切り 13 列、列位置を実測で特定)
- MCP サーバ実装 + `opencode.jsonc` 登録 + initialize / tools/list / tools/call 動作確認

### 解決したバグ

| 症状 | 原因 | 修正 |
|------|------|------|
| `1,41,150.0 B` | ロケールがカンマ小数点 | `Format-Num` で InvariantCulture 固定 |
| `ToString` オーバーロード例外 | ログの日時が文字列のまま残っていた | `ConvertFrom-ReportTime` で正規化してからパース |
| `更新リストから項目を除外できません: 0x80244037` | `IUpdateColl.RemoveAt` が使えない | `New-Object -ComObject Microsoft.Update.UpdateColl` を作り `Add()` |
| 対象が 0 件になる | `AcceptEula()` の戻り値を失敗扱いしていた | 戻り値を無視 (副作用のみ利用) |
| ドライバが Software と表示 | `IUpdate.Type` が整数 | 数値でマップ |
| `0x80240032` | criteria の `IsBrowseOnly=0` | criteria から除外 |
| pause/resume が動作しない | ServiceManager COM が pause メンバーを公開しない | レジストリ (`UX\Settings`) 方式へ変更 |
| `status` のイベント日時が `03-10-2026` | 既定のロケール書式 | `Format-DateTime` に統一 |
| 個人情報が含まれ得る | 出力に無加工だった | 既定でマスキング、`-NoRedact` で解除可 |

### 未検証 (副作用があるため未実行)

- `install` (実適用) / `reset` / `cache -Clean` / `set-target` / `clear-target` /
  `schedule install|remove|run`

いずれも `-DryRun` で対象を確認してから実行すること。

### 既知の状態

- ソフトウェア更新の未適用は 0 件 (ドライバのみ 6 件が未適用)
- `IsOOBEInProgress = 1` が設定されている (検索拒否の候補要因だが、0 にしても解決せず)
- 履歴に `Windows 11, version 26H2` の `Aborted` (`0x80240034` WU_E_UH_NEEDANOTHER_SEARCH) が繰り返し
- 実効値側の `PausedFeatureStatus` は pause 設定直後にまだ反映されない (数分かかる。警告で通知)

## トラブルシュート

| 症状 | 対処 |
|------|------|
| `管理者権限が必要です` で止まる | 管理者ターミナルで実行 (`status` 系は管理者不要) |
| 検索が遅い (20〜40 秒) | 仕様。`search` は Microsoft Update へ問い合わせる |
| `レポートを読み取れません` | `ReportingEvents.log` が未生成。`UsoClient StartScan` を先に実行 |
| ログが文字化け | `chcp 65001` または `pwsh -NoProfile` で実行。ファイルは UTF-8 |