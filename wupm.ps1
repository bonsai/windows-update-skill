# wupm.ps1 - Windows Update 管理 CLI (Windows 11 / PowerShell 7)
[CmdletBinding()]
param(
    [ValidateSet(
        "status", "search", "install", "history", "hotfix",
        "policy", "set-target", "clear-target", "pause", "resume",
        "cache", "log", "report", "reset", "schedule", "selftest"
    )]
    [string]$Action = "status",

    # search / install
    [ValidateSet("Software", "Driver", "All")]
    [string]$Type = "All",
    [string]$Title,
    [int]$MaxResults = 80,
    [switch]$DryRun,
    [switch]$NoRestart,
    [switch]$IncludeHidden,
    [switch]$IncludeDrivers,
    [switch]$OfflineSearch,
    [switch]$Orchestrate,

    # history / log
    [int]$Count = 25,
    [string]$Search,
    [int]$Days = 7,
    [ValidateSet("Info", "Warning", "Error", "All")]
    [string]$Level = "Warning",

    # cache / reset
    [switch]$Clean,
    [switch]$ReRegister,
    [switch]$ClearRenameOps,
    [switch]$RunDism,

    # schedule
    [ValidateSet("install", "remove", "status", "run")]
    [string]$TaskState = "status",
    [string]$At = "03:00",

    # policy
    [string]$TargetRelease,
    [int]$PauseDays = 7,
    [ValidateSet("Feature", "Quality")]
    [string]$DeferType = "Feature",
    [int]$DeferDays = 30,

    [switch]$Deep,
    [switch]$NoRedact,
    [switch]$Json,
    [string]$LogFile
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$OutputEncoding = [Console]::OutputEncoding
$ProgressPreference = "SilentlyContinue"

$ScriptPath = $PSCommandPath
$TaskName   = "WindowsUpdate-Auto"
$AppId      = "wupm-skill"
$HKLMUpdatePolicy = "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UpdatePolicy\Settings"
$HKLMPolicyState  = "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UpdatePolicy\PolicyState"
$HKLMWUPolicy     = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
$HKLMAUPolicy     = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
$HKLMUXSettings   = "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"
$WULog            = "Microsoft-Windows-WindowsUpdateClient/Operational"

function Get-PauseState {
    <# Microsoft.Update.ServiceManager は PowerShell の COM 遅延バインディングで
       GetIsPauseUpdatesEffective / PauseUpdatesExpiryTime を公開しない。
       そのためレジストリ (UpdatePolicy\Settings = 実効値, UX\Settings = 設定値) から読む。 #>
    $st = Get-RegProps $HKLMUpdatePolicy
    $ux = Get-RegProps $HKLMUXSettings
    $state = [ordered]@{
        PausedFeatureStatus = $(if ($st -and $st.Contains("PausedFeatureStatus")) { $st["PausedFeatureStatus"] } else { $null })
        PausedQualityStatus = $(if ($st -and $st.Contains("PausedQualityStatus")) { $st["PausedQualityStatus"] } else { $null })
        PausedFeatureExpiry = $(if ($st -and $st.Contains("PausedFeatureExpiryTime")) { $st["PausedFeatureExpiryTime"] } else { $null })
        PausedQualityExpiry = $(if ($st -and $st.Contains("PausedQualityExpiryTime")) { $st["PausedQualityExpiryTime"] } else { $null })
        ConfiguredExpiry    = $(if ($ux -and $ux.Contains("PauseUpdatesExpiryTime")) { $ux["PauseUpdatesExpiryTime"] } else { $null })
        ConfiguredStart     = $(if ($ux -and $ux.Contains("PauseUpdatesStartTime")) { $ux["PauseUpdatesStartTime"] } else { $null })
    }
    $state["IsPaused"] = (($state.PausedFeatureStatus -eq 1) -or ($state.PausedQualityStatus -eq 1))
    return $state
}

function Write-PauseRegistry {
    param([int]$Days, [ValidateSet("Feature", "Quality", "Both")][string]$DeferType, [switch]$Clear)
    $expiry = (Get-Date).AddDays($Days).ToString("yyyy-MM-dd")
    $start = (Get-Date).ToString("yyyy-MM-dd")
    if (-not (Test-Path $HKLMUXSettings)) { New-Item -Path $HKLMUXSettings -Force | Out-Null }

    if ($Clear) {
        foreach ($n in @("PauseUpdatesExpiryTime", "PauseUpdatesStartTime",
                         "PauseFeatureUpdatesStartTime", "PauseFeatureUpdatesExpiryTime",
                         "PauseQualityUpdatesStartTime", "PauseQualityUpdatesExpiryTime")) {
            Remove-ItemProperty -Path $HKLMUXSettings -Name $n -ErrorAction SilentlyContinue
        }
        return $null
    }
    New-ItemProperty -Path $HKLMUXSettings -Name PauseUpdatesStartTime  -Value $start  -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $HKLMUXSettings -Name PauseUpdatesExpiryTime -Value $expiry -PropertyType String -Force | Out-Null
    if ($DeferType -in @("Feature", "Both")) {
        New-ItemProperty -Path $HKLMUXSettings -Name PauseFeatureUpdatesStartTime  -Value $start  -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $HKLMUXSettings -Name PauseFeatureUpdatesExpiryTime -Value $expiry -PropertyType String -Force | Out-Null
    }
    if ($DeferType -in @("Quality", "Both")) {
        New-ItemProperty -Path $HKLMUXSettings -Name PauseQualityUpdatesStartTime  -Value $start  -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $HKLMUXSettings -Name PauseQualityUpdatesExpiryTime -Value $expiry -PropertyType String -Force | Out-Null
    }
    return $expiry
}

function Invoke-UsoRefresh {
    foreach ($s in @("RefreshSettings", "Pause")) {
        try { Start-Process "$env:SystemRoot\System32\UsoClient.exe" -ArgumentList $s -WindowStyle Hidden -Wait -ErrorAction Stop } catch { }
    }
}

# ---------------------------------------------------------------- privacy

# 出力から個人情報を除去する。默认 ON。-NoRedact で無効化 (ローカル調査のみ)。
$script:Redact = -not $NoRedact
$script:Identity = New-Object System.Collections.Generic.List[string]
foreach ($candidate in @(
    $env:USERNAME, $env:COMPUTERNAME, $env:USERDOMAIN, "$env:USERDOMAIN\$env:USERNAME",
    [Environment]::UserName, [Environment]::MachineName, [Environment]::UserDomainName
)) {
    if ($candidate -and $candidate.Length -ge 3 -and -not $script:Identity.Contains($candidate)) {
        $script:Identity.Add($candidate) | Out-Null
    }
}

function Protect-PII($Value) {
    if (-not $script:Redact) { return $Value }
    if ($null -eq $Value) { return $Value }
    $t = $Value -as [string]
    if ($null -eq $t) { return $Value }
    if ($t -eq '') { return $t }

    # GUID (CorrelationId / EventId / UpdateId)
    $t = [regex]::Replace($t, '(?i)(?<![0-9a-f-])[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}(?![0-9a-f-])',
        { param($m) $m.Value.Substring(0, 4) + '-****-****-****-************' })

    # メールアドレス
    $t = [regex]::Replace($t, '(?i)\b[\w.\-+]+@(?:[\w\-]+\.)+[a-z]{2,}\b', '***@***')

    # IPv4 アドレス
    $t = [regex]::Replace($t, '\b(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})\b', '$1.x.x.x')

    # ローカルアカウント名 /  computername / ドメイン (リテラル一致)
    foreach ($id in $script:Identity) {
        if ([string]::IsNullOrWhiteSpace($id) -or $id.Length -lt 3) { continue }
        $t = [regex]::Replace($t, "(?i)(?<![A-Za-z0-9_\-])" + [regex]::Escape($id) + "(?![A-Za-z0-9_\-])", '***')
    }
    return $t
}

function Protect-Object($Node) {
    if (-not $script:Redact -or $null -eq $Node) { return $Node }
    if ($Node -is [string]) { return Protect-PII $Node }
    if ($Node -is [System.Collections.IDictionary]) {
        foreach ($k in @($Node.Keys)) { $Node[$k] = Protect-Object $Node[$k] }
        return $Node
    }
    if ($Node -is [psobject] -and $Node.PSObject.Properties.Name -contains 'PSObject') {
        foreach ($p in @($Node.PSObject.Properties)) {
            if ($p.Name -match '^PS') { continue }
            try { $p.Value = Protect-Object $p.Value } catch { }
        }
        return $Node
    }
    if ($Node -is [System.Collections.IEnumerable]) {
        foreach ($item in $Node) { [void](Protect-Object $item) }
        return $Node
    }
    return $Node
}

# ---------------------------------------------------------------- infra

function Write-Section($Title) {
    Write-Host ""
    Write-Host "===== $(Protect-PII $Title) =====" -ForegroundColor Cyan
}

function Write-Info($Message)  { Write-Host (Protect-PII $Message) }
function Write-Ok($Message)    { Write-Host (Protect-PII $Message) -ForegroundColor Green }
function Write-Warn2($Message) { Write-Host (Protect-PII $Message) -ForegroundColor Yellow }
function Write-Err2($Message)  { Write-Host (Protect-PII $Message) -ForegroundColor Red }

function Test-Admin {
    return ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Require-Admin($What) {
    if (-not (Test-Admin)) {
        Write-Err2 "管理者権限が必要です: $What"
        Write-Info "管理者ターミナルで再実行するか: pwsh -Command ""Start-Process pwsh -Verb RunAs -ArgumentList '-NoProfile','-File','$ScriptPath','-Action','$Action'"""
        exit 2
    }
}

function Get-RegProps($Path) {
    if (-not (Test-Path $Path)) { return $null }
    $p = Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue
    if (-not $p) { return $null }
    $h = [ordered]@{}
    foreach ($prop in $p.PSObject.Properties) {
        if ($prop.Name -notmatch '^PS') { $h[$prop.Name] = $prop.Value }
    }
    return $h
}

$script:Culture = [System.Globalization.CultureInfo]::InvariantCulture

function Format-Num($Value, [string]$Format = "0.0") {
    return ([double]$Value).ToString($Format, $script:Culture)
}

function ConvertTo-HumanSize([double]$Bytes) {
    if ($Bytes -le 0) { return "0 B" }
    $u = @("B", "KB", "MB", "GB", "TB")
    $i = [math]::Min([math]::Floor([math]::Log($Bytes) / [math]::Log(1MB)), $u.Count - 1)
    return ("{0} {1}" -f (Format-Num ($Bytes / [math]::Pow(1MB, $i))), $u[$i])
}

function Get-DirSize([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return 0 }
    try {
        $sum = (Get-ChildItem -LiteralPath $Path -Recurse -Force -File -ErrorAction SilentlyContinue |
                Measure-Object -Property Length -Sum).Sum
        if ($null -eq $sum) { return 0 }
        return [double]$sum
    } catch { return 0 }
}

$script:OutLines = New-Object System.Collections.Generic.List[string]
function Emit-Line($Message) {
    Write-Host $Message
    $script:OutLines.Add($Message) | Out-Null
}

$script:Result = [ordered]@{ action = $Action; summary = [ordered]@{}; items = @() }
function Set-Summary($Key, $Value) { $script:Result.summary[$Key] = $Value }

function Complete-Result($Columns) {
    $script:Result.items = @($script:Result.items)
    if ($LogFile) {
        try {
            $dir = Split-Path -Parent $LogFile
            if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            $header = "--- $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Action] ---"
            Add-Content -LiteralPath $LogFile -Value (@($header) + $script:OutLines) -Encoding utf8
        } catch { Write-Warn2 "ログ書き込み失敗: $_" }
    }
    if ($Json) { Protect-Object $script:Result | ConvertTo-Json -Depth 6 }
    exit 0
}

function New-Table($Rows, $Columns) {
    if (-not $Rows -or $Rows.Count -eq 0) {
        Emit-Line "  (該当なし)"
        return
    }
    $table = $Rows | Select-Object $Columns
    $table | Format-Table -AutoSize | Out-String -Width 200 | ForEach-Object { Emit-Line $_.TrimEnd() }
}

function Write-OrchestratorFallback {
    Write-Section "Update Orchestrator レポート (フォールバック)"
    $rep = Get-OrchestratorReport
    if (-not $rep -or @($rep).Count -eq 0) {
        Emit-Line "  レポートを読み取れません。"
        Set-Summary "Fallback" "unavailable"
        return
    }
    $detected = Get-LastDetectedCount $rep
    $lastDet = @($rep | Where-Object { $_.Event -eq "AGENT_DETECTION_FINISHED" } | Sort-Object Time -Descending | Select-Object -First 1)
    Emit-Line ("  最終検出: {0}" -f (Format-DateTime $lastDet[0].Time))
    Emit-Line ("  検出された未適用更新: {0} 件" -f $(if ($null -ne $detected) { $detected } else { "不明" }))
    Set-Summary "LastDetectedCount" $detected
    Set-Summary "Fallback" "ReportingEvents.log"

    $act = @($rep | Where-Object { $_.Event -match "AGENT_(INSTALLING|DOWNLOAD)" } | Sort-Object Time -Descending | Select-Object -First 8)
    if ($act.Count) {
        Write-Info "  最近のインストール / ダウンロード:"
        $act | ForEach-Object {
            Emit-Line ("    [{0}] {1}" -f (Format-DateTime $_.Time "MM-dd HH:mm:ss"), $_.Message)
        }
    }
    Write-Warn2 "WUA COM 検索が利用できません。更新の適用は設定画面または:"
    Write-Info "  UsoClient StartDownload / StartInstall"
    Write-Info "  設定 > Windows Update > 更新チェック"
}

function Invoke-OrchestrateInstall {
    <# WUA COM が使えない環境向け。Update Orchestrator (UsoClient) に処理を要求する。 #>
    $uso = "$env:SystemRoot\System32\UsoClient.exe"
    if (-not (Test-Path $uso)) { Write-Err2 "UsoClient.exe が見つかりません"; return $false }
    $steps = @("RefreshSettings", "StartScan")
    if ($DryRun) {
        Write-Info "[DryRun] 要求する処理: $($steps -join ' -> ')"
        return $true
    }
    foreach ($s in $steps) {
        Write-Info "UsoClient $s ..."
        try {
            Start-Process $uso -ArgumentList $s -WindowStyle Hidden -Wait
            Start-Sleep -Seconds 2
            Write-Ok "  完了: $s"
        } catch { Write-Warn2 "  失敗: $s ($_)" }
    }
    Write-Warn2 "ダウンロード・インストールはバックグラウンドで継続します。"
    Write-Info "進捗確認: wupm.ps1 -Action report -Count 10"
    try { Start-Process $uso -ArgumentList "StartDownload" -WindowStyle Hidden } catch { }
    if (-not $NoRestart) {
        Write-Info "再起動が要求された場合は -Action status で RebootPending を確認してください。"
    }
    return $true
}

# ---------------------------------------------------------------- COM: search / install

function New-UpdateSession {
    return New-Object -ComObject Microsoft.Update.Session
}

function Invoke-UpdateSearch {
    param([string]$Criteria, [switch]$Offline)
    $session = New-UpdateSession
    $searcher = $session.CreateUpdateSearcher()
    if ($Offline) {
        $searcher.Online = $false
    } else {
        try { $searcher.Online = $true } catch { }
    }
    Write-Warn2 "検索中 (Microsoft Update へ問い合わせ。数十秒かかることがあります)..."
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try {
        $result = $searcher.Search($Criteria)
    } catch {
        $sw.Stop()
        $hr = [int]$_.Exception.HResult
        $session = $null
        if ($hr -eq 0x8024402C) {
            Write-Warn2 ("検索が拒否されました: 0x{0:X8} (WU_E_MAX_CAPACITY_REACHED) / {1} 秒" -f $hr, (Format-Num $sw.Elapsed.TotalSeconds))
            Write-Warn2 "この環境では WUA COM 検索が利用できません。Update Orchestrator レポートへフォールバックします。"
            return $null
        }
        throw
    }
    $sw.Stop()
    Emit-Line ("  検索完了: {0} 件 / {1} 秒" -f $result.Updates.Count, (Format-Num $sw.Elapsed.TotalSeconds))
    return [pscustomobject]@{ Session = $session; Searcher = $searcher; Result = $result }
}

# ---------------------------------------------------------- Update Orchestrator report

function Format-DateTime($Value, [string]$Format = "yyyy-MM-dd HH:mm:ss") {
    if ($null -eq $Value) { return "-" }
    if ($Value -is [datetime]) { return $Value.ToString($Format, $script:Culture) }
    if ($Value -is [datetimeoffset]) { return $Value.ToString($Format, $script:Culture) }
    return [string]$Value
}

function ConvertFrom-ReportTime([string]$Raw) {
    # 例: 2026-10-01 02:52:34:152+0900 (オフセット直前のコロは不正なので除去)
    if (-not $Raw) { return $null }
    $norm = $Raw.Trim() -replace ':(?=[+-]\d{4}$)', ''
    try { return [DateTimeOffset]::Parse($norm).LocalDateTime } catch { }
    $m = [regex]::Match($norm, '^(?<d>\d{4}-\d{2}-\d{2}) (?<t>\d{2}:\d{2}:\d{2})')
    if ($m.Success) {
        try { return [datetime]::ParseExact(($m.Groups['d'].Value + " " + $m.Groups['t'].Value), "yyyy-MM-dd HH:mm:ss", $script:Culture) } catch { }
    }
    return $null
}

function Get-OrchestratorReport {
    <# ReportingEvents.log はタブ区切り構造化ログ。
       列: 0=EventId 1=Time 2=? 3="<id> [EVENT_NAME]" 4=? 5=UpdateGuid 6=? 7=HResult
           8=Service 9=Result 10=Category 11=Message 12=CorrelationId
       例:
         {GUID}<TAB>2026-10-01 02:52:34:152+0900<TAB>1<TAB>148 [AGENT_DETECTION_FINISHED]<TAB>101<TAB>{GUID}<TAB>0<TAB>0<TAB>Update;...<TAB>Success<TAB>Software Synchronization<TAB>Windows Update Client successfully detected 8 updates.<TAB>PTg+Foz... #>
    $path = "$env:SystemRoot\SoftwareDistribution\ReportingEvents.log"
    if (-not (Test-Path -LiteralPath $path)) { return @() }
    $out = @()
    try {
        foreach ($line in (Get-Content -LiteralPath $path -Tail 4000 -ErrorAction SilentlyContinue)) {
            if ($line -notmatch "`t\d+\s+\[[A-Z_]+\]") { continue }
            $f = $line -split "`t"
            if ($f.Count -lt 12) { continue }
            if ($f[3] -notmatch '\[([A-Z_]+)\]') { continue }
            $ts = ConvertFrom-ReportTime $f[1]
            $out += [pscustomobject]@{
                Time    = $(if ($ts) { $ts } else { $f[1] })
                Event   = $Matches[1]
                HResult = $f[7]
                Result  = $f[9]
                Service = $f[8]
                Message = $f[11]
            }
        }
    } catch { }
    return $out
}

function Get-LastDetectedCount($Report) {
    $d = @($Report | Where-Object { $_.Event -eq "AGENT_DETECTION_FINISHED" } | Sort-Object Time -Descending)
    if ($d.Count -eq 0) { return $null }
    if ($d[0].Message -match 'detected\s+(\d+)\s+updates') { return [int]$Matches[1] }
    return $null
}

function Get-UpdateRow($Update) {
    $kb = @()
    try { if ($Update.KBArticleIDs) { $kb = @($Update.KBArticleIDs) } } catch { }
    $downloaded = $false
    try { $downloaded = [bool]$Update.IsDownloaded } catch { }
    $reboot = $false
    try { $reboot = [bool]$Update.RebootRequired } catch { }
    $mandatory = $false
    try { $mandatory = [bool]$Update.IsMandatory } catch { }
    # IUpdate.Type は COM 経由で整数が返る (0=Other 1=Software 2=Driver 3=Firmware)
    $type = "Other"
    try {
        $type = @{ 0 = "Other"; 1 = "Software"; 2 = "Driver"; 3 = "Firmware" }[[int]$Update.Type]
    } catch { }
    return [pscustomobject]@{
        Title      = $Update.Title
        KB         = ($kb -join ",")
        Type       = $type
        Downloaded = $downloaded
        Mandatory  = $mandatory
        Reboot     = $reboot
    }
}

function Get-SearchCriteria {
    # 注意: "IsBrowseOnly=0" を併用すると WU_E_INSTALL_NOT_ALLOWED (0x80240032) になる環境がある。
    #       使用可能な組み合わせは IsInstalled=0 / Type='Software' / Type='Driver' のみ。
    $hidden = if ($IncludeHidden) { "" } else { " and IsHidden=0" }
    switch ($Type) {
        "Driver"   { return "IsInstalled=0 and Type='Driver'" }
        "Software" { return "IsInstalled=0 and Type='Software'$hidden" }
        default    { return "IsInstalled=0$hidden" }
    }
}

function Invoke-DoInstall {
    param($Bundle, [string]$Filter)

    # IUpdateColl.RemoveAt / Remove は使用できない (0x80240037 / 0x80131501)。
    # 検索結果コレクションは読み取り専用のため、新しい UpdateColl へ Add() して組み立てる。
    $coll = New-Object -ComObject Microsoft.Update.UpdateColl
    $src  = $Bundle.Result.Updates
    $names = @()
    $skipped = @()

    for ($i = 0; $i -lt $src.Count; $i++) {
        $u = $null
        try { $u = $src.Item($i) } catch { continue }
        if (-not $u) { continue }
        if ($Filter -and $u.Title -notmatch $Filter) {
            $skipped += $u.Title
            continue
        }
        # AcceptEula の戻り値は「既に受け入れ済みか」を示すだけで、失敗判定には使わない。
        # 副作用 (EULA 受け入れ) が本質。ドライバはこれを要求することが多い。
        try { $u.AcceptEula() | Out-Null } catch { }
        try {
            $coll.Add($u) | Out-Null
            $names += $u.Title
        } catch {
            Write-Err2 "更新を追加できません: $($u.Title) ($_)"
            exit 3
        }
    }

    Emit-Line ""
    Emit-Line ("  検索でヒット: {0} 件 / 対象: {1} 件 / 除外: {2} 件" -f $src.Count, $coll.Count, $skipped.Count)
    if ($skipped.Count -gt 0) {
        $skipped | Select-Object -First 10 | ForEach-Object { Emit-Line "    除外: $_" }
        if ($skipped.Count -gt 10) { Emit-Line "    ... 他 $($skipped.Count - 10) 件" }
    }

    if ($coll.Count -eq 0) {
        Write-Warn2 "インストール対象の更新がありません。"
        Set-Summary "Planned" 0
        return
    }

    Emit-Line ""
    Emit-Line "インストール対象 ($($coll.Count) 件):"
    $names | ForEach-Object { Emit-Line "  - $_" }

    if ($DryRun) {
        Set-Summary "Planned" $coll.Count
        Write-Ok "DryRun: 実際のインストールはしていません。"
        return
    }

    Write-Warn2 "ダウンロードとインストールを開始します (数十分かかることがあります)..."
    $installer = $Bundle.Session.CreateUpdateInstaller()
    $installer.Updates = $coll
    $installer.ForceQuiet = $true
    $installer.ForceReboot = (-not $NoRestart)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $res = $installer.Install()
    $sw.Stop()

    $codeMap = @{ 2 = "Succeeded"; 3 = "SucceededWithErrors"; 4 = "Failed"; 5 = "Aborted" }
    $code = [int]$res.ResultCode
    $rebootRequired = [bool]$res.RebootRequired
    Emit-Line ""
    Emit-Line ("  ResultCode : {0} ({1})" -f $code, $(if ($codeMap[$code]) { $codeMap[$code] } else { "Unknown" }))
    Emit-Line ("  HResult    : 0x{0:X8}" -f [int]$res.HResult)
    Emit-Line ("  RebootRequired : {0}" -f $rebootRequired)
    Emit-Line ("  所要時間  : {0} 分" -f (Format-Num $sw.Elapsed.TotalMinutes))
    Set-Summary "Planned" $coll.Count
    Set-Summary "ResultCode" $code
    Set-Summary "Result" $(if ($codeMap[$code]) { $codeMap[$code] } else { "Unknown($code)" })
    Set-Summary "HResult" ("0x{0:X8}" -f [int]$res.HResult)
    Set-Summary "RebootRequired" $rebootRequired
    if ($rebootRequired) {
        Write-Warn2 "再起動が必要です (-NoRestart を付けない場合のみ自動再起動します)。"
    }
    $script:Result.items = @($names | ForEach-Object { [pscustomobject]@{ Title = $_ } })
    if ($code -ne 2) { exit 3 }
}

# ---------------------------------------------------------------- actions

switch ($Action) {

    "status" {
        Write-Section "Windows Update ステータス"
        $os = Get-CimInstance Win32_OperatingSystem
        $cv = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        Emit-Line ("  OS            : {0} {1} (Build {2}.{3})" -f $os.Caption, $os.Version, $cv.CurrentBuildNumber, $cv.UBR)
        $instDate = "不明"
        try {
            if ($cv.InstallDate) { $instDate = ([DateTimeOffset]::FromUnixTimeSeconds([int64]$cv.InstallDate)).LocalDateTime.ToString("yyyy-MM-dd") }
        } catch { }
        Emit-Line ("  インストール  : {0}" -f $instDate)
        $last = (Get-CimInstance Win32_QuickFixEngineering | Sort-Object InstalledOn -Descending | Select-Object -First 1)
        Emit-Line ("  最終ホットフィクス: {0}  {1}" -f $(if ($last) { $last.HotFixID } else { "-" }),
                                                 $(if ($last) { $last.InstalledOn } else { "" }))
        Set-Summary "OS" $os.Caption
        Set-Summary "Build" ("{0}.{1}" -f $cv.CurrentBuildNumber, $cv.UBR)
        Set-Summary "LastHotfix" $(if ($last) { $last.HotFixID } else { $null })

        Write-Section "再起動フラグ"
        $flags = @()
        $cbs = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"
        $wu  = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
        $pfro = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -ErrorAction SilentlyContinue).PendingFileRenameOperations
        $flags += [pscustomobject]@{ Flag = "CBS RebootPending"; State = $cbs }
        $flags += [pscustomobject]@{ Flag = "WU Auto Update\RebootRequired"; State = $wu }
        $flags += [pscustomobject]@{ Flag = "PendingFileRenameOperations"; State = $(if ($pfro) { "あり ($($pfro.Count) 件)" } else { "なし" }) }
        New-Table $flags @("Flag", "State")
        $needReboot = $cbs -or $wu -or [bool]$pfro
        Set-Summary "RebootPending" $needReboot
        if ($needReboot) { Write-Warn2 "再起動が保留されています。" }

        Write-Section "関連サービス"
        $svc = Get-Service wuauserv, bits, cryptsvc, DoSvc, UsoSvc, WaaSMedicSvc -ErrorAction SilentlyContinue |
               Select-Object Name, Status, StartType
        New-Table $svc @("Name", "Status", "StartType")

        Write-Section "ポリシー / 一時停止"
        $state = Get-RegProps $HKLMPolicyState
        $ux    = Get-RegProps $HKLMUXSettings
        $rows = @()
        if ($state) {
            foreach ($k in @("IsDeferralIsActive", "DeferFeatureUpdates", "DeferQualityUpdates", "BranchReadinessLevel", "IsWUfBConfigured")) {
                if ($state.Contains($k)) { $rows += [pscustomobject]@{ Key = $k; Value = $state[$k] } }
            }
        }
        if ($ux -and $ux.Contains("IsContinuousInnovationOptedIn")) {
            $rows += [pscustomobject]@{ Key = "IsContinuousInnovationOptedIn (Insider)"; Value = $ux["IsContinuousInnovationOptedIn"] }
        }
        $pauseState = Get-PauseState
        $rows += [pscustomobject]@{ Key = "IsPaused"; Value = $pauseState.IsPaused }
        $rows += [pscustomobject]@{ Key = "PausedFeatureStatus"; Value = $pauseState.PausedFeatureStatus }
        $rows += [pscustomobject]@{ Key = "PausedQualityStatus"; Value = $pauseState.PausedQualityStatus }
        $rows += [pscustomobject]@{ Key = "ConfiguredExpiry"; Value = $pauseState.ConfiguredExpiry }
        New-Table $rows @("Key", "Value")
        Set-Summary "Paused" $pauseState.IsPaused

        Write-Section "最近のイベント (24時間, Warning 以上)"
        try {
            $ev = Get-WinEvent -FilterHashtable @{ LogName = $WULog; StartTime = (Get-Date).AddHours(-24); Level = 1, 2, 3 } -MaxEvents 12 -ErrorAction Stop
            $ev | Select-Object @{n = "TimeCreated"; e = { Format-DateTime $_.TimeCreated } }, Id, @{n = "Level"; e = { $_.LevelDisplayName } },
                              @{n = "Message"; e = { ($_.Message -replace '\s+', ' ').Substring(0, [Math]::Min(110, ($_.Message -replace '\s+', ' ').Length)) } } |
                Format-Table -AutoSize | Out-String -Width 200 | ForEach-Object { Emit-Line $_.TrimEnd() }
        } catch { Emit-Line "  (イベントなし / ログなし)" }

        if ($Deep) {
            Write-Section "未適用更新 (検索)"
            $b = Invoke-UpdateSearch (Get-SearchCriteria)
            if ($null -eq $b) {
                Write-OrchestratorFallback
            } else {
                $rowsU = @()
                for ($i = 0; $i -lt $b.Result.Updates.Count; $i++) { $rowsU += Get-UpdateRow $b.Result.Updates.Item($i) }
                New-Table $rowsU @("Title", "KB", "Type", "Reboot")
                Set-Summary "PendingCount" $rowsU.Count
            }
        } else {
            Write-Info "(未適用更新の数は -Deep / search で見ます)"
        }

        Complete-Result @()
    }

    "search" {
        Write-Section "未適用更新 ($Type)"
        $b = Invoke-UpdateSearch (Get-SearchCriteria) -Offline:$OfflineSearch
        if ($null -eq $b) {
            Write-OrchestratorFallback
            Complete-Result @()
        }
        $rows = @()
        for ($i = 0; $i -lt $b.Result.Updates.Count; $i++) {
            $rows += Get-UpdateRow $b.Result.Updates.Item($i)
        }
        if ($Title) { $rows = @($rows | Where-Object { $_.Title -match $Title }) }
        if ($rows.Count -gt $MaxResults) {
            Write-Warn2 "$($rows.Count) 件中 $MaxResults 件を表示します (-MaxResults で変更)。"
            $rows = $rows[0..($MaxResults - 1)]
        }
        New-Table $rows @("Title", "KB", "Type", "Reboot", "Downloaded", "Mandatory")
        Set-Summary "Count" $rows.Count
        Set-Summary "SoftwareCount" @($rows | Where-Object Type -eq "Software").Count
        Set-Summary "DriverCount" @($rows | Where-Object Type -eq "Driver").Count
        $script:Result.items = $rows
        Complete-Result @("Title", "KB", "Type", "Reboot")
    }

    "install" {
        Require-Admin "更新のインストール"
        Write-Section "更新のインストール"
        $b = Invoke-UpdateSearch (Get-SearchCriteria)
        if ($null -eq $b) {
            Write-OrchestratorFallback
            Write-Section "代替手段"
            Write-Info "以下のコマンドで Update Orchestrator に適用を要求できます:"
            Write-Info "  UsoClient StartDownload   (未適用の更新をダウンロード)"
            Write-Info "  UsoClient StartInstall    (ダウンロード済み分をインストール)"
            Write-Info "  UsoClient ScanInstallWait (スキャン -> ダウンロード -> インストール)"
            Write-Info "または -Orchestrate 付きで再実行すると上記を自動実行します。"
            if ($Orchestrate) {
                Write-Section "Update Orchestrator へ要求"
                Invoke-OrchestrateInstall -NoRestart:$NoRestart
                Set-Summary "InstallMethod" "orchestrator"
            }
            Set-Summary "InstallMethodFallback" $(if ($Orchestrate) { "orchestrator" } else { "manual" })
            Complete-Result @()
        }
        $all = @()
        for ($i = 0; $i -lt $b.Result.Updates.Count; $i++) { $all += Get-UpdateRow $b.Result.Updates.Item($i) }
        Set-Summary "Found" $all.Count

        if ($IncludeDrivers -and $Type -ne "Driver") {
            Write-Section "ドライバも適用対象にします"
            $bd = Invoke-UpdateSearch "IsInstalled=0 and Type='Driver'"
            if ($null -eq $bd) {
                Write-Warn2 "ドライバ検索は利用できないためスキップします。"
            } else {
                $drv = @()
                for ($i = 0; $i -lt $bd.Result.Updates.Count; $i++) { $drv += Get-UpdateRow $bd.Result.Updates.Item($i) }
                $all += $drv
                Set-Summary "Found" $all.Count
            }
        }

        $filter = if ($Title) { $Title } else { $null }
        Invoke-DoInstall $b $filter
        Complete-Result @("Title")
    }

    "history" {
        Write-Section "更新履歴 (最新 $Count 件)"
        $s = New-UpdateSession
        $searcher = $s.CreateUpdateSearcher()
        $total = $searcher.GetTotalHistoryCount()
        $take = [Math]::Min($Count, $total)
        $h = $searcher.QueryHistory(0, $take)
        $codeMap = @{ 2 = "Installed"; 3 = "Failed"; 4 = "Aborted"; 5 = "Superseded"; 6 = "DetectionError"; 7 = "NonApplicable" }
        $rows = @()
        for ($i = 0; $i -lt $h.Count; $i++) {
            $e = $h.Item($i)
            $code = [int]$e.ResultCode
            $rows += [pscustomobject]@{
                Date     = (Format-DateTime $e.Date "yyyy-MM-dd HH:mm")
                Result   = $(if ($codeMap[$code]) { $codeMap[$code] } else { "Code$code" })
                HResult  = ("0x{0:X8}" -f [int]$e.HResult)
                # ClientApplicationID は外部アプリの識別子なので出さない (実行元の記録を除く)
                Client   = $(try { if ($e.ClientApplicationID) { "yes" } else { "" } } catch { "" })
                Title    = $e.Title
            }
        }
        if ($Search) { $rows = @($rows | Where-Object { $_.Title -match $Search }) }
        New-Table $rows @("Date", "Result", "HResult", "Title")
        Set-Summary "TotalHistory" $total
        Set-Summary "Shown" $rows.Count
        $script:Result.items = $rows
        Complete-Result @("Date", "Result", "Title")
    }

    "hotfix" {
        Write-Section "インストール済みホットフィクス (CIM)"
        # InstalledBy (Windows アカウント名) は個人情報のため出力しない
        $rows = Get-CimInstance Win32_QuickFixEngineering |
                Sort-Object InstalledOn -Descending |
                Select-Object @{n = "InstalledOn"; e = { $_.InstalledOn.ToString("yyyy-MM-dd") } },
                              HotFixID, Description
        New-Table $rows @("InstalledOn", "HotFixID", "Description")
        Set-Summary "Count" @($rows).Count
        Complete-Result @("InstalledOn", "HotFixID", "Description")
    }

    "policy" {
        Write-Section "Windows Update ポリシー"

        Write-Section "UpdatePolicy\Settings"
        $st = Get-RegProps $HKLMUpdatePolicy
        if ($st) {
            $rows = foreach ($k in $st.Keys) { [pscustomobject]@{ Key = $k; Value = $st[$k] } }
            New-Table $rows @("Key", "Value")
        } else { Emit-Line "  (なし)" }

        Write-Section "UpdatePolicy\PolicyState"
        $ps = Get-RegProps $HKLMPolicyState
        if ($ps) {
            $rows = foreach ($k in $ps.Keys) { [pscustomobject]@{ Key = $k; Value = $ps[$k] } }
            New-Table $rows @("Key", "Value")
        } else { Emit-Line "  (なし)" }

        Write-Section "Group Policy (HKLM\SOFTWARE\Policies\...\WindowsUpdate)"
        if (Test-Path $HKLMWUPolicy) {
            $rows = @()
            $wp = Get-RegProps $HKLMWUPolicy
            foreach ($k in $wp.Keys) { $rows += [pscustomobject]@{ Key = $k; Value = ($wp[$k] -join ",") } }
            $ap = Get-RegProps $HKLMAUPolicy
            foreach ($k in $ap.Keys) { $rows += [pscustomobject]@{ Key = "AU\$k"; Value = ($ap[$k] -join ",") } }
            New-Table $rows @("Key", "Value")
        } else {
            Emit-Line "  (ポリシーは未設定 = Windows Update 既定動作)"
            Set-Summary "Managed" $false
        }
        if (Test-Path $HKLMWUPolicy) { Set-Summary "Managed" $true }

        Write-Section "ServiceManager / 一時停止 (レジストリ実効値)"
        try {
            $sm = New-Object -ComObject Microsoft.Update.ServiceManager
            Emit-Line ("  ServerSelection / ClientApplicationID : {0}" -f $(try { $sm.ClientApplicationID } catch { "(未設定)" }))
            Emit-Line "  COM ServiceManager は GetIsPauseUpdatesEffective を公開しないため、レジストリで判定:"
            $pst = Get-PauseState
            New-Table @(
                [pscustomobject]@{ Key = "IsPaused"; Value = $pst.IsPaused }
                [pscustomobject]@{ Key = "PausedFeatureStatus"; Value = $pst.PausedFeatureStatus }
                [pscustomobject]@{ Key = "PausedQualityStatus"; Value = $pst.PausedQualityStatus }
                [pscustomobject]@{ Key = "PausedFeatureExpiry"; Value = $pst.PausedFeatureExpiry }
                [pscustomobject]@{ Key = "ConfiguredExpiry (UX)"; Value = $pst.ConfiguredExpiry }
            ) @("Key", "Value")
        } catch { Write-Warn2 "ServiceManager 取得失敗: $_" }
        Complete-Result @()
    }

    "set-target" {
        Require-Admin "ターゲットリリースの設定"
        if (-not $TargetRelease) { Write-Err2 "-TargetRelease を指定してください (例: 23H2 / 25H2 / 26100)"; exit 2 }
        $raw = ($TargetRelease -replace '[^0-9]', '')
        if ($raw.Length -eq 0) { Write-Err2 "-TargetRelease の形式が不正です: $TargetRelease"; exit 2 }
        $value = if ($raw.Length -gt 4) { [int]$raw } else { [int]([math]::Floor([int]$raw / 100) * 100 + ($raw % 100)) }
        if ($raw.Length -le 4) {
            $value = [int]$raw * 100
            Write-Info "MMHR 形式 [$raw] を TargetReleaseVersion = $value に変換しました。"
        }
        New-Item -Path $HKLMUpdatePolicy -Force | Out-Null
        New-ItemProperty -Path $HKLMUpdatePolicy -Name TargetReleaseVersion -Value $value -PropertyType DWord -Force | Out-Null
        Set-Summary "TargetReleaseVersion" $value
        Write-Ok "TargetReleaseVersion = $value を設定しました。"
        try {
            Start-Process "$env:SystemRoot\System32\UsoClient.exe" -ArgumentList "RefreshSettings" -WindowStyle Hidden -Wait
            Emit-Line "  UsoClient RefreshSettings 実行済み (反映に数分かかる場合があります)。"
        } catch { Write-Warn2 "UsoClient 実行失敗: $_" }
        Complete-Result @()
    }

    "clear-target" {
        Require-Admin "ターゲットリリースの解除"
        Remove-ItemProperty -Path $HKLMUpdatePolicy -Name TargetReleaseVersion -ErrorAction SilentlyContinue
        Write-Ok "TargetReleaseVersion を解除しました (最新リリースを追従)。"
        try { Start-Process "$env:SystemRoot\System32\UsoClient.exe" -ArgumentList "RefreshSettings" -WindowStyle Hidden -Wait } catch { }
        Complete-Result @()
    }

    "pause" {
        Require-Admin "更新の一時停止"
        Write-Section "更新の一時停止 ($PauseDays 日 / $DeferType)"
        $before = Get-PauseState
        Emit-Line ("  現在: IsPaused={0} Feature={1} Quality={2}" -f $before.IsPaused, $before.PausedFeatureStatus, $before.PausedQualityStatus)
        if ($DryRun) {
            Write-Info "[DryRun] 実行しない。以下のレジストリに書き込む想定:"
            Write-Info "  HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"
            Write-Info "  PauseUpdatesExpiryTime = $((Get-Date).AddDays($PauseDays).ToString('yyyy-MM-dd'))"
            Write-Info "  Pause{Feature|Quality}UpdatesExpiryTime = 同上"
            Set-Summary "PausedUntil" "dry-run"
            Complete-Result @()
        }
        try {
            $expiry = Write-PauseRegistry -Days $PauseDays -DeferType $DeferType
            Invoke-UsoRefresh
            Start-Sleep -Seconds 3
            Write-Ok "$DeferType 更新を $PauseDays 日間 (~$expiry まで) 一時停止しました。"
            Set-Summary "PausedUntil" $expiry
            $after = Get-PauseState
            Write-Info ("  反映確認: IsPaused={0} FeatureStatus={1} QualityStatus={2}" -f $after.IsPaused, $after.PausedFeatureStatus, $after.PausedQualityStatus)
            if (-not $after.IsPaused) {
                Write-Warn2 "レジストリには設定しましたが OS 側の実効値がまだ反映されていません。数分待って policy / status を確認してください。"
            }
        } catch {
            Write-Err2 "一時停止の設定に失敗しました: $_"
            exit 3
        }
        Complete-Result @()
    }

    "resume" {
        Require-Admin "一時停止の解除"
        Write-Section "一時停止の解除"
        $before = Get-PauseState
        Emit-Line ("  現在: IsPaused={0} Feature={1} Quality={2}" -f $before.IsPaused, $before.PausedFeatureStatus, $before.PausedQualityStatus)
        if ($DryRun) {
            Write-Info "[DryRun] 実行しない。UX\Settings の Pause* レジストリ値を削除する想定。"
            Complete-Result @()
        }
        try {
            Write-PauseRegistry -Clear | Out-Null
            Invoke-UsoRefresh
            Start-Sleep -Seconds 3
            Write-Ok "一時停止の設定値を削除しました。"
            $after = Get-PauseState
            Write-Info ("  反映確認: IsPaused={0} FeatureStatus={1} QualityStatus={2}" -f $after.IsPaused, $after.PausedFeatureStatus, $after.PausedQualityStatus)
            Set-Summary "Paused" $after.IsPaused
        } catch {
            Write-Err2 "解除に失敗しました: $_"
            exit 3
        }
        Complete-Result @()
    }

    "cache" {
        Write-Section "更新キャッシュ"
        $targets = [ordered]@{
            "SoftwareDistribution\Download"      = "$env:SystemRoot\SoftwareDistribution\Download"
            "SoftwareDistribution\DataStore"     = "$env:SystemRoot\SoftwareDistribution\DataStore"
            "DeliveryOptimization Cache"         = "$env:SystemRoot\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache"
            "DeliveryOptimization Config"        = "$env:SystemRoot\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Config"
        }
        $rows = @()
        $total = 0
        foreach ($k in $targets.Keys) {
            $sz = Get-DirSize $targets[$k]
            $total += $sz
            $rows += [pscustomobject]@{ Cache = $k; Size = (ConvertTo-HumanSize $sz); Path = $targets[$k] }
        }
        New-Table $rows @("Cache", "Size")
        Emit-Line ("  合計: {0}" -f (ConvertTo-HumanSize $total))
        Set-Summary "TotalBytes" $total
        Set-Summary "Total" (ConvertTo-HumanSize $total)

        if ($Clean) {
            Require-Admin "キャッシュ削除"
            Write-Section "キャッシュ削除"
            if (Get-Command Delete-DeliveryOptimizationCache -ErrorAction SilentlyContinue) {
                try {
                    Delete-DeliveryOptimizationCache -Force
                    Write-Ok "DeliveryOptimization キャッシュを cmdlet で削除しました。"
                } catch { Write-Warn2 "Delete-DeliveryOptimizationCache 失敗: $_" }
            }
            foreach ($svc in @("wuauserv", "bits", "DoSvc", "UsoSvc")) {
                try { Stop-Service -Name $svc -Force -ErrorAction Stop; Write-Info "停止: $svc" } catch { }
            }
            Start-Sleep -Seconds 3
            foreach ($k in @("SoftwareDistribution\Download", "DeliveryOptimization Cache")) {
                $p = $targets[$k]
                if (-not (Test-Path -LiteralPath $p)) { continue }
                try {
                    Get-ChildItem -LiteralPath $p -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Ok "削除: $k"
                } catch { Write-Warn2 "削除失敗 $k : $_" }
            }
            foreach ($svc in @("bits", "DoSvc", "wuauserv")) {
                try { Start-Service -Name $svc -ErrorAction SilentlyContinue; Write-Info "開始: $svc" } catch { }
            }
            try { Start-Process "$env:SystemRoot\System32\UsoClient.exe" -ArgumentList "StartScan" -WindowStyle Hidden -Wait } catch { }
            Set-Summary "Cleaned" $true
            Write-Ok "キャッシュ削除完了。次回の検索で再取得されます。"
        }
        $script:Result.items = $rows
        Complete-Result @("Cache", "Size")
    }

    "log" {
        Write-Section "Windows Update イベントログ (直近 $Days 日 / $Level)"
        $lv = switch ($Level) {
            "Error"   { @(1, 2) }
            "Warning" { @(1, 2, 3) }
            "Info"    { @(1, 2, 3, 4) }
            default   { @(1, 2, 3, 4, 5) }
        }
        $logs = @($WULog, "Microsoft-Windows-Servicing")
        $rows = @()
        foreach ($lg in $logs) {
            if (-not (Get-WinEvent -ListLog $lg -ErrorAction SilentlyContinue)) { continue }
            try {
                $ev = Get-WinEvent -FilterHashtable @{ LogName = $lg; StartTime = (Get-Date).AddDays(-$Days); Level = $lv } -ErrorAction Stop
            } catch { continue }
            foreach ($e in $ev) {
                $msg = ($e.Message -replace '\s+', ' ').Trim()
                if ($msg.Length -gt 160) { $msg = $msg.Substring(0, 160) + "..." }
                $rows += [pscustomobject]@{
                    Time  = (Format-DateTime $e.TimeCreated)
                    Log   = $lg.Split("/")[-1]
                    Id    = $e.Id
                    Level = $e.LevelDisplayName
                    Message = $msg
                }
            }
        }
        if (-not $rows) { Emit-Line "  (該当イベントなし)"; Set-Summary "Count" 0; Complete-Result @(); break }
        $rows = @($rows | Sort-Object Time -Descending | Select-Object -First $Count)
        New-Table $rows @("Time", "Id", "Level", "Message")
        Set-Summary "Count" $rows.Count
        $script:Result.items = $rows
        Complete-Result @("Time", "Id", "Level", "Message")
    }

    "report" {
        Write-Section "Update Orchestrator レポート (直近 $Count 件)"
        $rep = Get-OrchestratorReport
        if (@($rep).Count -eq 0) {
            Emit-Line "  レポートを読み取れません: $env:SystemRoot\SoftwareDistribution\ReportingEvents.log"
            Set-Summary "Count" 0
            Complete-Result @()
        }
        $rows = @($rep | Sort-Object Time -Descending)
        $det = Get-LastDetectedCount $rep
        if ($null -ne $det) {
            Write-Info ("  最終検出の未適用更新: {0} 件" -f $det)
            Set-Summary "LastDetectedCount" $det
        }
        Write-Info ("  パース済みイベント: {0} 件 / 、直近 {1} 件を表示" -f $rows.Count, [Math]::Min($Count, $rows.Count))
        $show = @($rows | Select-Object -First $Count)
        $table = $show | Select-Object @{n = "Time"; e = { Format-DateTime $_.Time } }, Event, Result, HResult,
                                 @{n = "Message"; e = { if ($_.Message.Length -gt 90) { $_.Message.Substring(0, 90) + "..." } else { $_.Message } } }
        $table | Format-Table -AutoSize | Out-String -Width 200 | ForEach-Object { Emit-Line $_.TrimEnd() }
        Set-Summary "Count" $rows.Count
        $script:Result.items = @($show | ForEach-Object {
            [pscustomobject]@{ Time = (Format-DateTime $_.Time); Event = $_.Event; Result = $_.Result; HResult = $_.HResult; Message = $_.Message }
        })
        Complete-Result @("Time", "Event", "Result", "Message")
    }

    "selftest" {
        Write-Section "自己診断 (selftest)"
        $rows = @()
        $rows += [pscustomobject]@{ Check = "OS"; Status = $(try { (Get-CimInstance Win32_OperatingSystem).Caption } catch { "NG: $_" }) }
        $rows += [pscustomobject]@{ Check = "Admin"; Status = $(if (Test-Admin) { "yes" } else { "no (write actions blocked)" }) }
        $rows += [pscustomobject]@{ Check = "PowerShell"; Status = $PSVersionTable.PSVersion.ToString() }
        $rows += [pscustomobject]@{ Check = "COM Microsoft.Update.Session"; Status = $(try { $s = New-Object -ComObject Microsoft.Update.Session; "ok" } catch { "NG: 0x$('{0:X8}' -f [int]$_.Exception.HResult)" }) }
        $rows += [pscustomobject]@{ Check = "COM history query"; Status = $(try { $s = New-Object -ComObject Microsoft.Update.Session; $n = $s.CreateUpdateSearcher().GetTotalHistoryCount(); "ok ($n records)" } catch { "NG" }) }
        $rows += [pscustomobject]@{ Check = "WUA online search"; Status = $(try { $s = New-Object -ComObject Microsoft.Update.Session; $r = $s.CreateUpdateSearcher(); $r.Online = $true; $x = $r.Search("IsInstalled=0"); "ok ($($x.Updates.Count))" } catch { "NG: 0x$('{0:X8}' -f [int]$_.Exception.HResult)" }) }
        $rows += [pscustomobject]@{ Check = "ReportingEvents.log"; Status = $(try { $f = "$env:SystemRoot\SoftwareDistribution\ReportingEvents.log"; if (Test-Path $f) { "ok ($([math]::Round((Get-Item $f).Length/1KB)) KB)" } else { "NG: not found" } } catch { "NG" }) }
        $rows += [pscustomobject]@{ Check = "Orchestrator parse"; Status = $(try { $r = @(Get-OrchestratorReport); "ok ($($r.Count) events)" } catch { "NG: $_" }) }
        $rows += [pscustomobject]@{ Check = "EventLog WU client"; Status = $(try { $l = Get-WinEvent -ListLog $WULog -ErrorAction Stop; "ok ($($l.RecordCount) records)" } catch { "NG: $($_.Exception.Message)" }) }
        $rows += [pscustomobject]@{ Check = "Scheduled task API"; Status = $(try { Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop | Out-Null; "ok (registered)" } catch { "ok (not registered)" }) }
        $rows += [pscustomobject]@{ Check = "Redaction ON"; Status = $script:Redact }
        $rows += [pscustomobject]@{ Check = "Identity tokens"; Status = "$($script:Identity.Count) masked" }
        New-Table $rows @("Check", "Status")

        Write-Section "Redaction unit test"
        $samples = @(
            @{ In = "user=$env:USERNAME"; Expect = "not contains $env:USERNAME" },
            @{ In = "pc=$env:COMPUTERNAME"; Expect = "not contains $env:COMPUTERNAME" },
            @{ In = "guid={C268EEF9-0B96-4382-9D94-93F38A510B72}"; Expect = "guid masked" },
            @{ In = "mail=someone@example.com"; Expect = "masked" },
            @{ In = "ip=192.168.1.42"; Expect = "masked" },
            @{ In = "path=C:\Users\Someone\AppData"; Expect = "unchanged (generic)" }
        )
        $trows = @()
        foreach ($s in $samples) {
            $out = Protect-PII $s.In
            $ok = switch -Regex ($s.Expect) {
                "^not contains" { -not $out.Contains($Matches[1]) }
                "guid masked"   { $out -notmatch 'C268EEF9' }
                "masked"        { $out -match '\*\*\*|x\.x' }
                "unchanged"     { $out -eq $s.In }
                default         { $true }
            }
            $trows += [pscustomobject]@{ Input = $s.In; Output = $out; Expect = $s.Expect; Pass = $(if ($ok) { "PASS" } else { "FAIL" }) }
        }
        New-Table $trows @("Input", "Output", "Expect", "Pass")
        Set-Summary "RedactionFails" @($trows | Where-Object Pass -eq "FAIL").Count
        Complete-Result @("Check", "Status")
    }

    "reset" {
        Require-Admin "Windows Update コンポーネントのリセット"
        Write-Section "1. サービス停止"
        foreach ($svc in @("wuauserv", "bits", "cryptsvc", "msiserver", "DoSvc", "UsoSvc", "WaaSMedicSvc")) {
            try {
                Stop-Service -Name $svc -Force -ErrorAction Stop
                Write-Ok "停止: $svc"
            } catch { Write-Warn2 "停止できない/停止不要: $svc" }
        }
        Start-Sleep -Seconds 3

        Write-Section "2. SoftwareDistribution の退避"
        $stamp = Get-Date -Format "yyyyMMddHHmmss"
        $sd = "$env:SystemRoot\SoftwareDistribution"
        if (Test-Path $sd) {
            $bak = "$sd.old-$stamp"
            try {
                Rename-Item -LiteralPath $sd -NewName (Split-Path -Leaf $bak) -ErrorAction Stop
                Write-Ok "退避: $bak"
                Set-Summary "BackupPath" $bak
            } catch { Write-Err2 "SoftwareDistribution の退避に失敗: $_" }
        }
        $docache = "$env:SystemRoot\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache"
        if (Test-Path $docache) {
            try {
                Get-ChildItem -LiteralPath $docache -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                Write-Ok "DeliveryOptimization キャッシュを削除"
            } catch { Write-Warn2 "DO キャッシュ削除失敗: $_" }
        }

        Write-Section "3. WSUS / ポリシー確認"
        if (Test-Path $HKLMWUPolicy) {
            Write-Warn2 "グループポリシーが設定されています:"
            (Get-RegProps $HKLMWUPolicy).GetEnumerator() | ForEach-Object { Write-Info "   $($_.Key) = $($_.Value)" }
            Write-Info "WSUS 原因的リセットが必要な場合は -ClearPolicy で削除します。"
        } else {
            Write-Ok "WSUS / 集中管理ポリシーは未設定です。"
        }
        if ($ClearRenameOps) {
            try {
                Remove-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -ErrorAction Stop
                Write-Ok "PendingFileRenameOperations を削除"
            } catch { Write-Warn2 "PendingFileRenameOperations 削除失敗: $_" }
        }

        if ($ReRegister) {
            Write-Section "4. DLL 再登録"
            $dlls = @("wuaueng.dll", "wucltui.dll", "wuapi.dll", "atl.dll", "shell32.dll", "ntdll.dll")
            foreach ($d in $dlls) {
                $path = Join-Path $env:SystemRoot "System32\$d"
                if (Test-Path $path) {
                    $p = Start-Process "$env:SystemRoot\System32\regsvr32.exe" -ArgumentList "/s", "/i", $path -Wait -PassThru
                    Write-Info "$d -> exit $($p.ExitCode)"
                }
            }
        }

        if ($RunDism) {
            Write-Section "5. DISM 修復 (数十分かかることがあります)"
            & "$env:SystemRoot\System32\Dism.exe" /Online /Cleanup-Image /RestoreHealth /NoRestart
            Write-Info "DISM 終了コード: $LASTEXITCODE"
        }

        Write-Section "サービス再開"
        foreach ($svc in @("cryptsvc", "bits", "wuauserv", "DoSvc", "UsoSvc")) {
            try { Start-Service -Name $svc -ErrorAction SilentlyContinue; Write-Ok "開始: $svc" } catch { Write-Warn2 "開始失敗: $svc" }
        }
        try { Start-Process "$env:SystemRoot\System32\UsoClient.exe" -ArgumentList "RefreshSettings" -WindowStyle Hidden -Wait } catch { }
        try { Start-Process "$env:SystemRoot\System32\UsoClient.exe" -ArgumentList "StartScan" -WindowStyle Hidden -Wait } catch { }
        Write-Ok "リセット完了。search / status を実行して確認してください。"
        Complete-Result @()
    }

    "schedule" {
        Write-Section "自動インストールタスク ($TaskName)"
        switch ($TaskState) {

            "status" {
                $t = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
                if (-not $t) { Emit-Line "  未登録です。"; Set-Summary "Registered" $false; break }
                $i = Get-ScheduledTaskInfo -TaskName $TaskName
                New-Table @([pscustomobject]@{
                    State      = $t.State
                    NextRun    = $(if ($i.NextRunTime) { $i.NextRunTime } else { "-" })
                    LastRun    = $(if ($i.LastRunTime) { $i.LastRunTime } else { "-" })
                    LastResult = $i.LastTaskResult
                    Principal  = $t.Principal.UserId
                }) @("State", "NextRun", "LastRun", "LastResult", "Principal")
                Set-Summary "Registered" $true
                Set-Summary "NextRun" $(if ($i.NextRunTime) { $i.NextRunTime.ToString("s") } else { $null })
            }

            "install" {
                Require-Admin "タスク登録"
                $time = [DateTime]::ParseExact($At, "HH:mm", $null)
                $arg = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$ScriptPath`" -Action install -NoRestart -LogFile `"$env:ProgramData\wupm\auto-update.log`""
                $action = New-ScheduledTaskAction -Execute (Get-Command powershell.exe).Source -Argument $arg
                $trigger = New-ScheduledTaskTrigger -Daily -At $time
                $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
                    -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 3)
                if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
                    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
                }
                Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal `
                    -Settings $settings -Description "wupm: daily silent Windows Update install (no reboot)" | Out-Null
                Write-Ok "登録完了: 毎日 $At に SYSTEM 権限で更新をサイレント適用 (再起動なし)。"
                Write-Info "ログ: $env:ProgramData\wupm\auto-update.log"
                Set-Summary "Registered" $true
                Set-Summary "At" $At
            }

            "run" {
                Write-Ok "タスクを今すぐ実行します (再起動なし)"
                Start-ScheduledTask -TaskName $TaskName
                Start-Sleep -Seconds 5
                $i = Get-ScheduledTaskInfo -TaskName $TaskName
                Emit-Line ("  State = {0} / LastResult = {1}" -f (Get-ScheduledTask -TaskName $TaskName).State, $i.LastTaskResult)
            }

            "remove" {
                Require-Admin "タスク解除"
                if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
                    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
                    Write-Ok "解除しました。"
                    Set-Summary "Registered" $false
                } else { Emit-Line "  未登録でした。" }
            }
        }
        Complete-Result @()
    }
}

Complete-Result @()