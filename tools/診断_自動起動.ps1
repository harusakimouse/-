<#
  0945_baibai.xlsm が勝手に立ち上がる — Windows側の原因診断スクリプト

  使い方:
    1. このファイルを右クリック → 「PowerShell で実行」
       または PowerShell を開いて:
         powershell -ExecutionPolicy Bypass -File .\診断_自動起動.ps1
    2. デスクトップに 自動起動診断_Windows_yyyyMMdd_HHmmss.txt が出力される
    3. その中身を貼り付ける

  読み取りのみ。設定は一切変更しません。
#>

$ErrorActionPreference = 'SilentlyContinue'
$TargetName = '0945_baibai'      # 探すファイル名（拡張子なし・部分一致）

$out = New-Object System.Collections.Generic.List[string]
function W([string]$s) { $out.Add($s) }

W ('=' * 60)
W " 自動起動 診断レポート (Windows側)"
W " 対象   : $TargetName"
W " 生成   : $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')"
W " PC     : $env:COMPUTERNAME / $env:USERNAME"
W ('=' * 60)
W ''

# ---------------------------------------------------------------
W '[1] タスク スケジューラ  ★最有力候補'
W '    (09:45 のような時刻で登録されていないか)'
W ''
$hitTask = $false
Get-ScheduledTask | ForEach-Object {
    $t = $_
    $actions = $t.Actions | ForEach-Object {
        "$($_.Execute) $($_.Arguments)"
    }
    $joined = ($actions -join ' ')
    if ($joined -match $TargetName -or $joined -match 'EXCEL' -or $joined -match '\.xlsm') {
        $hitTask = $true
        $info = $t | Get-ScheduledTaskInfo
        W "  ◆ タスク名 : $($t.TaskName)"
        W "    パス     : $($t.TaskPath)"
        W "    状態     : $($t.State)"
        W "    次回実行 : $($info.NextRunTime)"
        W "    前回実行 : $($info.LastRunTime)"
        foreach ($a in $actions) { W "    実行内容 : $a" }
        foreach ($tr in $t.Triggers) {
            W "    トリガー : $($tr.CimClass.CimClassName)  開始=$($tr.StartBoundary)  有効=$($tr.Enabled)"
        }
        W ''
    }
}
if (-not $hitTask) { W '  (該当なし)'; W '' }

# ---------------------------------------------------------------
W '[2] Excel の XLSTART フォルダ  ★原因No.2'
W '    (ここに置いたファイルは Excel 起動時に必ず開かれる)'
W ''
$xlstartPaths = @(
    "$env:APPDATA\Microsoft\Excel\XLSTART"
    "$env:ProgramFiles\Microsoft Office\root\Office16\XLSTART"
    "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\XLSTART"
    "$env:ProgramFiles\Microsoft Office\Office16\XLSTART"
)
foreach ($p in $xlstartPaths) {
    if (Test-Path $p) {
        $files = Get-ChildItem $p -File
        if ($files) {
            W "  ◆ $p"
            $files | ForEach-Object { W "      → $($_.Name)   ($($_.Length) bytes, $($_.LastWriteTime))" }
        } else {
            W "  - $p  (空 = 正常)"
        }
    }
}
W ''

# ---------------------------------------------------------------
W '[3] Excel の「起動時にすべてのファイルを開くフォルダー」(AltStartup)'
W ''
$found = $false
Get-ChildItem 'HKCU:\Software\Microsoft\Office' -ErrorAction SilentlyContinue |
    Where-Object { $_.PSChildName -match '^\d+\.\d+$' } | ForEach-Object {
        $k = "HKCU:\Software\Microsoft\Office\$($_.PSChildName)\Excel\Options"
        $v = (Get-ItemProperty $k -Name 'AltStartup' -ErrorAction SilentlyContinue).AltStartup
        if ($v) { W "  ◆ $k"; W "      AltStartup = $v"; $found = $true }
    }
if (-not $found) { W '  (未設定 = 正常)' }
W ''

# ---------------------------------------------------------------
W '[4] スタートアップ フォルダ'
W ''
$startupPaths = @(
    "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
    "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"
)
foreach ($p in $startupPaths) {
    W "  $p"
    $files = Get-ChildItem $p -File
    if ($files) { $files | ForEach-Object { W "      → $($_.Name)" } }
    else { W '      (空 = 正常)' }
}
W ''

# ---------------------------------------------------------------
W '[5] レジストリの自動実行 (Run / RunOnce)'
W ''
$runKeys = @(
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce'
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run'
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce'
)
$hitRun = $false
foreach ($k in $runKeys) {
    $props = Get-ItemProperty $k -ErrorAction SilentlyContinue
    if ($props) {
        $props.PSObject.Properties |
            Where-Object { $_.Name -notmatch '^PS' } | ForEach-Object {
                if ("$($_.Value)" -match $TargetName -or "$($_.Value)" -match 'EXCEL|\.xlsm|\.vbs|\.bat') {
                    W "  ◆ $k"
                    W "      $($_.Name) = $($_.Value)"
                    $hitRun = $true
                }
            }
    }
}
if (-not $hitRun) { W '  (Excel関連の登録なし)' }
W ''

# ---------------------------------------------------------------
W '[6] Windows「サインイン後にアプリを自動的に再起動する」設定'
W '    ONだと、シャットダウン時に開いていたExcelが次回サインインで復活する'
W ''
$ral = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name 'RestartApps' -ErrorAction SilentlyContinue).RestartApps
if ($null -eq $ral) { W '  RestartApps = (未設定)' }
else { W "  RestartApps = $ral   (1 = ON ★これが原因の可能性あり / 0 = OFF)" }
W ''

# ---------------------------------------------------------------
W "[7] $TargetName という名前のファイルを検索"
W ''
$searchRoots = @($env:USERPROFILE, 'C:\') | Select-Object -Unique
$seen = @{}
foreach ($root in $searchRoots) {
    Get-ChildItem -Path $root -Recurse -Filter "*$TargetName*" -ErrorAction SilentlyContinue -Force |
        Select-Object -First 40 | ForEach-Object {
            if (-not $seen.ContainsKey($_.FullName)) {
                $seen[$_.FullName] = $true
                W "  - $($_.FullName)   ($($_.Length) bytes, $($_.LastWriteTime))"
            }
        }
    if ($seen.Count -gt 0) { break }
}
if ($seen.Count -eq 0) { W '  (見つからず)' }
W ''

# ---------------------------------------------------------------
W '[8] 現在動いている Excel プロセス'
W ''
$procs = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($procs) {
    $procs | ForEach-Object {
        W "  - PID $($_.Id)  起動時刻 $($_.StartTime)  メモリ $([math]::Round($_.WorkingSet64/1MB,1))MB"
    }
} else { W '  (起動していない)' }
W ''

# ---------------------------------------------------------------
W '[9] 直近のイベントログ (Excel の起動記録 / 過去24時間)'
W ''
$since = (Get-Date).AddDays(-1)
$evts = Get-WinEvent -FilterHashtable @{
    LogName='Microsoft-Windows-TaskScheduler/Operational'; StartTime=$since
} -MaxEvents 200 -ErrorAction SilentlyContinue |
    Where-Object { $_.Message -match 'EXCEL|xlsm' } | Select-Object -First 20
if ($evts) {
    $evts | ForEach-Object {
        W "  [$($_.TimeCreated)] ID=$($_.Id)"
        W "     $(($_.Message -split "`n")[0])"
    }
} else {
    W '  (該当なし / タスクスケジューラの操作ログが無効の可能性)'
}
W ''

W ('=' * 60)
W ' レポート終了'
W ('=' * 60)

# ---------------------------------------------------------------
$desktop = [Environment]::GetFolderPath('Desktop')
if (-not $desktop) { $desktop = "$env:USERPROFILE\Desktop" }
$path = Join-Path $desktop ("自動起動診断_Windows_{0}.txt" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
$out -join "`r`n" | Out-File -FilePath $path -Encoding UTF8

$out -join "`r`n" | Write-Host
Write-Host ''
Write-Host "レポートを出力しました: $path" -ForegroundColor Green
Write-Host 'この内容を貼り付けてください。' -ForegroundColor Green
