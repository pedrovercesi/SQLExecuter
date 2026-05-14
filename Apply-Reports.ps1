#Requires -Version 5.1
<#
    Apply-Reports.ps1
    Interactive runner that applies the SQL scripts in this folder against the target database.
    Iterates subfolders in numeric order (0, 1, 2, ...) and, inside each folder, runs the .sql
    files in numeric order via sqlcmd. Supports resume through .deploy-state.json and writes
    a log under logs/.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath
)

$ErrorActionPreference = 'Stop'

# Resolve script root robustly: $PSScriptRoot can be empty if dot-sourced,
# pasted in REPL, or invoked via certain wrappers. Fallback chain:
#   1) $PSScriptRoot
#   2) $MyInvocation.MyCommand.Path's directory
#   3) current working directory
$Root = $PSScriptRoot
if (-not $Root -and $MyInvocation.MyCommand.Path) {
    $Root = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if (-not $Root) { $Root = (Get-Location).Path }

if (-not $ConfigPath) { $ConfigPath = Join-Path $Root 'deploy-config.json' }
$StateFile = Join-Path $Root '.deploy-state.json'
$LogDir    = Join-Path $Root 'logs'

# ---------------------------------------------------------------------------
# UI helpers
# ---------------------------------------------------------------------------
function Write-Color {
    param([string]$Text, [ConsoleColor]$Color = 'White', [switch]$NoNewline)
    $prev = $Host.UI.RawUI.ForegroundColor
    try { $Host.UI.RawUI.ForegroundColor = $Color; if ($NoNewline) { Write-Host $Text -NoNewline } else { Write-Host $Text } }
    finally { $Host.UI.RawUI.ForegroundColor = $prev }
}
function Write-Header {
    param([string]$Text)
    Write-Host ''
    Write-Color ('=' * 72) Cyan
    Write-Color $Text Cyan
    Write-Color ('=' * 72) Cyan
}
function Read-Choice {
    param([string]$Prompt, [string[]]$Choices)
    while ($true) {
        Write-Host ''
        Write-Color "$Prompt " Yellow -NoNewline
        Write-Color ("[" + ($Choices -join '/') + "] ") Gray -NoNewline
        $k = Read-Host
        if ($null -eq $k) { $k = '' }
        $k = ([string]$k).Trim().ToUpper()
        if ($Choices -contains $k) { return $k }
        Write-Color "  Invalid answer. Choose: $($Choices -join ', ')" Red
    }
}

# ---------------------------------------------------------------------------
# Config + sqlcmd resolution
# ---------------------------------------------------------------------------
function Read-Config {
    if (-not (Test-Path $ConfigPath)) { throw "Config file not found: $ConfigPath" }
    $cfg = Get-Content $ConfigPath -Raw | ConvertFrom-Json
    foreach ($f in 'server','database','authentication') {
        if (-not $cfg.$f) { throw "Missing '$f' field in deploy-config.json" }
    }
    if ($cfg.authentication -notin 'Windows','SQL') {
        throw "authentication must be 'Windows' or 'SQL', got: $($cfg.authentication)"
    }
    if ($cfg.server -eq 'SERVER\INSTANCE') {
        throw "deploy-config.json still has the 'SERVER\INSTANCE' placeholder. Edit it before running."
    }
    return $cfg
}

function Resolve-Sqlcmd {
    param([string]$Hint)
    if ($Hint -and (Get-Command $Hint -ErrorAction SilentlyContinue)) { return (Get-Command $Hint).Source }
    if (Get-Command sqlcmd -ErrorAction SilentlyContinue) { return (Get-Command sqlcmd).Source }
    $candidates = @(
        "${env:ProgramFiles}\Microsoft SQL Server\Client SDK\ODBC\*\Tools\Binn\SQLCMD.EXE",
        "${env:ProgramFiles(x86)}\Microsoft SQL Server\Client SDK\ODBC\*\Tools\Binn\SQLCMD.EXE",
        "${env:ProgramFiles}\Microsoft SQL Server\*\Tools\Binn\SQLCMD.EXE"
    )
    foreach ($c in $candidates) {
        $found = Get-ChildItem $c -ErrorAction SilentlyContinue | Sort-Object FullName -Descending | Select-Object -First 1
        if ($found) { return $found.FullName }
    }
    throw "sqlcmd not found. Install the SQL Server Command Line Utilities or set sqlcmdPath in deploy-config.json."
}

function Build-SqlcmdArgs {
    param($Cfg, [string]$ScriptPath, [string]$Password)
    $a = @('-S', $Cfg.server, '-d', $Cfg.database, '-i', $ScriptPath, '-b', '-V', '16')
    if ($Cfg.authentication -eq 'Windows') {
        $a += '-E'
    } else {
        $a += @('-U', $Cfg.username, '-P', $Password)
    }
    return ,$a
}

function Test-Connection-Sql {
    param($Cfg, [string]$SqlcmdExe, [string]$Password)
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("conn_test_" + [Guid]::NewGuid().ToString('N') + ".sql")
    Set-Content -Path $tmp -Value "SELECT @@VERSION;" -Encoding utf8
    try {
        $args_ = Build-SqlcmdArgs -Cfg $Cfg -ScriptPath $tmp -Password $Password
        $out = & $SqlcmdExe @args_ 2>&1
        if ($LASTEXITCODE -ne 0) { throw "Connection failed ($LASTEXITCODE): $out" }
        return ($out -join "`n")
    } finally {
        Remove-Item $tmp -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# State + folders
# ---------------------------------------------------------------------------
function Read-State {
    if (Test-Path $StateFile) {
        try {
            $raw = Get-Content $StateFile -Raw | ConvertFrom-Json
            $h = @{}
            foreach ($p in $raw.PSObject.Properties) { $h[$p.Name] = $p.Value }
            return $h
        } catch { return @{} }
    }
    return @{}
}
function Save-State {
    param([hashtable]$State)
    $State | ConvertTo-Json -Depth 6 | Set-Content -Path $StateFile -Encoding utf8
}

function Get-NumericPrefix {
    param([string]$Name)
    if ($Name -match '^(\d+)') { return [int]$Matches[1] }
    return [int]::MaxValue
}

function Get-OrderedFolders {
    Get-ChildItem -Path $Root -Directory |
        Where-Object { $_.Name -match '^\d+\.' } |
        Sort-Object @{Expression = { Get-NumericPrefix $_.Name }}
}

function Test-IsSkipped {
    param([string]$FolderName, $SkipList)
    if (-not $SkipList) { return $false }
    foreach ($s in $SkipList) {
        if ($FolderName -ieq $s) { return $true }
    }
    return $false
}

function Get-OrderedScripts {
    param([System.IO.DirectoryInfo]$Folder)
    Get-ChildItem -Path $Folder.FullName -Filter '*.sql' -File |
        Sort-Object @{Expression = { Get-NumericPrefix $_.Name }}, Name
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
$cfg = Read-Config
$sqlcmd = Resolve-Sqlcmd -Hint $cfg.sqlcmdPath
Write-Header "Apply-Reports"
Write-Host "Root folder : $Root"
Write-Host "sqlcmd      : $sqlcmd"
Write-Host "Server      : $($cfg.server)"
Write-Host "Database    : $($cfg.database)"
Write-Host "Auth        : $($cfg.authentication)"

$password = $null
if ($cfg.authentication -eq 'SQL') {
    if (-not $cfg.username) { throw "authentication=SQL but username is empty in deploy-config.json" }
    Write-Host "User        : $($cfg.username)"
    $sec = Read-Host "Password for $($cfg.username)" -AsSecureString
    $password = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
    )
}

Write-Host ''
Write-Color "Testing connection..." Gray
try {
    $ver = Test-Connection-Sql -Cfg $cfg -SqlcmdExe $sqlcmd -Password $password
    Write-Color "OK" Green
    Write-Host ($ver.Split("`n") | Select-Object -First 2 | Out-String).TrimEnd()
} catch {
    Write-Color "FAILED: $_" Red
    exit 1
}

$state = Read-State
$allFolders = Get-OrderedFolders
if (-not $allFolders) { Write-Color "No numbered subfolder found." Red; exit 1 }

$skipList = @()
if ($cfg.PSObject.Properties.Name -contains 'skipFolders' -and $cfg.skipFolders) {
    $skipList = @($cfg.skipFolders)
}
$folders = @($allFolders | Where-Object { -not (Test-IsSkipped -FolderName $_.Name -SkipList $skipList) })
$skipped = @($allFolders | Where-Object { Test-IsSkipped -FolderName $_.Name -SkipList $skipList })

$okCount = ($state.GetEnumerator() | Where-Object { $_.Value.status -eq 'OK' }).Count
Write-Host ''
Write-Color "Folders detected: $($allFolders.Count) (to run: $($folders.Count), ignored: $($skipped.Count)). Previously applied successfully: $okCount." Gray
if ($skipped.Count -gt 0) {
    Write-Color "Ignored via config (skipFolders):" Yellow
    foreach ($f in $skipped) { Write-Host "  - $($f.Name)" }
}

# Prepare log
if (-not (Test-Path $LogDir)) { New-Item -Type Directory -Path $LogDir | Out-Null }
$logFile = Join-Path $LogDir ("deploy-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
function Log { param([string]$Line) Add-Content -Path $logFile -Value ("[{0:yyyy-MM-dd HH:mm:ss}] {1}" -f (Get-Date), $Line) -Encoding utf8 }
Log "=== Run started by $env:USERNAME on $env:COMPUTERNAME against $($cfg.server)/$($cfg.database) ==="

$summary = @{ OK = 0; FAILED = 0; PARTIAL = 0; SKIPPED = 0 }
$idx = 0
$abortAll = $false
foreach ($folder in $folders) {
    if ($abortAll) { break }
    $idx++
    $prev = $state[$folder.Name]
    $scripts = Get-OrderedScripts -Folder $folder
    Write-Header ("[{0}/{1}] {2}" -f $idx, $folders.Count, $folder.Name)
    Write-Host ("Scripts: {0}" -f $scripts.Count)
    foreach ($s in $scripts) { Write-Host "  - $($s.Name)" }
    if ($prev) {
        $col = switch ($prev.status) { 'OK' {'Green'} 'PARTIAL' {'Yellow'} default {'Red'} }
        Write-Color ("Previous state: {0} at {1}" -f $prev.status, $prev.timestamp) $col
    }

    $choice = Read-Choice "Apply this folder?" @('Y','N','Q')
    if ($choice -eq 'Q') { Write-Color "Cancelled by user." Yellow; break }
    if ($choice -eq 'N') {
        Write-Color "Skipped." Gray
        Log "SKIP folder $($folder.Name) (user)"
        $summary.SKIPPED++
        continue
    }

    $folderResult = @{ status = 'OK'; timestamp = (Get-Date).ToString('s'); scriptsRun = @() }
    $abortFolder = $false
    foreach ($s in $scripts) {
        $tries = 0
        while ($true) {
            $tries++
            Write-Host ''
            Write-Color "  -> $($s.Name)" Cyan
            Log "RUN $($folder.Name)/$($s.Name) (attempt $tries)"
            $args_ = Build-SqlcmdArgs -Cfg $cfg -ScriptPath $s.FullName -Password $password
            $output = & $sqlcmd @args_ 2>&1
            $code = $LASTEXITCODE
            $output | ForEach-Object { Log "    $_" }
            if ($code -eq 0) {
                Write-Color "     OK" Green
                $folderResult.scriptsRun += @{ script = $s.Name; status = 'OK' }
                break
            } else {
                Write-Color "     FAILED (exit $code)" Red
                $tail = ($output | Select-Object -Last 20) -join "`n"
                Write-Host $tail
                $err = Read-Choice "Action?" @('R','S','A')
                if ($err -eq 'R') { continue }
                if ($err -eq 'S') {
                    Write-Color "     Script skipped." Yellow
                    Log "SKIP script $($s.Name) (user, after error)"
                    $folderResult.scriptsRun += @{ script = $s.Name; status = 'SKIPPED' }
                    if ($folderResult.status -eq 'OK') { $folderResult.status = 'PARTIAL' }
                    break
                }
                if ($err -eq 'A') {
                    Write-Color "     Folder aborted." Red
                    Log "ABORT folder $($folder.Name) at $($s.Name) (user)"
                    $folderResult.scriptsRun += @{ script = $s.Name; status = 'FAILED' }
                    $folderResult.status = 'FAILED'
                    $abortFolder = $true
                    break
                }
            }
        }
        if ($abortFolder) { break }
    }

    $state[$folder.Name] = $folderResult
    Save-State -State $state
    switch ($folderResult.status) {
        'OK'      { Write-Color "Folder completed (OK)." Green;             $summary.OK++ }
        'PARTIAL' { Write-Color "Folder completed with skips (PARTIAL)." Yellow; $summary.PARTIAL++ }
        'FAILED'  { Write-Color "Folder aborted (FAILED)." Red;             $summary.FAILED++ }
    }

    if ($abortFolder) {
        $cont = Read-Choice "Continue to the next folder?" @('Y','Q')
        if ($cont -eq 'Q') { $abortAll = $true }
    }
}

Write-Header "Summary"
Write-Color ("OK      : {0}" -f $summary.OK) Green
Write-Color ("PARTIAL : {0}" -f $summary.PARTIAL) Yellow
Write-Color ("FAILED  : {0}" -f $summary.FAILED) Red
Write-Color ("SKIPPED : {0}" -f $summary.SKIPPED) Gray
Write-Host ''
Write-Host "Log   : $logFile"
Write-Host "State : $StateFile"
Log "=== Run finished. OK=$($summary.OK) PARTIAL=$($summary.PARTIAL) FAILED=$($summary.FAILED) SKIPPED=$($summary.SKIPPED) ==="
