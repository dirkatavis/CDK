<#
.SYNOPSIS
  Parse Initialize_RO.log and extract RO numbers and MVA numbers.

.DESCRIPTION
  Reads a log path from the `-LogPath` parameter or from the `Initialize_RO->Log` key
  in `config/config.ini` (located at repository root). The resolved log file MUST
  exist — there are NO fallback paths. The script parses lines matching the pattern
  `MVA: <digits> - RO: <digits>` and writes two files next to the log:
    - parse_initialize_log_ro.txt  (one RO per line)
    - parse_initialize_log_mva.txt (one MVA per line)

.NOTES
  - Targets PowerShell 5.1+ (uses core features only).
  - Overwrites output files each run and preserves original log order.
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$false)]
    [string]$LogPath
)

function Fail([string]$msg, [int]$code=1) {
    Write-Error $msg
    exit $code
}

$PSScriptRootResolved = Split-Path -Parent $MyInvocation.MyCommand.Definition

if (-not $LogPath) {
    $configPath = Join-Path $PSScriptRootResolved '..\config\config.ini'
    $configPath = Resolve-Path $configPath -ErrorAction SilentlyContinue
    if (-not $configPath) { Fail "config/config.ini not found at expected location. Provide -LogPath explicitly." 2 }
    $configPath = $configPath.ProviderPath

    try {
        $iniLines = Get-Content -LiteralPath $configPath -ErrorAction Stop
    } catch {
        Fail "Failed reading config file: $configPath" 3
    }

    $section = 'Initialize_RO'
    $inSection = $false
    $logValue = $null
    foreach ($line in $iniLines) {
        $trim = $line.Trim()
        if ($trim -match '^[;#]') { continue }
        if ($trim -match '^\[(.+)\]') {
            $inSection = ($matches[1].Trim() -ieq $section)
            continue
        }
        if ($inSection -and $trim -match '^Log\s*=\s*(.+)$') {
            $logValue = $matches[1].Trim()
            break
        }
    }

    if (-not $logValue) { Fail "'Log' key not found in [$section] of config/config.ini. Provide -LogPath explicitly." 4 }

    # If the config value is relative, interpret it relative to repo root
    $repoRoot = Resolve-Path (Join-Path $PSScriptRootResolved '..')
    if (-not [System.IO.Path]::IsPathRooted($logValue)) {
        $LogPath = Join-Path $repoRoot $logValue
    } else {
        $LogPath = $logValue
    }
}

$LogPath = (Resolve-Path -LiteralPath $LogPath -ErrorAction SilentlyContinue)
if (-not $LogPath) { Fail "Resolved log path does not exist. Provide a valid -LogPath or update config/config.ini." 5 }
$LogPath = $LogPath.ProviderPath

$outFolder = Split-Path -Parent $LogPath
$roFile = Join-Path $outFolder 'parse_initialize_log_ro.txt'
$mvaFile = Join-Path $outFolder 'parse_initialize_log_mva.txt'

# Overwrite (truncate) output files
if (Test-Path $roFile) { Remove-Item -LiteralPath $roFile -Force }
New-Item -LiteralPath $roFile -ItemType File -Force | Out-Null
if (Test-Path $mvaFile) { Remove-Item -LiteralPath $mvaFile -Force }
New-Item -LiteralPath $mvaFile -ItemType File -Force | Out-Null

Write-Output "Parsing log: $LogPath"

$pattern = [regex] 'MVA:\s*(\d{6,9})\s*-\s*RO:\s*(\d+)'
$matchCount = 0
$linesProcessed = 0

try {
    Get-Content -LiteralPath $LogPath -Encoding UTF8 | ForEach-Object {
        $linesProcessed++
        $line = $_
        $m = $pattern.Match($line)
        if ($m.Success) {
            $mva = $m.Groups[1].Value
            $ro  = $m.Groups[2].Value
            Add-Content -LiteralPath $roFile -Value $ro
            Add-Content -LiteralPath $mvaFile -Value $mva
            $matchCount++
        }
    }
} catch {
    Fail "Error reading or parsing log file: $_" 6
}

if ($matchCount -eq 0) {
    Write-Warning "No matching 'MVA: <digits> - RO: <digits>' lines were found in the log."
}

Write-Output "Lines scanned: $linesProcessed"
Write-Output "Entries written: ROs=$matchCount, MVAs=$matchCount"
Write-Output "RO output: $roFile"
Write-Output "MVA output: $mvaFile"
exit 0
