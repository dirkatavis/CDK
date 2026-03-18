$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Push-Location $repoRoot

try {
    $iniFiles = git ls-files "*.ini"

    if (-not $iniFiles) {
        Write-Host "No tracked .ini files found."
        exit 0
    }

    foreach ($file in $iniFiles) {
        git update-index --assume-unchanged -- $file | Out-Null
    }

    Write-Host "Marked assume-unchanged for tracked .ini files:"
    foreach ($file in $iniFiles) {
        Write-Host "  $file"
    }

    Write-Host ""
    Write-Host "Verification (prefix 'h' means assume-unchanged):"
    git ls-files -v "*.ini"
}
finally {
    Pop-Location
}