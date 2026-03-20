param(
    [Parameter(Mandatory = $false)]
    [string]$Title,

    [Parameter(Mandatory = $true)]
    [string]$BodyFile,

    [Parameter(Mandatory = $false)]
    [string]$Base = "main",

    [Parameter(Mandatory = $false)]
    [string]$Head,

    [Parameter(Mandatory = $false)]
    [string]$Repo,

    [switch]$Draft
)

$ErrorActionPreference = "Stop"

function Invoke-GhJson {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Args
    )

    $output = & gh @Args
    if ($LASTEXITCODE -ne 0) {
        throw "gh command failed: gh $($Args -join ' ')"
    }

    if ([string]::IsNullOrWhiteSpace($output)) {
        return $null
    }

    return $output | ConvertFrom-Json
}

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    throw "git is required but was not found in PATH."
}

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    # gh CLI not available — fall back to browser-based PR creation.
    Write-Warning "GitHub CLI (gh) not found. Install from https://cli.github.com to automate PR creation."
    Write-Warning "Falling back to browser-based PR creation..."

    $currentHead = (& git branch --show-current).Trim()
    $usedHead = if (-not [string]::IsNullOrWhiteSpace($Head)) { $Head } else { $currentHead }

    $remoteUrl = (& git remote get-url origin 2>$null).Trim()
    # Normalize SSH → HTTPS and strip .git suffix
    $remoteUrl = $remoteUrl -replace '^git@github\.com:', 'https://github.com/'
    $remoteUrl = $remoteUrl -replace '\.git$', ''

    $prUrl = "$remoteUrl/compare/$Base...${usedHead}?expand=1"

    if (Test-Path -LiteralPath $BodyFile -PathType Leaf) {
        $bodyText = Get-Content -Raw $BodyFile
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.Clipboard]::SetText($bodyText)
        Write-Host "PR body copied to clipboard - paste it into the GitHub editor."
    }

    Write-Host "Opening: $prUrl"
    Start-Process $prUrl
    exit 0
}

$repoRoot = (& git rev-parse --show-toplevel).Trim()
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($repoRoot)) {
    throw "Unable to determine repository root. Run this script inside a git repository."
}

if ([string]::IsNullOrWhiteSpace($Head)) {
    $Head = (& git branch --show-current).Trim()
}

if ([string]::IsNullOrWhiteSpace($Head)) {
    throw "Unable to determine current branch. Provide -Head explicitly."
}

$resolvedBodyFile = if ([System.IO.Path]::IsPathRooted($BodyFile)) {
    $BodyFile
} else {
    Join-Path $repoRoot $BodyFile
}

if (-not (Test-Path -LiteralPath $resolvedBodyFile -PathType Leaf)) {
    throw "Body file not found: $resolvedBodyFile"
}

$listArgs = @("pr", "list", "--state", "open", "--head", $Head, "--json", "number,title,url")
if (-not [string]::IsNullOrWhiteSpace($Repo)) {
    $listArgs += @("--repo", $Repo)
}

$existingPrs = Invoke-GhJson -Args $listArgs
$existingPr = $null
if ($existingPrs -is [array] -and $existingPrs.Count -gt 0) {
    $existingPr = $existingPrs[0]
} elseif ($existingPrs -and $existingPrs.number) {
    $existingPr = $existingPrs
}

if ($existingPr) {
    $editArgs = @("pr", "edit", "$($existingPr.number)", "--body-file", $resolvedBodyFile)
    if (-not [string]::IsNullOrWhiteSpace($Title)) {
        $editArgs += @("--title", $Title)
    }
    if (-not [string]::IsNullOrWhiteSpace($Repo)) {
        $editArgs += @("--repo", $Repo)
    }

    & gh @editArgs
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to edit PR #$($existingPr.number)."
    }

    Write-Host "Updated PR #$($existingPr.number): $($existingPr.url)"
    exit 0
}

if ([string]::IsNullOrWhiteSpace($Title)) {
    throw "-Title is required when no open PR exists for branch '$Head'."
}

$createArgs = @("pr", "create", "--title", $Title, "--body-file", $resolvedBodyFile, "--base", $Base, "--head", $Head)
if ($Draft.IsPresent) {
    $createArgs += "--draft"
}
if (-not [string]::IsNullOrWhiteSpace($Repo)) {
    $createArgs += @("--repo", $Repo)
}

$createOutput = & gh @createArgs
if ($LASTEXITCODE -ne 0) {
    throw "Failed to create PR for head '$Head' into base '$Base'."
}

$viewArgs = @("pr", "list", "--state", "open", "--head", $Head, "--json", "number,url")
if (-not [string]::IsNullOrWhiteSpace($Repo)) {
    $viewArgs += @("--repo", $Repo)
}

$createdPrs = Invoke-GhJson -Args $viewArgs
$createdPr = $null
if ($createdPrs -is [array] -and $createdPrs.Count -gt 0) {
    $createdPr = $createdPrs[0]
} elseif ($createdPrs -and $createdPrs.number) {
    $createdPr = $createdPrs
}

if (-not $createdPr) {
    throw "PR was created, but lookup failed. Output:`n$createOutput"
}

$finalEditArgs = @("pr", "edit", "$($createdPr.number)", "--body-file", $resolvedBodyFile)
if (-not [string]::IsNullOrWhiteSpace($Repo)) {
    $finalEditArgs += @("--repo", $Repo)
}

& gh @finalEditArgs
if ($LASTEXITCODE -ne 0) {
    throw "PR created (#$($createdPr.number)) but failed to re-apply body."
}

Write-Host "Created PR #$($createdPr.number): $($createdPr.url)"