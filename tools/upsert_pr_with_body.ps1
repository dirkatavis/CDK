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
    Write-Warning "GitHub CLI (gh) not found. Attempting REST API via stored git credentials..."

    # --- Resolve head branch ---
    $usedHead = if (-not [string]::IsNullOrWhiteSpace($Head)) { $Head } else {
        (& git branch --show-current).Trim()
    }

    # --- Resolve owner/repo from remote ---
    $remoteUrl = (& git remote get-url origin 2>$null).Trim()
    $remoteUrl = $remoteUrl -replace '^git@github\.com:', 'https://github.com/'
    $remoteUrl = $remoteUrl -replace '\.git$', ''
    if ($remoteUrl -notmatch 'github\.com/([^/]+)/([^/]+)$') {
        throw "Cannot parse owner/repo from remote URL: $remoteUrl"
    }
    $ghOwner = $Matches[1]
    $ghRepo  = $Matches[2]

    # --- Resolve body file ---
    $resolvedBodyFile = if ([System.IO.Path]::IsPathRooted($BodyFile)) { $BodyFile } else {
        Join-Path (& git rev-parse --show-toplevel).Trim() $BodyFile
    }
    if (-not (Test-Path -LiteralPath $resolvedBodyFile -PathType Leaf)) {
        throw "Body file not found: $resolvedBodyFile"
    }
    $bodyText = (Get-Content -Raw $resolvedBodyFile) -replace "`r`n", "`n" -replace "`r", "`n"

    # --- Extract token from git credential store (same creds used for push) ---
    $token = $null
    try {
        $credLines = "protocol=https`nhost=github.com`n" | & git credential fill 2>$null
        foreach ($line in ($credLines -split "`r?`n")) {
            if ($line -match '^password=(.+)') { $token = $Matches[1].Trim(); break }
        }
    } catch { }

    if (-not $token) {
        # Last resort: browser fallback with body on clipboard
        Write-Warning "No stored GitHub credentials found. Store a PAT via Windows Credential Manager or 'git credential approve'."
        Write-Warning "Falling back to browser..."
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.Clipboard]::SetText($bodyText)
        Write-Host "PR body copied to clipboard - paste it into the GitHub editor."
        Start-Process "$remoteUrl/compare/$Base...${usedHead}?expand=1"
        exit 0
    }

    $apiBase = "https://api.github.com/repos/$ghOwner/$ghRepo"
    $headers = @{
        Authorization          = "token $token"
        Accept                 = "application/vnd.github+json"
        "X-GitHub-Api-Version" = "2022-11-28"
    }

    # --- Check for existing open PR on this head ---
    $existingPr = $null
    try {
        $prs = Invoke-RestMethod -Uri "$apiBase/pulls?state=open&head=${ghOwner}:${usedHead}" -Headers $headers
        if ($prs.Count -gt 0) { $existingPr = $prs[0] }
    } catch { }

    if ($existingPr) {
        $patch = @{ body = $bodyText }
        if (-not [string]::IsNullOrWhiteSpace($Title)) { $patch.title = $Title }
        try {
            $patchJson = [System.Text.Encoding]::UTF8.GetBytes(($patch | ConvertTo-Json -Depth 5 -Compress))
            Invoke-RestMethod -Uri "$apiBase/pulls/$($existingPr.number)" -Method Patch `
                -Headers $headers -Body $patchJson `
                -ContentType "application/json" | Out-Null
            Write-Host "Updated PR #$($existingPr.number): $($existingPr.html_url)"
        } catch {
            throw "Failed to update PR #$($existingPr.number) ($($existingPr.html_url)): $_"
        }
        exit 0
    }

    if ([string]::IsNullOrWhiteSpace($Title)) {
        throw "-Title is required when no open PR exists for branch '$usedHead'."
    }

    $payload = @{
        title = $Title
        body  = $bodyText
        head  = $usedHead
        base  = $Base
        draft = [bool]$Draft.IsPresent
    }
    try {
        $payloadJson = [System.Text.Encoding]::UTF8.GetBytes(($payload | ConvertTo-Json -Depth 5 -Compress))
        $created = Invoke-RestMethod -Uri "$apiBase/pulls" -Method Post `
            -Headers $headers -Body $payloadJson `
            -ContentType "application/json"
        Write-Host "Created PR #$($created.number): $($created.html_url)"
    } catch {
        throw "Failed to create PR for branch '$usedHead': $_"
    }
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