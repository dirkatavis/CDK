# CDK Dependency Validation Guide

This guide explains how to validate your CDK environment before running automation scripts.

## Quick Start

Before running any CDK script for the first time, run:

```cmd
cd C:\Temp_alt\CDK
cscript.exe tools\validate_dependencies.vbs
```

This performs a complete pre-flight check of your environment.

## What Gets Validated

### Check 1: CDK_BASE Environment Variable
**What it checks:** Is the `CDK_BASE` user environment variable set?

**Why it matters:** All scripts depend on this variable to locate the CDK repository root.

**If it fails:**
- Run: `cscript.exe tools\setup_cdk_base.vbs`
- This sets `CDK_BASE` to point to your CDK repo root
- You may need to restart your terminal or VS Code after setup

### Check 2: Repository Root Marker (.cdkroot)
**What it checks:** Does a `.cdkroot` file exist at the repo root?

**Why it matters:** This marker file confirms the repository structure is intact.

**If it fails:** This is a warning only. The file should exist but its absence won't prevent scripts from running.

### Check 3: PathHelper.vbs
**What it checks:** Does `common\PathHelper.vbs` exist?

**Why it matters:** All scripts require PathHelper to resolve paths from the central configuration.

**If it fails:** Ensure the `common\PathHelper.vbs` file exists. This file should be:
- Location: `{CDK_REPO_ROOT}\common\PathHelper.vbs`
- It's essential for all path resolution

### Check 4: config.ini
**What it checks:** Does `config.ini` exist at the repository root?

**Why it matters:** config.ini is the central configuration file for all scripts. It defines paths and parameters.

**If it fails:** Ensure `config.ini` exists at:
- Location: `{CDK_REPO_ROOT}\config.ini`
- This file should not be moved or renamed

### Check 5: Configured File Paths
**What it checks:** Do all paths referenced in config.ini actually exist?

**Why it matters:** Scripts read paths from config.ini. If referenced files are missing, scripts will fail.

**If it fails:** Verify that:
- Input CSV files exist in the expected locations
- Output directories exist
- Referenced scripts exist

## Validation Results

### ✓ ALL CHECKS PASSED
Everything is ready! You can run CDK scripts:
```cmd
cscript.exe PostFinalCharges\PostFinalCharges.vbs
cscript.exe Close_ROs\Close_ROs_Pt1.vbs
' ... other scripts
```

### ⊘ WARNINGS DETECTED
Your environment will likely work, but review the warnings. Common warnings:
- Missing `.cdkroot` marker (non-critical)
- Log files that don't exist yet (will be created on first run)

### ✗ FAILURES DETECTED
Scripts will NOT run until you fix these issues. Follow the remediation steps shown for each failure.

## Troubleshooting

### "CDK_BASE environment variable not set"
```cmd
cd C:\Temp_alt\CDK
cscript.exe tools\setup_cdk_base.vbs
```
Then close and reopen your terminal.

### "CDK_BASE points to non-existent folder"
Edit your user environment variables:
1. Windows Key + X → System
2. Advanced system settings → Environment Variables
3. Under User variables, find `CDK_BASE`
4. Edit it to point to your CDK repository root (e.g., `C:\Temp_alt\CDK`)
5. Click OK and close your terminal, then reopen

### "PathHelper.vbs not found"
Ensure your CDK repository has the complete structure:
```
CDK\
├── common\
│   ├── PathHelper.vbs     ← Should be here
│   └── ValidateSetup.vbs
├── config.ini             ← Should be here
├── .cdkroot               ← Should be here
└── ... other folders
```

### "config.ini not found"
Ensure `config.ini` exists at your repository root. It should contain sections like:
```ini
[PostFinalCharges]
CSV=PostFinalCharges\CashoutRoList.csv
Log=PostFinalCharges\PostFinalCharges.log
```

## For Developers: Adding Validation to Your Scripts

To add dependency validation to a new script, include this at startup:

### Option 1: Strict Validation (Recommended for critical scripts)
```vbscript
' At the top of your script, after Option Explicit:
' <reference path="../../common/PathHelper.vbs" />
' <reference path="../../common/ValidateSetup.vbs" />

Option Explicit

' ... your other includes ...

' Validate before doing anything else
Sub Main()
    MustHaveValidDependencies
    ' ... rest of your script ...
End Sub

Main
```

### Option 2: Soft Validation (For utility scripts)
```vbscript
Option Explicit

Sub Main()
    If Not ValidateScriptDependencies() Then
        WScript.Echo "Warning: Some dependencies may be missing."
        WScript.Echo "Run: cscript.exe tools\validate_dependencies.vbs"
        ' Continue anyway, or exit:
        ' WScript.Quit 1
    End If
    ' ... rest of your script ...
End Sub

Main
```

### Option 3: Silent Validation (For standalone utilities)
```vbscript
' Just check if repo root is available, don't enforce
Dim repoRoot
repoRoot = GetRepoRootSafe()
If repoRoot = "" Then
    WScript.Echo "ERROR: CDK_BASE not set. Run: tools\setup_cdk_base.vbs"
    WScript.Quit 1
End If
```

## For System Administrators: Distribution Checklist

When distributing CDK to other users:

1. ✓ Ensure `.cdkroot` marker file is in the repo root
2. ✓ Ensure `config.ini` is present with all required sections
3. ✓ Ensure `common\PathHelper.vbs` and `common\ValidateSetup.vbs` exist
4. ✓ Provide these instructions to end users
5. ✓ Have users run: `cscript.exe tools\setup_cdk_base.vbs` during initial setup
6. ✓ Have users run: `cscript.exe tools\validate_dependencies.vbs` to verify installation

## Support

If you encounter validation errors not covered here:

1. Run: `cscript.exe tools\validate_dependencies.vbs` (note verbosity)
2. Check the config.ini file for typos or missing sections
3. Verify all directory paths match your actual folder structure
4. Ensure the CDK folder structure matches the expected layout

For help with specific scripts, see:
- `PostFinalCharges\README.md` (if present)
- `docs\PATH_CONFIGURATION.md`
